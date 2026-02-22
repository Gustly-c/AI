from __future__ import annotations

from collections import defaultdict
from dataclasses import dataclass

from ortools.sat.python import cp_model

from scheduler_desktop.models import (
    DAYS,
    SLOTS,
    AppState,
    Assignment,
    Discipline,
    Room,
    ScheduleEntry,
    Stream,
    StudentGroup,
    Teacher,
    slot_key,
)


class PlanningError(RuntimeError):
    pass


@dataclass(frozen=True)
class SessionRef:
    assignment_id: str
    session_index: int

    @property
    def key(self) -> str:
        return f"{self.assignment_id}:{self.session_index}"


def weeks_overlap(a: Assignment, b: Assignment) -> bool:
    return not (a.end_week < b.start_week or b.end_week < a.start_week)


def validate_state(state: AppState) -> None:
    teacher_ids = {t.id for t in state.teachers}
    room_ids = {r.id for r in state.rooms}
    group_ids = {g.id for g in state.groups}
    stream_ids = {s.id for s in state.streams}
    discipline_ids = {d.id for d in state.disciplines}

    for stream in state.streams:
        for group_id in stream.group_ids:
            if group_id not in group_ids:
                raise PlanningError(f"Поток {stream.id}: неизвестная группа {group_id}")
        if stream.preferred_room_id and stream.preferred_room_id not in room_ids:
            raise PlanningError(f"Поток {stream.id}: неизвестная аудитория {stream.preferred_room_id}")

    for assignment in state.assignments:
        if assignment.teacher_id not in teacher_ids:
            raise PlanningError(f"Назначение {assignment.id}: неизвестный преподаватель {assignment.teacher_id}")
        if assignment.discipline_id not in discipline_ids:
            raise PlanningError(f"Назначение {assignment.id}: неизвестная дисциплина {assignment.discipline_id}")
        if assignment.stream_id and assignment.stream_id not in stream_ids:
            raise PlanningError(f"Назначение {assignment.id}: неизвестный поток {assignment.stream_id}")
        for group_id in assignment.group_ids:
            if group_id not in group_ids:
                raise PlanningError(f"Назначение {assignment.id}: неизвестная группа {group_id}")
        if assignment.room_id and assignment.room_id not in room_ids:
            raise PlanningError(f"Назначение {assignment.id}: неизвестная аудитория {assignment.room_id}")
        if assignment.lock_room_id and assignment.lock_room_id not in room_ids:
            raise PlanningError(f"Назначение {assignment.id}: lock_room_id {assignment.lock_room_id} не найдена")
        if assignment.sessions_per_week <= 0:
            raise PlanningError(f"Назначение {assignment.id}: sessions_per_week должно быть > 0")
        if assignment.start_week > assignment.end_week:
            raise PlanningError(f"Назначение {assignment.id}: start_week больше end_week")


class ScheduleGenerator:
    def __init__(self, state: AppState) -> None:
        self.state = state
        self.teachers = {t.id: t for t in state.teachers}
        self.rooms = {r.id: r for r in state.rooms}
        self.groups = {g.id: g for g in state.groups}
        self.streams = {s.id: s for s in state.streams}
        self.disciplines = {d.id: d for d in state.disciplines}
        self.assignments = {a.id: a for a in state.assignments}
        self.max_week = max((a.end_week for a in state.assignments), default=1)

    def generate(self, time_limit_sec: int = 12) -> list[ScheduleEntry]:
        validate_state(self.state)
        model = cp_model.CpModel()

        sessions: list[SessionRef] = []
        for assignment in self.state.assignments:
            for idx in range(assignment.sessions_per_week):
                sessions.append(SessionRef(assignment_id=assignment.id, session_index=idx))

        if not sessions:
            return []

        candidate_timeslots: dict[str, list[tuple[int, int]]] = {}
        candidate_rooms: dict[str, list[str]] = {}

        for session in sessions:
            assignment = self.assignments[session.assignment_id]
            discipline = self.disciplines[assignment.discipline_id]
            teacher = self.teachers[assignment.teacher_id]
            assignment_group_ids = self._assignment_group_ids(assignment)

            self._validate_contract_window(assignment, teacher)
            for group_id in assignment_group_ids:
                self._validate_group_window(assignment, self.groups[group_id])

            slots = self._build_timeslot_candidates(assignment, teacher)
            if not slots:
                raise PlanningError(
                    f"Назначение {assignment.id}: нет доступных слотов (рабочие дни/блокировки/смена группы)"
                )

            rooms = self._build_room_candidates(assignment, discipline)
            if not rooms:
                raise PlanningError(
                    f"Назначение {assignment.id}: нет подходящих аудиторий "
                    f"(вместимость/тип/фиксированная аудитория)"
                )

            candidate_timeslots[session.key] = slots
            candidate_rooms[session.key] = rooms

        y_vars: dict[tuple[str, int, int], cp_model.IntVar] = {}
        x_vars: dict[tuple[str, int, int, str], cp_model.IntVar] = {}

        for session in sessions:
            skey = session.key
            slot_choice_vars: list[cp_model.IntVar] = []
            for day, slot in candidate_timeslots[skey]:
                y = model.NewBoolVar(f"y_{skey}_{day}_{slot}")
                y_vars[(skey, day, slot)] = y
                slot_choice_vars.append(y)

                room_vars: list[cp_model.IntVar] = []
                for room_id in candidate_rooms[skey]:
                    x = model.NewBoolVar(f"x_{skey}_{day}_{slot}_{room_id}")
                    x_vars[(skey, day, slot, room_id)] = x
                    room_vars.append(x)

                model.Add(sum(room_vars) == y)

            model.Add(sum(slot_choice_vars) == 1)

        by_week_and_day_teacher: dict[tuple[int, int, str], list[cp_model.IntVar]] = defaultdict(list)
        by_week_teacher: dict[tuple[int, str], list[cp_model.IntVar]] = defaultdict(list)
        by_week_slot_teacher: dict[tuple[int, int, int, str], list[cp_model.IntVar]] = defaultdict(list)
        by_week_slot_group: dict[tuple[int, int, int, str], list[cp_model.IntVar]] = defaultdict(list)
        by_week_slot_room: dict[tuple[int, int, int, str], list[cp_model.IntVar]] = defaultdict(list)

        for session in sessions:
            assignment = self.assignments[session.assignment_id]
            teacher_id = assignment.teacher_id
            skey = session.key
            assignment_group_ids = self._assignment_group_ids(assignment)

            for week in range(assignment.start_week, assignment.end_week + 1):
                for day, slot in candidate_timeslots[skey]:
                    y = y_vars[(skey, day, slot)]
                    by_week_and_day_teacher[(week, day, teacher_id)].append(y)
                    by_week_teacher[(week, teacher_id)].append(y)
                    by_week_slot_teacher[(week, day, slot, teacher_id)].append(y)

                    for group_id in assignment_group_ids:
                        by_week_slot_group[(week, day, slot, group_id)].append(y)

                    for room_id in candidate_rooms[skey]:
                        x = x_vars[(skey, day, slot, room_id)]
                        by_week_slot_room[(week, day, slot, room_id)].append(x)

        for vars_ in by_week_slot_teacher.values():
            model.Add(sum(vars_) <= 1)
        for vars_ in by_week_slot_group.values():
            model.Add(sum(vars_) <= 1)
        for vars_ in by_week_slot_room.values():
            model.Add(sum(vars_) <= 1)

        for (week, day, teacher_id), vars_ in by_week_and_day_teacher.items():
            teacher = self.teachers[teacher_id]
            model.Add(sum(vars_) <= teacher.max_classes_per_day)
        for (week, teacher_id), vars_ in by_week_teacher.items():
            teacher = self.teachers[teacher_id]
            model.Add(sum(vars_) <= teacher.max_classes_per_week)

        penalties: list[cp_model.LinearExpr] = []
        for session in sessions:
            assignment = self.assignments[session.assignment_id]
            discipline = self.disciplines[assignment.discipline_id]
            teacher = self.teachers[assignment.teacher_id]
            assignment_group_ids = self._assignment_group_ids(assignment)
            skey = session.key

            for day, slot in candidate_timeslots[skey]:
                for room_id in candidate_rooms[skey]:
                    x = x_vars[(skey, day, slot, room_id)]
                    room = self.rooms[room_id]

                    penalty = 0
                    if teacher.default_room_id and room_id != teacher.default_room_id:
                        penalty += 2
                    if assignment.room_id and room_id != assignment.room_id:
                        penalty += 5
                    if discipline.fixed_room_id and room_id != discipline.fixed_room_id:
                        penalty += 7
                    if slot >= 6:
                        penalty += 1

                    group_shift_penalty = 0
                    for group_id in assignment_group_ids:
                        group = self.groups[group_id]
                        if slot > group.shift_end_slot:
                            group_shift_penalty += 4
                        if slot < group.shift_start_slot:
                            group_shift_penalty += 4
                    penalty += group_shift_penalty

                    if penalty:
                        penalties.append(penalty * x)

        if penalties:
            model.Minimize(sum(penalties))

        solver = cp_model.CpSolver()
        solver.parameters.max_time_in_seconds = time_limit_sec
        solver.parameters.num_search_workers = 8
        solver.parameters.log_search_progress = False
        status = solver.Solve(model)

        if status not in (cp_model.OPTIMAL, cp_model.FEASIBLE):
            raise PlanningError(
                "Не удалось построить расписание с текущими ограничениями. "
                "Ослабьте ограничения или замените ресурсы."
            )

        result: list[ScheduleEntry] = []
        for session in sessions:
            assignment = self.assignments[session.assignment_id]
            skey = session.key
            chosen_day = None
            chosen_slot = None
            chosen_room = None
            for day, slot in candidate_timeslots[skey]:
                if solver.Value(y_vars[(skey, day, slot)]) == 1:
                    chosen_day = day
                    chosen_slot = slot
                    for room_id in candidate_rooms[skey]:
                        x = x_vars[(skey, day, slot, room_id)]
                        if solver.Value(x) == 1:
                            chosen_room = room_id
                            break
                    break

            if chosen_day is None or chosen_slot is None or chosen_room is None:
                raise PlanningError(f"Внутренняя ошибка: не выбран слот для {skey}")

            result.append(
                ScheduleEntry(
                    assignment_id=assignment.id,
                    discipline_id=assignment.discipline_id,
                    teacher_id=assignment.teacher_id,
                    group_ids=assignment_group_ids,
                    day=chosen_day,
                    slot=chosen_slot,
                    room_id=chosen_room,
                    start_week=assignment.start_week,
                    end_week=assignment.end_week,
                )
            )

        result.sort(key=lambda s: (s.day, s.slot, s.assignment_id))
        return result

    def _validate_contract_window(self, assignment: Assignment, teacher: Teacher) -> None:
        if assignment.start_week < teacher.contract_start_week or assignment.end_week > teacher.contract_end_week:
            raise PlanningError(
                f"Назначение {assignment.id}: период вне контракта преподавателя {teacher.id} "
                f"({teacher.contract_start_week}-{teacher.contract_end_week})"
            )

    def _validate_group_window(self, assignment: Assignment, group: StudentGroup) -> None:
        if assignment.start_week < group.program_start_week or assignment.end_week > group.program_end_week:
            raise PlanningError(
                f"Назначение {assignment.id}: период вне учебного плана группы {group.id} "
                f"({group.program_start_week}-{group.program_end_week})"
            )

    def _build_timeslot_candidates(
        self,
        assignment: Assignment,
        teacher: Teacher,
    ) -> list[tuple[int, int]]:
        slots: list[tuple[int, int]] = []
        locked = assignment.lock_day is not None and assignment.lock_slot is not None
        assignment_group_ids = self._assignment_group_ids(assignment)

        for day in range(len(DAYS)):
            for slot in SLOTS:
                if locked and (day != assignment.lock_day or slot != assignment.lock_slot):
                    continue
                if day not in teacher.work_days:
                    continue
                if slot_key(day, slot) in teacher.blocked_slots:
                    continue

                allowed_for_all_groups = True
                for group_id in assignment_group_ids:
                    group = self.groups[group_id]
                    if day in group.blocked_days:
                        allowed_for_all_groups = False
                        break
                    if slot < group.shift_start_slot or slot > group.shift_end_slot:
                        allowed_for_all_groups = False
                        break
                    if slot_key(day, slot) in group.blocked_slots:
                        allowed_for_all_groups = False
                        break
                if not allowed_for_all_groups:
                    continue

                slots.append((day, slot))
        return slots

    def _build_room_candidates(
        self,
        assignment: Assignment,
        discipline: Discipline,
    ) -> list[str]:
        stream_preferred_room_id = None
        if assignment.stream_id:
            stream = self.streams.get(assignment.stream_id)
            if stream:
                stream_preferred_room_id = stream.preferred_room_id

        requested_room = (
            assignment.lock_room_id
            or assignment.room_id
            or discipline.fixed_room_id
            or stream_preferred_room_id
        )
        if requested_room:
            room = self.rooms[requested_room]
            if self._room_fits_assignment(room, assignment, discipline):
                return [requested_room]
            return []

        rooms: list[str] = []
        for room in self.rooms.values():
            if self._room_fits_assignment(room, assignment, discipline):
                rooms.append(room.id)
        return rooms

    def _room_fits_assignment(self, room: Room, assignment: Assignment, discipline: Discipline) -> bool:
        total_students = sum(self.groups[group_id].size for group_id in self._assignment_group_ids(assignment))
        if room.capacity < total_students:
            return False

        if discipline.required_room_features:
            room_features = set(room.features)
            for feature in discipline.required_room_features:
                if feature not in room_features:
                    return False

        return True

    def _assignment_group_ids(self, assignment: Assignment) -> list[str]:
        resolved: list[str] = []
        if assignment.stream_id:
            stream = self.streams.get(assignment.stream_id)
            if stream:
                resolved.extend(stream.group_ids)
        resolved.extend(assignment.group_ids)

        unique_ids = sorted(set(resolved))
        if not unique_ids:
            raise PlanningError(
                f"Назначение {assignment.id}: не заданы группы (ни в group_ids, ни в stream_id)"
            )
        return unique_ids
