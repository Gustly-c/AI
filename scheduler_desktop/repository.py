from __future__ import annotations

import json
from pathlib import Path

from scheduler_desktop.excel_io import export_state_to_excel, import_state_from_excel
from scheduler_desktop.models import (
    AppState,
    Assignment,
    Discipline,
    Room,
    Stream,
    StudentGroup,
    Teacher,
)


class StateRepository:
    def __init__(self, path: Path) -> None:
        self.path = path

    def load(self) -> AppState:
        if not self.path.exists():
            return sample_state()
        if self.path.suffix.lower() == ".xlsx":
            return import_state_from_excel(self.path)
        data = json.loads(self.path.read_text(encoding="utf-8"))
        return AppState.from_dict(data)

    def save(self, state: AppState) -> None:
        self.path.parent.mkdir(parents=True, exist_ok=True)
        if self.path.suffix.lower() == ".xlsx":
            export_state_to_excel(self.path, state)
            return
        self.path.write_text(json.dumps(state.to_dict(), ensure_ascii=False, indent=2), encoding="utf-8")


JsonRepository = StateRepository


def sample_state() -> AppState:
    teachers = [
        Teacher(
            id="T1",
            name="Иванов И.И.",
            department="Кафедра ИТ",
            position="Доцент",
            degree="PhD",
            default_room_id="R101",
            blocked_slots=["2:1", "2:2"],
            max_classes_per_day=4,
        ),
        Teacher(
            id="T2",
            name="Петрова А.А.",
            department="Кафедра Математики",
            position="Старший преподаватель",
            degree="MSc",
            default_room_id="R201",
            blocked_slots=["4:5", "4:6"],
            max_classes_per_day=3,
        ),
        Teacher(
            id="T3",
            name="Садыков Б.Б.",
            department="Кафедра ИТ",
            position="Профессор",
            degree="ScD",
            default_room_id="LAB1",
            blocked_slots=["0:1", "5:7", "5:8"],
            max_classes_per_day=3,
        ),
    ]

    rooms = [
        Room(id="R101", name="Аудитория 101", capacity=50, building="A", features=["lecture"]),
        Room(id="R201", name="Аудитория 201", capacity=35, building="B", features=["lecture"]),
        Room(id="LAB1", name="Лаборатория 1", capacity=30, building="A", features=["lab", "practice"]),
        Room(id="LAB2", name="Лаборатория 2", capacity=28, building="C", features=["lab"]),
    ]

    groups = [
        StudentGroup(id="G1", name="SE-23-1", size=24, shift_start_slot=1, shift_end_slot=5, blocked_days=[2]),
        StudentGroup(id="G2", name="SE-23-2", size=23, shift_start_slot=3, shift_end_slot=8, blocked_days=[]),
        StudentGroup(id="G3", name="IS-23-1", size=27, shift_start_slot=1, shift_end_slot=6, blocked_days=[4]),
    ]
    streams = [
        Stream(id="S1", name="SE поток", group_ids=["G1", "G2"], preferred_room_id="R101"),
    ]

    disciplines = [
        Discipline(id="D1", name="Алгоритмы", credits=5, required_room_features=["lecture"], kind="lecture"),
        Discipline(id="D2", name="Базы данных (лаб)", credits=4, required_room_features=["lab"], kind="lab"),
        Discipline(id="D3", name="Практика ИТ", credits=3, required_room_features=["practice"], kind="practice"),
        Discipline(
            id="D4",
            name="Высшая математика",
            credits=5,
            required_room_features=["lecture"],
            kind="lecture",
        ),
    ]

    assignments = [
        Assignment(
            id="A1",
            discipline_id="D1",
            teacher_id="T1",
            stream_id="S1",
            stream_name="SE поток",
            start_week=1,
            end_week=15,
            sessions_per_week=2,
            room_id=None,
        ),
        Assignment(
            id="A2",
            discipline_id="D2",
            teacher_id="T3",
            group_ids=["G1"],
            start_week=1,
            end_week=15,
            sessions_per_week=1,
            room_id="LAB1",
        ),
        Assignment(
            id="A3",
            discipline_id="D2",
            teacher_id="T3",
            group_ids=["G2"],
            start_week=1,
            end_week=15,
            sessions_per_week=1,
            room_id="LAB2",
        ),
        Assignment(
            id="A4",
            discipline_id="D4",
            teacher_id="T2",
            group_ids=["G3"],
            start_week=1,
            end_week=15,
            sessions_per_week=2,
        ),
        Assignment(
            id="A5",
            discipline_id="D3",
            teacher_id="T1",
            group_ids=["G3"],
            start_week=5,
            end_week=7,
            sessions_per_week=2,
            room_id="LAB1",
        ),
    ]

    return AppState(
        teachers=teachers,
        rooms=rooms,
        groups=groups,
        streams=streams,
        disciplines=disciplines,
        assignments=assignments,
        schedule=[],
    )
