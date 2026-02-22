from __future__ import annotations

from dataclasses import asdict, dataclass, field
from typing import Any

DAYS = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat"]
SLOTS = [1, 2, 3, 4, 5, 6, 7, 8]


def slot_key(day: int, slot: int) -> str:
    return f"{day}:{slot}"


@dataclass
class Teacher:
    id: str
    name: str
    department: str
    position: str
    degree: str
    default_room_id: str | None = None
    work_days: list[int] = field(default_factory=lambda: [0, 1, 2, 3, 4, 5])
    blocked_slots: list[str] = field(default_factory=list)
    max_classes_per_day: int = 4
    max_classes_per_week: int = 20
    contract_start_week: int = 1
    contract_end_week: int = 15


@dataclass
class Room:
    id: str
    name: str
    capacity: int
    building: str
    features: list[str] = field(default_factory=list)


@dataclass
class StudentGroup:
    id: str
    name: str
    size: int
    shift_start_slot: int = 1
    shift_end_slot: int = 8
    blocked_days: list[int] = field(default_factory=list)
    blocked_slots: list[str] = field(default_factory=list)
    program_start_week: int = 1
    program_end_week: int = 15


@dataclass
class Stream:
    id: str
    name: str
    group_ids: list[str]
    preferred_room_id: str | None = None
    notes: str = ""


@dataclass
class Discipline:
    id: str
    name: str
    credits: int
    required_room_features: list[str] = field(default_factory=list)
    fixed_room_id: str | None = None
    kind: str = "lecture"  # lecture | lab | practice
    split_by_subgroups: bool = False
    practice_as_lab_exception: bool = True


@dataclass
class Assignment:
    id: str
    discipline_id: str
    teacher_id: str
    group_ids: list[str] = field(default_factory=list)
    stream_id: str | None = None
    stream_name: str = ""
    start_week: int = 1
    end_week: int = 15
    sessions_per_week: int = 2
    duration_slots: int = 1
    room_id: str | None = None
    lock_day: int | None = None
    lock_slot: int | None = None
    lock_room_id: str | None = None
    lock_teacher_id: str | None = None
    notes: str = ""


@dataclass
class ScheduleEntry:
    assignment_id: str
    discipline_id: str
    teacher_id: str
    group_ids: list[str]
    day: int
    slot: int
    room_id: str
    start_week: int
    end_week: int


@dataclass
class AppState:
    teachers: list[Teacher] = field(default_factory=list)
    rooms: list[Room] = field(default_factory=list)
    groups: list[StudentGroup] = field(default_factory=list)
    streams: list[Stream] = field(default_factory=list)
    disciplines: list[Discipline] = field(default_factory=list)
    assignments: list[Assignment] = field(default_factory=list)
    schedule: list[ScheduleEntry] = field(default_factory=list)

    def to_dict(self) -> dict[str, Any]:
        return {
            "teachers": [asdict(t) for t in self.teachers],
            "rooms": [asdict(r) for r in self.rooms],
            "groups": [asdict(g) for g in self.groups],
            "streams": [asdict(s) for s in self.streams],
            "disciplines": [asdict(d) for d in self.disciplines],
            "assignments": [asdict(a) for a in self.assignments],
            "schedule": [asdict(s) for s in self.schedule],
        }

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> "AppState":
        return cls(
            teachers=[Teacher(**item) for item in data.get("teachers", [])],
            rooms=[Room(**item) for item in data.get("rooms", [])],
            groups=[StudentGroup(**item) for item in data.get("groups", [])],
            streams=[Stream(**item) for item in data.get("streams", [])],
            disciplines=[Discipline(**item) for item in data.get("disciplines", [])],
            assignments=[Assignment(**item) for item in data.get("assignments", [])],
            schedule=[ScheduleEntry(**item) for item in data.get("schedule", [])],
        )
