from __future__ import annotations

import json
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

import customtkinter as ctk
from openpyxl import Workbook

from scheduler_desktop.excel_io import export_state_to_excel, import_state_from_excel
from scheduler_desktop.models import DAYS, SLOTS, AppState, Assignment, ScheduleEntry, slot_key
from scheduler_desktop.planning import PlanningError, ScheduleGenerator
from scheduler_desktop.repository import StateRepository, sample_state

DAY_LABELS_RU = ["Пн", "Вт", "Ср", "Чт", "Пт", "Сб"]


class SchedulerDesktopApp(ctk.CTk):
    def __init__(self, repository: StateRepository) -> None:
        super().__init__()
        self.repo = repository
        self.app_state = self.repo.load()
        self.title("UniSchedule Generator")
        self.geometry("1420x860")
        self.minsize(1220, 760)
        ctk.set_appearance_mode("light")
        ctk.set_default_color_theme("green")
        self.configure(fg_color="#f4f7f8")
        self._setup_treeview_style()

        self.pages: dict[str, ctk.CTkFrame] = {}
        self._build_layout()
        self._build_data_page()
        self._build_generation_page()
        self._build_rooms_page()
        self._build_teachers_page()
        self._build_replacements_page()
        self._build_analysis_page()
        self.show_page("data")
        self.refresh_ui()

    def _setup_treeview_style(self) -> None:
        style = ttk.Style(self)
        style.theme_use("default")
        style.configure(
            "Treeview",
            background="#ffffff",
            fieldbackground="#ffffff",
            rowheight=30,
            borderwidth=0,
            font=("Segoe UI", 10),
        )
        style.configure(
            "Treeview.Heading",
            background="#e9eef0",
            relief="flat",
            borderwidth=0,
            font=("Segoe UI Semibold", 10),
        )
        style.map("Treeview", background=[("selected", "#d7efe4")], foreground=[("selected", "#153a2f")])

    def _build_layout(self) -> None:
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        sidebar = ctk.CTkFrame(self, corner_radius=0, fg_color="#153a2f", width=250)
        sidebar.grid(row=0, column=0, sticky="nsew")
        sidebar.grid_propagate(False)
        sidebar.grid_rowconfigure(10, weight=1)

        ctk.CTkLabel(
            sidebar,
            text="UniSchedule",
            font=("Segoe UI Semibold", 28),
            text_color="#e9f5ef",
        ).grid(row=0, column=0, padx=24, pady=(26, 4), sticky="w")
        ctk.CTkLabel(
            sidebar,
            text="Генерация расписаний ВУЗа",
            font=("Segoe UI", 12),
            text_color="#b6d8c9",
        ).grid(row=1, column=0, padx=24, pady=(0, 22), sticky="w")

        nav_items = [
            ("data", "Данные"),
            ("generation", "Генерация"),
            ("rooms", "Занятость аудиторий"),
            ("teachers", "Расписание ППС"),
            ("replace", "Замены"),
            ("analysis", "Анализ ПО"),
        ]
        self.nav_buttons: dict[str, ctk.CTkButton] = {}
        for idx, (key, label) in enumerate(nav_items, start=2):
            btn = ctk.CTkButton(
                sidebar,
                text=label,
                font=("Segoe UI Semibold", 15),
                anchor="w",
                corner_radius=10,
                fg_color="transparent",
                hover_color="#245342",
                text_color="#d7efe4",
                command=lambda k=key: self.show_page(k),
            )
            btn.grid(row=idx, column=0, padx=16, pady=5, sticky="ew")
            self.nav_buttons[key] = btn

        self.status_label = ctk.CTkLabel(
            sidebar,
            text="",
            font=("Segoe UI", 11),
            text_color="#b6d8c9",
            justify="left",
            wraplength=210,
        )
        self.status_label.grid(row=11, column=0, padx=16, pady=16, sticky="sw")

        self.content = ctk.CTkFrame(self, fg_color="#f4f7f8", corner_radius=0)
        self.content.grid(row=0, column=1, sticky="nsew")
        self.content.grid_columnconfigure(0, weight=1)
        self.content.grid_rowconfigure(0, weight=1)

        for key in ["data", "generation", "rooms", "teachers", "replace", "analysis"]:
            page = ctk.CTkFrame(self.content, fg_color="#f4f7f8")
            page.grid(row=0, column=0, sticky="nsew")
            page.grid_remove()
            self.pages[key] = page

    def show_page(self, key: str) -> None:
        for page_key, page in self.pages.items():
            if page_key == key:
                page.grid()
            else:
                page.grid_remove()

        for btn_key, btn in self.nav_buttons.items():
            btn.configure(fg_color="#2f6d58" if btn_key == key else "transparent")

    def _build_data_page(self) -> None:
        page = self.pages["data"]
        page.grid_columnconfigure(0, weight=1)
        page.grid_rowconfigure(2, weight=1)

        header = ctk.CTkFrame(page, fg_color="#e3efea", corner_radius=18)
        header.grid(row=0, column=0, sticky="ew", padx=22, pady=(18, 10))
        header.grid_columnconfigure(6, weight=1)

        ctk.CTkLabel(
            header,
            text="Данные и ограничения",
            font=("Segoe UI Semibold", 24),
            text_color="#163b30",
        ).grid(row=0, column=0, padx=20, pady=14, sticky="w")

        ctk.CTkButton(header, text="Демо-набор", command=self.load_sample).grid(row=0, column=1, padx=6, pady=14)
        ctk.CTkButton(header, text="Импорт Excel", command=self.import_excel).grid(row=0, column=2, padx=6, pady=14)
        ctk.CTkButton(header, text="Экспорт Excel", command=self.export_excel).grid(row=0, column=3, padx=6, pady=14)
        ctk.CTkButton(header, text="Кредиты -> пары", command=self.recompute_sessions_from_credits).grid(
            row=0, column=4, padx=6, pady=14
        )
        ctk.CTkButton(header, text="Сохранить", command=self.save_state).grid(row=0, column=5, padx=(6, 20), pady=14)

        cards = ctk.CTkFrame(page, fg_color="#f4f7f8")
        cards.grid(row=1, column=0, sticky="ew", padx=22, pady=(0, 10))
        cards.grid_columnconfigure((0, 1, 2, 3, 4), weight=1)
        self.card_labels: dict[str, ctk.CTkLabel] = {}
        for idx, key in enumerate(["teachers", "rooms", "groups", "disciplines", "assignments"]):
            card = ctk.CTkFrame(cards, fg_color="#ffffff", corner_radius=14, border_width=1, border_color="#d3dfda")
            card.grid(row=0, column=idx, padx=6, sticky="ew")
            ctk.CTkLabel(card, text=key.upper(), text_color="#5f796f", font=("Segoe UI", 11)).grid(
                row=0, column=0, padx=10, pady=(9, 3), sticky="w"
            )
            lbl = ctk.CTkLabel(card, text="0", text_color="#153a2f", font=("Segoe UI Semibold", 24))
            lbl.grid(row=1, column=0, padx=10, pady=(0, 10), sticky="w")
            self.card_labels[key] = lbl

        body = ctk.CTkFrame(page, fg_color="#f4f7f8")
        body.grid(row=2, column=0, sticky="nsew", padx=22, pady=(0, 20))
        body.grid_columnconfigure(0, weight=5)
        body.grid_columnconfigure(1, weight=3)
        body.grid_rowconfigure(0, weight=1)

        left = ctk.CTkFrame(body, fg_color="#ffffff", corner_radius=16, border_width=1, border_color="#d3dfda")
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        left.grid_columnconfigure(0, weight=1)
        left.grid_rowconfigure(1, weight=1)
        ctk.CTkLabel(left, text="JSON данных проекта", font=("Segoe UI Semibold", 16)).grid(
            row=0, column=0, padx=14, pady=(12, 8), sticky="w"
        )
        self.json_box = ctk.CTkTextbox(left, font=("Consolas", 12), wrap="none")
        self.json_box.grid(row=1, column=0, padx=12, pady=(0, 8), sticky="nsew")
        btns = ctk.CTkFrame(left, fg_color="transparent")
        btns.grid(row=2, column=0, padx=12, pady=(0, 12), sticky="ew")
        ctk.CTkButton(btns, text="Обновить JSON из состояния", command=self.populate_json_box).pack(side="left", padx=(0, 8))
        ctk.CTkButton(btns, text="Применить JSON в приложение", command=self.apply_json_box).pack(side="left")

        right = ctk.CTkFrame(body, fg_color="#ffffff", corner_radius=16, border_width=1, border_color="#d3dfda")
        right.grid(row=0, column=1, sticky="nsew")
        right.grid_columnconfigure(0, weight=1)
        right.grid_rowconfigure(4, weight=1)

        ctk.CTkLabel(right, text="График ППС (дни/слоты)", font=("Segoe UI Semibold", 16)).grid(
            row=0, column=0, padx=12, pady=(12, 8), sticky="w"
        )
        self.teacher_pick_avail = ctk.CTkComboBox(right, values=[""], width=230)
        self.teacher_pick_avail.grid(row=1, column=0, padx=12, pady=4, sticky="w")
        ctk.CTkButton(right, text="Редактировать занятость", command=self.edit_teacher_availability).grid(
            row=2, column=0, padx=12, pady=8, sticky="w"
        )

        ctk.CTkLabel(right, text="Назначения дисциплин", font=("Segoe UI Semibold", 16)).grid(
            row=3, column=0, padx=12, pady=(12, 6), sticky="w"
        )
        columns = ("id", "disc", "teacher", "groups", "weeks", "pairs")
        self.assignment_tree = ttk.Treeview(right, columns=columns, show="headings", height=11)
        for col, text, width in [
            ("id", "ID", 55),
            ("disc", "Дисциплина", 180),
            ("teacher", "ППС", 150),
            ("groups", "Группы/поток", 140),
            ("weeks", "Недели", 90),
            ("pairs", "Пар/нед", 80),
        ]:
            self.assignment_tree.heading(col, text=text)
            self.assignment_tree.column(col, width=width, stretch=(col in {"disc", "teacher", "groups"}))
        self.assignment_tree.grid(row=4, column=0, padx=12, pady=(0, 12), sticky="nsew")

    def _build_generation_page(self) -> None:
        page = self.pages["generation"]
        page.grid_columnconfigure(0, weight=1)
        page.grid_rowconfigure(2, weight=1)

        top = ctk.CTkFrame(page, fg_color="#ffffff", corner_radius=16, border_width=1, border_color="#d3dfda")
        top.grid(row=0, column=0, sticky="ew", padx=22, pady=(18, 10))
        top.grid_columnconfigure(4, weight=1)
        ctk.CTkLabel(top, text="Генерация расписания", font=("Segoe UI Semibold", 24), text_color="#163b30").grid(
            row=0, column=0, padx=16, pady=14, sticky="w"
        )
        ctk.CTkButton(top, text="Сгенерировать", command=self.generate_schedule).grid(row=0, column=1, padx=8, pady=14)
        ctk.CTkButton(top, text="Очистить расписание", command=self.clear_schedule).grid(row=0, column=2, padx=8, pady=14)
        ctk.CTkButton(top, text="Экспорт расписания Excel", command=self.export_schedule_excel).grid(
            row=0, column=3, padx=8, pady=14
        )
        self.generation_result = ctk.CTkLabel(top, text="", font=("Segoe UI", 13), text_color="#3f6257")
        self.generation_result.grid(row=0, column=4, padx=16, sticky="w")

        legend = ctk.CTkLabel(
            page,
            text="Строка = фиксированный слот недели, колонка 'Недели' показывает период активности дисциплины.",
            font=("Segoe UI", 12),
            text_color="#3f6257",
        )
        legend.grid(row=1, column=0, padx=26, pady=(0, 8), sticky="w")

        schedule_frame = ctk.CTkFrame(page, fg_color="#ffffff", corner_radius=16, border_width=1, border_color="#d3dfda")
        schedule_frame.grid(row=2, column=0, sticky="nsew", padx=22, pady=(0, 20))
        schedule_frame.grid_columnconfigure(0, weight=1)
        schedule_frame.grid_rowconfigure(0, weight=1)

        columns = ("assignment", "discipline", "teacher", "groups", "day", "slot", "room", "weeks")
        self.schedule_tree = ttk.Treeview(schedule_frame, columns=columns, show="headings", height=16)
        for col, text, width in [
            ("assignment", "Назначение", 90),
            ("discipline", "Дисциплина", 180),
            ("teacher", "ППС", 140),
            ("groups", "Группы", 160),
            ("day", "День", 70),
            ("slot", "Пара", 65),
            ("room", "Аудитория", 100),
            ("weeks", "Недели", 90),
        ]:
            self.schedule_tree.heading(col, text=text)
            self.schedule_tree.column(col, width=width, stretch=(col in {"discipline", "teacher", "groups"}))
        self.schedule_tree.grid(row=0, column=0, padx=12, pady=12, sticky="nsew")

    def _build_rooms_page(self) -> None:
        page = self.pages["rooms"]
        page.grid_columnconfigure(0, weight=1)
        page.grid_rowconfigure(1, weight=1)

        top = ctk.CTkFrame(page, fg_color="#ffffff", corner_radius=16, border_width=1, border_color="#d3dfda")
        top.grid(row=0, column=0, sticky="ew", padx=22, pady=(18, 10))
        ctk.CTkLabel(top, text="Просмотр занятости аудитории", font=("Segoe UI Semibold", 22)).grid(
            row=0, column=0, padx=16, pady=14, sticky="w"
        )
        self.room_pick = ctk.CTkComboBox(top, values=[""], width=200)
        self.room_pick.grid(row=0, column=1, padx=8, pady=14)
        self.room_week_entry = ctk.CTkEntry(top, width=80, placeholder_text="Неделя")
        self.room_week_entry.grid(row=0, column=2, padx=8, pady=14)
        ctk.CTkButton(top, text="Показать", command=self.refresh_room_grid).grid(row=0, column=3, padx=8, pady=14)

        grid = ctk.CTkFrame(page, fg_color="#ffffff", corner_radius=16, border_width=1, border_color="#d3dfda")
        grid.grid(row=1, column=0, sticky="nsew", padx=22, pady=(0, 20))
        grid.grid_columnconfigure(0, weight=1)
        grid.grid_rowconfigure(0, weight=1)

        room_columns = ("pair", *DAY_LABELS_RU)
        self.room_grid = ttk.Treeview(grid, columns=room_columns, show="headings")
        self.room_grid.heading("pair", text="Пара")
        self.room_grid.column("pair", width=65, stretch=False)
        for day in DAY_LABELS_RU:
            self.room_grid.heading(day, text=day)
            self.room_grid.column(day, width=170, stretch=True)
        self.room_grid.grid(row=0, column=0, padx=12, pady=12, sticky="nsew")

    def _build_teachers_page(self) -> None:
        page = self.pages["teachers"]
        page.grid_columnconfigure(0, weight=1)
        page.grid_rowconfigure(1, weight=1)

        top = ctk.CTkFrame(page, fg_color="#ffffff", corner_radius=16, border_width=1, border_color="#d3dfda")
        top.grid(row=0, column=0, sticky="ew", padx=22, pady=(18, 10))
        ctk.CTkLabel(top, text="Расписание преподавателя", font=("Segoe UI Semibold", 22)).grid(
            row=0, column=0, padx=16, pady=14, sticky="w"
        )
        self.teacher_pick = ctk.CTkComboBox(top, values=[""], width=230)
        self.teacher_pick.grid(row=0, column=1, padx=8, pady=14)
        self.teacher_week_entry = ctk.CTkEntry(top, width=80, placeholder_text="Неделя")
        self.teacher_week_entry.grid(row=0, column=2, padx=8, pady=14)
        ctk.CTkButton(top, text="Показать", command=self.refresh_teacher_grid).grid(row=0, column=3, padx=8, pady=14)

        grid = ctk.CTkFrame(page, fg_color="#ffffff", corner_radius=16, border_width=1, border_color="#d3dfda")
        grid.grid(row=1, column=0, sticky="nsew", padx=22, pady=(0, 20))
        grid.grid_columnconfigure(0, weight=1)
        grid.grid_rowconfigure(0, weight=1)

        teacher_columns = ("pair", *DAY_LABELS_RU)
        self.teacher_grid = ttk.Treeview(grid, columns=teacher_columns, show="headings")
        self.teacher_grid.heading("pair", text="Пара")
        self.teacher_grid.column("pair", width=65, stretch=False)
        for day in DAY_LABELS_RU:
            self.teacher_grid.heading(day, text=day)
            self.teacher_grid.column(day, width=170, stretch=True)
        self.teacher_grid.grid(row=0, column=0, padx=12, pady=12, sticky="nsew")

    def _build_replacements_page(self) -> None:
        page = self.pages["replace"]
        page.grid_columnconfigure(0, weight=1)
        page.grid_rowconfigure(1, weight=1)

        panel = ctk.CTkFrame(page, fg_color="#ffffff", corner_radius=16, border_width=1, border_color="#d3dfda")
        panel.grid(row=0, column=0, sticky="ew", padx=22, pady=(18, 10))
        ctk.CTkLabel(panel, text="Замена аудитории или ППС", font=("Segoe UI Semibold", 22)).grid(
            row=0, column=0, padx=16, pady=(12, 10), sticky="w"
        )

        ctk.CTkLabel(panel, text="Назначение").grid(row=1, column=0, padx=16, pady=6, sticky="w")
        self.assignment_pick_replace = ctk.CTkComboBox(panel, values=[""], width=220)
        self.assignment_pick_replace.grid(row=1, column=1, padx=6, pady=6, sticky="w")
        ctk.CTkLabel(panel, text="Новый ППС").grid(row=1, column=2, padx=6, pady=6, sticky="w")
        self.replace_teacher_pick = ctk.CTkComboBox(panel, values=[""], width=220)
        self.replace_teacher_pick.grid(row=1, column=3, padx=6, pady=6, sticky="w")
        ctk.CTkLabel(panel, text="Новая аудитория").grid(row=1, column=4, padx=6, pady=6, sticky="w")
        self.replace_room_pick = ctk.CTkComboBox(panel, values=[""], width=220)
        self.replace_room_pick.grid(row=1, column=5, padx=6, pady=6, sticky="w")

        self.lock_slot_switch = ctk.CTkSwitch(panel, text="Зафиксировать текущий слот")
        self.lock_slot_switch.grid(row=1, column=6, padx=(10, 16), pady=6)
        ctk.CTkButton(panel, text="Применить замену", command=self.apply_replacement).grid(
            row=2, column=0, padx=16, pady=(8, 14), sticky="w"
        )
        ctk.CTkButton(panel, text="Снять фиксацию слота", command=self.clear_slot_lock).grid(
            row=2, column=1, padx=6, pady=(8, 14), sticky="w"
        )
        ctk.CTkButton(panel, text="Перегенерировать", command=self.generate_schedule).grid(
            row=2, column=2, padx=6, pady=(8, 14), sticky="w"
        )

        grid = ctk.CTkFrame(page, fg_color="#ffffff", corner_radius=16, border_width=1, border_color="#d3dfda")
        grid.grid(row=1, column=0, sticky="nsew", padx=22, pady=(0, 20))
        grid.grid_columnconfigure(0, weight=1)
        grid.grid_rowconfigure(0, weight=1)

        columns = ("assignment", "disc", "teacher", "room", "day", "slot", "weeks")
        self.replace_tree = ttk.Treeview(grid, columns=columns, show="headings")
        for col, text, width in [
            ("assignment", "Назначение", 90),
            ("disc", "Дисциплина", 170),
            ("teacher", "ППС", 150),
            ("room", "Аудитория", 110),
            ("day", "День", 70),
            ("slot", "Пара", 70),
            ("weeks", "Недели", 90),
        ]:
            self.replace_tree.heading(col, text=text)
            self.replace_tree.column(col, width=width, stretch=(col in {"disc", "teacher"}))
        self.replace_tree.grid(row=0, column=0, padx=12, pady=12, sticky="nsew")

    def _build_analysis_page(self) -> None:
        page = self.pages["analysis"]
        page.grid_columnconfigure(0, weight=1)
        page.grid_rowconfigure(1, weight=1)

        header = ctk.CTkFrame(page, fg_color="#ffffff", corner_radius=16, border_width=1, border_color="#d3dfda")
        header.grid(row=0, column=0, sticky="ew", padx=22, pady=(18, 10))
        ctk.CTkLabel(header, text="Анализ систем расписаний (Sprut и др.)", font=("Segoe UI Semibold", 22)).grid(
            row=0, column=0, padx=16, pady=14, sticky="w"
        )

        box_frame = ctk.CTkFrame(page, fg_color="#ffffff", corner_radius=16, border_width=1, border_color="#d3dfda")
        box_frame.grid(row=1, column=0, sticky="nsew", padx=22, pady=(0, 20))
        box_frame.grid_columnconfigure(0, weight=1)
        box_frame.grid_rowconfigure(0, weight=1)
        self.analysis_box = ctk.CTkTextbox(box_frame, wrap="word", font=("Segoe UI", 12))
        self.analysis_box.grid(row=0, column=0, padx=12, pady=12, sticky="nsew")

    def refresh_ui(self) -> None:
        self.populate_json_box()
        self.refresh_cards()
        self.refresh_assignment_tree()
        self.refresh_schedule_tree()
        self.refresh_replace_tree()
        self.refresh_comboboxes()
        self.refresh_analysis()

    def refresh_cards(self) -> None:
        self.card_labels["teachers"].configure(text=str(len(self.app_state.teachers)))
        self.card_labels["rooms"].configure(text=str(len(self.app_state.rooms)))
        self.card_labels["groups"].configure(text=str(len(self.app_state.groups)))
        self.card_labels["disciplines"].configure(text=str(len(self.app_state.disciplines)))
        self.card_labels["assignments"].configure(text=str(len(self.app_state.assignments)))

    def populate_json_box(self) -> None:
        payload = json.dumps(self.app_state.to_dict(), ensure_ascii=False, indent=2)
        self.json_box.delete("1.0", "end")
        self.json_box.insert("1.0", payload)

    def apply_json_box(self) -> None:
        try:
            data = json.loads(self.json_box.get("1.0", "end"))
            self.app_state = AppState.from_dict(data)
            self.refresh_ui()
            self._set_status("JSON применен")
        except Exception as exc:
            messagebox.showerror("Ошибка JSON", str(exc))

    def load_sample(self) -> None:
        self.app_state = sample_state()
        self.refresh_ui()
        self._set_status("Загружен демонстрационный набор")

    def import_excel(self) -> None:
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if not path:
            return
        try:
            self.app_state = import_state_from_excel(Path(path))
            self.refresh_ui()
            self._set_status(f"Импортировано: {Path(path).name}")
        except Exception as exc:
            messagebox.showerror("Импорт", str(exc))

    def export_excel(self) -> None:
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if not path:
            return
        try:
            export_state_to_excel(Path(path), self.app_state)
            self._set_status(f"Экспортировано: {Path(path).name}")
        except Exception as exc:
            messagebox.showerror("Экспорт", str(exc))

    def save_state(self) -> None:
        try:
            self.repo.save(self.app_state)
            self._set_status(f"Сохранено в {self.repo.path.name}")
        except Exception as exc:
            messagebox.showerror("Сохранение", str(exc))

    def recompute_sessions_from_credits(self) -> None:
        disc_map = {d.id: d for d in self.app_state.disciplines}
        changed = 0
        for assignment in self.app_state.assignments:
            discipline = disc_map.get(assignment.discipline_id)
            if not discipline:
                continue
            weeks = max(1, assignment.end_week - assignment.start_week + 1)
            contact_hours = discipline.credits * 15
            sessions = max(1, round(contact_hours / (weeks * 1.5)))
            assignment.sessions_per_week = sessions
            changed += 1
        self.refresh_assignment_tree()
        self.populate_json_box()
        self._set_status(f"Обновлено пар/нед по кредитам: {changed}")

    def refresh_assignment_tree(self) -> None:
        for item in self.assignment_tree.get_children():
            self.assignment_tree.delete(item)
        disc_map = {d.id: d.name for d in self.app_state.disciplines}
        teacher_map = {t.id: t.name for t in self.app_state.teachers}
        stream_map = {s.id: s for s in self.app_state.streams}
        for assignment in self.app_state.assignments:
            groups_label = ", ".join(self._assignment_effective_group_ids(assignment))
            if assignment.stream_id and assignment.stream_id in stream_map:
                stream = stream_map[assignment.stream_id]
                groups_label = f"{stream.name}: {groups_label}"
            self.assignment_tree.insert(
                "",
                "end",
                values=(
                    assignment.id,
                    disc_map.get(assignment.discipline_id, assignment.discipline_id),
                    teacher_map.get(assignment.teacher_id, assignment.teacher_id),
                    groups_label,
                    f"{assignment.start_week}-{assignment.end_week}",
                    assignment.sessions_per_week,
                ),
            )

    def generate_schedule(self) -> None:
        try:
            schedule = ScheduleGenerator(self.app_state).generate()
            self.app_state.schedule = schedule
            self.refresh_schedule_tree()
            self.refresh_replace_tree()
            self.refresh_room_grid()
            self.refresh_teacher_grid()
            self.populate_json_box()
            self.generation_result.configure(text=f"Сгенерировано: {len(schedule)} строк")
            self._set_status(f"Расписание построено: {len(schedule)} строк")
        except PlanningError as exc:
            messagebox.showerror("Генерация", str(exc))
            self._set_status("Генерация не удалась")

    def clear_schedule(self) -> None:
        self.app_state.schedule = []
        self.refresh_schedule_tree()
        self.refresh_replace_tree()
        self.refresh_room_grid()
        self.refresh_teacher_grid()
        self.populate_json_box()
        self.generation_result.configure(text="Расписание очищено")
        self._set_status("Расписание очищено")

    def refresh_schedule_tree(self) -> None:
        for item in self.schedule_tree.get_children():
            self.schedule_tree.delete(item)

        disc_map = {d.id: d.name for d in self.app_state.disciplines}
        teacher_map = {t.id: t.name for t in self.app_state.teachers}
        room_map = {r.id: r.name for r in self.app_state.rooms}
        for entry in sorted(self.app_state.schedule, key=lambda x: (x.day, x.slot, x.assignment_id)):
            self.schedule_tree.insert(
                "",
                "end",
                values=(
                    entry.assignment_id,
                    disc_map.get(entry.discipline_id, entry.discipline_id),
                    teacher_map.get(entry.teacher_id, entry.teacher_id),
                    ", ".join(entry.group_ids),
                    DAY_LABELS_RU[entry.day],
                    entry.slot,
                    room_map.get(entry.room_id, entry.room_id),
                    f"{entry.start_week}-{entry.end_week}",
                ),
            )

    def refresh_comboboxes(self) -> None:
        teacher_values = [f"{t.id} | {t.name}" for t in self.app_state.teachers]
        room_values = [f"{r.id} | {r.name}" for r in self.app_state.rooms]
        assignment_values = [a.id for a in self.app_state.assignments]

        self.teacher_pick.configure(values=teacher_values or [""])
        self.teacher_pick_avail.configure(values=teacher_values or [""])
        self.replace_teacher_pick.configure(values=[""] + teacher_values)
        self.room_pick.configure(values=room_values or [""])
        self.replace_room_pick.configure(values=[""] + room_values)
        self.assignment_pick_replace.configure(values=assignment_values or [""])

        if teacher_values and not self.teacher_pick.get():
            self.teacher_pick.set(teacher_values[0])
        if teacher_values and not self.teacher_pick_avail.get():
            self.teacher_pick_avail.set(teacher_values[0])
        if room_values and not self.room_pick.get():
            self.room_pick.set(room_values[0])
        if assignment_values and not self.assignment_pick_replace.get():
            self.assignment_pick_replace.set(assignment_values[0])

    def refresh_room_grid(self) -> None:
        for item in self.room_grid.get_children():
            self.room_grid.delete(item)
        selected_room = self._id_from_combo(self.room_pick.get())
        week = self._read_week(self.room_week_entry.get())
        if not selected_room:
            return

        matrix = self._build_matrix()
        for slot in SLOTS:
            row = [matrix.get((week, day, slot, "room", selected_room), "") for day in range(len(DAYS))]
            self.room_grid.insert("", "end", values=(slot, *row))

    def refresh_teacher_grid(self) -> None:
        for item in self.teacher_grid.get_children():
            self.teacher_grid.delete(item)
        teacher_id = self._id_from_combo(self.teacher_pick.get())
        week = self._read_week(self.teacher_week_entry.get())
        if not teacher_id:
            return

        matrix = self._build_matrix()
        for slot in SLOTS:
            row = [matrix.get((week, day, slot, "teacher", teacher_id), "") for day in range(len(DAYS))]
            self.teacher_grid.insert("", "end", values=(slot, *row))

    def _build_matrix(self) -> dict[tuple[int, int, int, str, str], str]:
        disc_map = {d.id: d.name for d in self.app_state.disciplines}
        room_map = {r.id: r.name for r in self.app_state.rooms}
        matrix: dict[tuple[int, int, int, str, str], str] = {}
        for entry in self.app_state.schedule:
            for week in range(entry.start_week, entry.end_week + 1):
                matrix[(week, entry.day, entry.slot, "room", entry.room_id)] = (
                    f"{entry.assignment_id} {disc_map.get(entry.discipline_id, '')}"
                )
                matrix[(week, entry.day, entry.slot, "teacher", entry.teacher_id)] = (
                    f"{entry.assignment_id} {disc_map.get(entry.discipline_id, '')} [{room_map.get(entry.room_id, '')}]"
                )
        return matrix

    def refresh_replace_tree(self) -> None:
        for item in self.replace_tree.get_children():
            self.replace_tree.delete(item)
        disc_map = {d.id: d.name for d in self.app_state.disciplines}
        teacher_map = {t.id: t.name for t in self.app_state.teachers}
        room_map = {r.id: r.name for r in self.app_state.rooms}
        for entry in self.app_state.schedule:
            self.replace_tree.insert(
                "",
                "end",
                values=(
                    entry.assignment_id,
                    disc_map.get(entry.discipline_id, entry.discipline_id),
                    teacher_map.get(entry.teacher_id, entry.teacher_id),
                    room_map.get(entry.room_id, entry.room_id),
                    DAY_LABELS_RU[entry.day],
                    entry.slot,
                    f"{entry.start_week}-{entry.end_week}",
                ),
            )

    def apply_replacement(self) -> None:
        assignment_id = self.assignment_pick_replace.get().strip()
        if not assignment_id:
            return
        assignment = self._find_assignment(assignment_id)
        if not assignment:
            messagebox.showerror("Замена", f"Назначение {assignment_id} не найдено")
            return

        teacher_id = self._id_from_combo(self.replace_teacher_pick.get())
        room_id = self._id_from_combo(self.replace_room_pick.get())
        if teacher_id:
            assignment.teacher_id = teacher_id
        if room_id:
            assignment.room_id = room_id

        if self.lock_slot_switch.get() == 1:
            entry = self._first_schedule_entry(assignment_id)
            if entry:
                assignment.lock_day = entry.day
                assignment.lock_slot = entry.slot
                assignment.lock_room_id = entry.room_id
        self.populate_json_box()
        self.refresh_assignment_tree()
        self._set_status(f"Замена применена для {assignment_id}")

    def clear_slot_lock(self) -> None:
        assignment_id = self.assignment_pick_replace.get().strip()
        assignment = self._find_assignment(assignment_id)
        if not assignment:
            return
        assignment.lock_day = None
        assignment.lock_slot = None
        assignment.lock_room_id = None
        self.populate_json_box()
        self._set_status(f"Фиксация снята для {assignment_id}")

    def _first_schedule_entry(self, assignment_id: str) -> ScheduleEntry | None:
        for entry in self.app_state.schedule:
            if entry.assignment_id == assignment_id:
                return entry
        return None

    def _find_assignment(self, assignment_id: str) -> Assignment | None:
        for assignment in self.app_state.assignments:
            if assignment.id == assignment_id:
                return assignment
        return None

    def _assignment_effective_group_ids(self, assignment: Assignment) -> list[str]:
        group_ids = set(assignment.group_ids)
        if assignment.stream_id:
            stream = next((s for s in self.app_state.streams if s.id == assignment.stream_id), None)
            if stream:
                group_ids.update(stream.group_ids)
        return sorted(group_ids)

    def edit_teacher_availability(self) -> None:
        teacher_id = self._id_from_combo(self.teacher_pick_avail.get())
        teacher = next((t for t in self.app_state.teachers if t.id == teacher_id), None)
        if not teacher:
            return

        window = ctk.CTkToplevel(self)
        window.title(f"График ППС: {teacher.name}")
        window.geometry("860x520")
        window.grab_set()
        window.grid_columnconfigure(1, weight=1)
        window.grid_rowconfigure(1, weight=1)

        ctk.CTkLabel(window, text="Рабочие дни", font=("Segoe UI Semibold", 15)).grid(
            row=0, column=0, padx=14, pady=(12, 8), sticky="w"
        )
        day_frame = ctk.CTkFrame(window, fg_color="transparent")
        day_frame.grid(row=0, column=1, padx=14, pady=(10, 8), sticky="w")
        day_vars: dict[int, ctk.IntVar] = {}
        for day in range(len(DAYS)):
            var = ctk.IntVar(value=1 if day in teacher.work_days else 0)
            day_vars[day] = var
            ctk.CTkCheckBox(day_frame, text=DAY_LABELS_RU[day], variable=var).pack(side="left", padx=6)

        ctk.CTkLabel(window, text="Недоступные слоты (галочка = нельзя занимать)", font=("Segoe UI Semibold", 14)).grid(
            row=1, column=0, columnspan=2, padx=14, pady=(8, 4), sticky="w"
        )
        matrix = ctk.CTkScrollableFrame(window, fg_color="#ffffff", width=820, height=350)
        matrix.grid(row=2, column=0, columnspan=2, padx=14, pady=(0, 10), sticky="nsew")

        slot_vars: dict[str, ctk.IntVar] = {}
        for day in range(len(DAYS)):
            ctk.CTkLabel(matrix, text=DAY_LABELS_RU[day], font=("Segoe UI Semibold", 13)).grid(
                row=0, column=day + 1, padx=6, pady=6
            )
        for row_idx, slot in enumerate(SLOTS, start=1):
            ctk.CTkLabel(matrix, text=f"{slot} пара").grid(row=row_idx, column=0, padx=6, pady=4)
            for day in range(len(DAYS)):
                key = slot_key(day, slot)
                var = ctk.IntVar(value=1 if key in teacher.blocked_slots else 0)
                slot_vars[key] = var
                ctk.CTkCheckBox(matrix, text="", width=20, variable=var).grid(
                    row=row_idx, column=day + 1, padx=6, pady=2
                )

        def commit() -> None:
            teacher.work_days = [day for day, var in day_vars.items() if var.get() == 1]
            teacher.blocked_slots = [key for key, var in slot_vars.items() if var.get() == 1]
            self.populate_json_box()
            self._set_status(f"График ППС обновлен: {teacher.id}")
            window.destroy()

        ctk.CTkButton(window, text="Сохранить график", command=commit).grid(
            row=3, column=0, padx=14, pady=(0, 12), sticky="w"
        )

    def export_schedule_excel(self) -> None:
        if not self.app_state.schedule:
            messagebox.showinfo("Экспорт", "Сначала сгенерируйте расписание")
            return
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if not path:
            return
        disc_map = {d.id: d.name for d in self.app_state.disciplines}
        teacher_map = {t.id: t.name for t in self.app_state.teachers}
        room_map = {r.id: r.name for r in self.app_state.rooms}
        wb = Workbook()
        ws = wb.active
        ws.title = "schedule"
        ws.append(["assignment_id", "discipline", "teacher", "groups", "day", "slot", "room", "weeks"])
        for entry in self.app_state.schedule:
            ws.append(
                [
                    entry.assignment_id,
                    disc_map.get(entry.discipline_id, entry.discipline_id),
                    teacher_map.get(entry.teacher_id, entry.teacher_id),
                    ",".join(entry.group_ids),
                    DAY_LABELS_RU[entry.day],
                    entry.slot,
                    room_map.get(entry.room_id, entry.room_id),
                    f"{entry.start_week}-{entry.end_week}",
                ]
            )
        wb.save(Path(path))
        self._set_status(f"Расписание Excel экспортировано: {Path(path).name}")

    def refresh_analysis(self) -> None:
        path = Path("docs/competitor_analysis.md")
        content = path.read_text(encoding="utf-8") if path.exists() else "Файл анализа не найден."
        self.analysis_box.delete("1.0", "end")
        self.analysis_box.insert("1.0", content)

    def _set_status(self, text: str) -> None:
        self.status_label.configure(text=text)

    @staticmethod
    def _id_from_combo(value: str) -> str:
        if "|" in value:
            return value.split("|", 1)[0].strip()
        return value.strip()

    @staticmethod
    def _read_week(value: str) -> int:
        try:
            return max(1, int(value))
        except ValueError:
            return 1
