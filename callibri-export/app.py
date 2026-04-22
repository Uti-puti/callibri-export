"""
app.py — GUI-приложение Callibri Export на CustomTkinter.
Вся бизнес-логика — в core.py.
"""

import os
import sys
import queue
import threading
import calendar
from datetime import datetime, timedelta
from tkinter import filedialog

import customtkinter as ctk
from dotenv import load_dotenv

import core
import providers

# Ленивый импорт gsheets — может отсутствовать если пакеты не установлены
try:
    import gsheets as gs
    HAS_GSHEETS = True
except ImportError:
    gs = None
    HAS_GSHEETS = False


# ── Диалог выбора даты (календарь) ──────────────────────────────────────────

class DatePickerDialog(ctk.CTkToplevel):
    MONTH_NAMES = [
        "Январь", "Февраль", "Март", "Апрель", "Май", "Июнь",
        "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь",
    ]
    WEEKDAY_NAMES = ["Пн", "Вт", "Ср", "Чт", "Пт", "Сб", "Вс"]

    def __init__(self, parent, initial_date=None, on_select=None):
        super().__init__(parent)
        self.title("Выбор даты")
        self.resizable(False, False)
        self.transient(parent)
        self.on_select = on_select

        today = datetime.now()
        d = initial_date or today
        self._year = d.year
        self._month = d.month
        self._selected_day = d.day if (d.year == today.year and d.month == today.month) else None

        self._build_ui()
        self._render_month()

        self.update_idletasks()
        try:
            px = parent.winfo_rootx()
            py = parent.winfo_rooty()
            self.geometry(f"+{px + 100}+{py + 100}")
        except Exception:
            pass

        self.grab_set()
        self.focus()

    def _build_ui(self):
        header = ctk.CTkFrame(self, fg_color="transparent")
        header.pack(fill="x", padx=8, pady=(8, 4))

        ctk.CTkButton(header, text="‹", width=32, command=self._prev_month).pack(side="left")

        self.lbl_title = ctk.CTkLabel(header, text="", font=ctk.CTkFont(size=14, weight="bold"))
        self.lbl_title.pack(side="left", expand=True, fill="x", padx=8)

        ctk.CTkButton(header, text="›", width=32, command=self._next_month).pack(side="left")

        self.grid_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.grid_frame.pack(padx=8, pady=4)

        for col, name in enumerate(self.WEEKDAY_NAMES):
            color = ("red3", "red") if col >= 5 else ("gray30", "gray70")
            ctk.CTkLabel(
                self.grid_frame, text=name, width=36, height=24,
                text_color=color, font=ctk.CTkFont(size=11, weight="bold"),
            ).grid(row=0, column=col, padx=1, pady=1)

        footer = ctk.CTkFrame(self, fg_color="transparent")
        footer.pack(fill="x", padx=8, pady=(4, 8))

        ctk.CTkButton(
            footer, text="Сегодня", width=80, height=26,
            fg_color="transparent", border_width=1, text_color=("gray10", "gray90"),
            command=self._pick_today,
        ).pack(side="left")

        ctk.CTkButton(
            footer, text="Отмена", width=80, height=26,
            fg_color="transparent", border_width=1, text_color=("gray10", "gray90"),
            command=self.destroy,
        ).pack(side="right")

    def _render_month(self):
        self.lbl_title.configure(text=f"{self.MONTH_NAMES[self._month - 1]} {self._year}")

        for widget in self.grid_frame.grid_slaves():
            info = widget.grid_info()
            if int(info.get("row", 0)) > 0:
                widget.destroy()

        cal = calendar.Calendar(firstweekday=0)
        weeks = cal.monthdayscalendar(self._year, self._month)
        today = datetime.now().date()

        for row_idx, week in enumerate(weeks, start=1):
            for col_idx, day in enumerate(week):
                if day == 0:
                    continue
                is_today = (day == today.day and self._month == today.month and self._year == today.year)
                is_selected = (day == self._selected_day)

                if is_selected:
                    fg_color = ("gray25", "gray75")
                elif is_today:
                    fg_color = ("#1f6aa5", "#144870")
                else:
                    fg_color = "transparent"

                btn = ctk.CTkButton(
                    self.grid_frame, text=str(day), width=36, height=28,
                    fg_color=fg_color,
                    border_width=0 if fg_color != "transparent" else 1,
                    text_color=("gray10", "gray90"),
                    command=lambda d=day: self._pick_day(d),
                )
                btn.grid(row=row_idx, column=col_idx, padx=1, pady=1)

    def _prev_month(self):
        if self._month == 1:
            self._month = 12
            self._year -= 1
        else:
            self._month -= 1
        self._selected_day = None
        self._render_month()

    def _next_month(self):
        if self._month == 12:
            self._month = 1
            self._year += 1
        else:
            self._month += 1
        self._selected_day = None
        self._render_month()

    def _pick_day(self, day):
        try:
            picked = datetime(self._year, self._month, day)
        except ValueError:
            return
        if self.on_select:
            self.on_select(picked)
        self.destroy()

    def _pick_today(self):
        if self.on_select:
            self.on_select(datetime.now())
        self.destroy()


# ── Диалог настроек проекта (поля + фильтры) ────────────────────────────────

class ProjectSettingsDialog(ctk.CTkToplevel):
    """Диалог редактирования полей и фильтров одного проекта."""

    def __init__(self, parent, proj_conf, provider, creds, gsheet_credentials=None):
        super().__init__(parent)
        self.proj_conf = proj_conf
        self.provider = provider
        self.creds = creds or {}
        self.gsheet_credentials = gsheet_credentials  # путь к credentials.json
        self.result = None  # будет dict с обновлёнными настройками

        site_id = proj_conf.get("site_id")
        folder = proj_conf.get("folder", "")
        self.title(f"Настройки [{provider.LABEL}] — {folder} ({site_id})")
        self.geometry("680x650")
        self.minsize(580, 520)
        self.grab_set()

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)  # tabview растягивается, кнопки прижаты вниз

        # --- Вкладки ---
        self.tabview = ctk.CTkTabview(self)
        self.tabview.grid(row=0, column=0, sticky="nsew", padx=10, pady=(10, 0))
        self.tabview.grid_rowconfigure(0, weight=1)
        self.tabview.grid_columnconfigure(0, weight=1)

        tab_fields = self.tabview.add("Поля")
        tab_filters = self.tabview.add("Фильтры")
        tab_gsheet = self.tabview.add("Google Sheets")

        self._build_fields_tab(tab_fields)
        self._build_filters_tab(tab_filters)
        self._build_gsheet_tab(tab_gsheet)

        # --- Кнопки ---
        btn_frame = ctk.CTkFrame(self, fg_color="transparent")
        btn_frame.grid(row=1, column=0, sticky="ew", padx=10, pady=10)

        ctk.CTkButton(btn_frame, text="Сохранить", width=120, command=self._on_save).pack(side="left", padx=(0, 8))
        ctk.CTkButton(
            btn_frame, text="Отмена", width=100,
            fg_color="transparent", border_width=1, text_color=("gray10", "gray90"),
            command=self.destroy,
        ).pack(side="left")

    # ── Вкладка «Поля» ───────────────────────────────────────────────────

    def _build_fields_tab(self, tab):
        tab.grid_columnconfigure(0, weight=1)
        tab.grid_columnconfigure(2, weight=1)
        tab.grid_rowconfigure(1, weight=1)

        current_fields = list(self.proj_conf.get("fields") or self.provider.DEFAULT_COLUMNS)
        available = [f for f in self.provider.ALL_FIELDS if f not in current_fields]

        # Заголовки
        ctk.CTkLabel(tab, text="Доступные поля").grid(row=0, column=0, pady=(4, 2))
        ctk.CTkLabel(tab, text="Выбранные поля").grid(row=0, column=2, pady=(4, 2))

        # Списки
        self.lst_available = ctk.CTkTextbox(tab, width=220, font=ctk.CTkFont(size=12))
        self.lst_available.grid(row=1, column=0, sticky="nsew", padx=(4, 2), pady=4)
        self.lst_available.bind("<ButtonRelease-1>", lambda e: self._on_field_click(self.lst_available))

        self.lst_selected = ctk.CTkTextbox(tab, width=220, font=ctk.CTkFont(size=12))
        self.lst_selected.grid(row=1, column=2, sticky="nsew", padx=(2, 4), pady=4)
        self.lst_selected.bind("<ButtonRelease-1>", lambda e: self._on_field_click(self.lst_selected))

        # Кнопки перемещения
        mid_frame = ctk.CTkFrame(tab, fg_color="transparent")
        mid_frame.grid(row=1, column=1, padx=4)

        ctk.CTkButton(mid_frame, text="-->", width=36, command=self._move_right).pack(pady=2)
        ctk.CTkButton(mid_frame, text="<--", width=36, command=self._move_left).pack(pady=2)
        ctk.CTkButton(mid_frame, text="^", width=36, command=self._move_up).pack(pady=(12, 2))
        ctk.CTkButton(mid_frame, text="v", width=36, command=self._move_down).pack(pady=2)
        ctk.CTkButton(
            mid_frame, text="Сброс", width=52, font=ctk.CTkFont(size=11),
            fg_color="transparent", border_width=1, text_color=("gray10", "gray90"),
            command=self._reset_fields,
        ).pack(pady=(12, 2))

        # Описание выбранного поля (внизу)
        self.lbl_field_desc = ctk.CTkLabel(
            tab, text="Кликни на поле для описания",
            text_color="gray", font=ctk.CTkFont(size=12),
            wraplength=500, anchor="w", justify="left",
        )
        self.lbl_field_desc.grid(row=2, column=0, columnspan=3, sticky="ew", padx=8, pady=(2, 4))

        self._available_fields = available
        self._selected_fields = current_fields
        self._refresh_field_lists()

    def _field_display(self, field_name):
        """Форматируем поле для отображения в списке: 'name — описание'."""
        desc = self.provider.FIELD_DESCRIPTIONS.get(field_name, "")
        if desc:
            return f"{field_name}  —  {desc}"
        return field_name

    def _field_from_display(self, display_text):
        """Извлекаем имя поля из строки отображения."""
        display_text = display_text.strip()
        if not display_text:
            return None
        # Формат: "field_name  —  описание" или просто "field_name"
        parts = display_text.split("  —  ", 1)
        return parts[0].strip()

    def _refresh_field_lists(self):
        self.lst_available.configure(state="normal")
        self.lst_available.delete("1.0", "end")
        for f in self._available_fields:
            self.lst_available.insert("end", self._field_display(f) + "\n")
        self.lst_available.configure(state="normal")

        self.lst_selected.configure(state="normal")
        self.lst_selected.delete("1.0", "end")
        for f in self._selected_fields:
            self.lst_selected.insert("end", self._field_display(f) + "\n")
        self.lst_selected.configure(state="normal")

    def _get_selected_line(self, textbox):
        """Получить имя поля из текущей строки и номер строки."""
        try:
            idx = textbox.index("insert")
            line_num = int(idx.split(".")[0])
            line_text = textbox.get(f"{line_num}.0", f"{line_num}.end").strip()
            field_name = self._field_from_display(line_text)
            return line_num, field_name
        except Exception:
            return None, None

    def _on_field_click(self, textbox):
        """При клике на поле — показать полное описание внизу."""
        _, field = self._get_selected_line(textbox)
        if field:
            desc = self.provider.FIELD_DESCRIPTIONS.get(field, "Нет описания")
            self.lbl_field_desc.configure(
                text=f"{field}:  {desc}",
                text_color=("gray10", "gray90"),
            )

    def _move_right(self):
        _, field = self._get_selected_line(self.lst_available)
        if field and field in self._available_fields:
            self._available_fields.remove(field)
            self._selected_fields.append(field)
            self._refresh_field_lists()

    def _move_left(self):
        _, field = self._get_selected_line(self.lst_selected)
        if field and field in self._selected_fields:
            self._selected_fields.remove(field)
            self._available_fields.append(field)
            self._refresh_field_lists()

    def _move_up(self):
        line_num, field = self._get_selected_line(self.lst_selected)
        if field and field in self._selected_fields:
            idx = self._selected_fields.index(field)
            if idx > 0:
                self._selected_fields[idx], self._selected_fields[idx - 1] = (
                    self._selected_fields[idx - 1], self._selected_fields[idx],
                )
                self._refresh_field_lists()
                self.lst_selected.mark_set("insert", f"{line_num - 1}.0")

    def _move_down(self):
        line_num, field = self._get_selected_line(self.lst_selected)
        if field and field in self._selected_fields:
            idx = self._selected_fields.index(field)
            if idx < len(self._selected_fields) - 1:
                self._selected_fields[idx], self._selected_fields[idx + 1] = (
                    self._selected_fields[idx + 1], self._selected_fields[idx],
                )
                self._refresh_field_lists()
                self.lst_selected.mark_set("insert", f"{line_num + 1}.0")

    def _reset_fields(self):
        self._selected_fields = list(self.provider.DEFAULT_COLUMNS)
        self._available_fields = [f for f in self.provider.ALL_FIELDS if f not in self._selected_fields]
        self._refresh_field_lists()

    # ── Вкладка «Фильтры» ────────────────────────────────────────────────

    def _build_filters_tab(self, tab):
        tab.grid_columnconfigure(0, weight=1)
        tab.grid_rowconfigure(0, weight=1)

        # Внутренний скролл на случай, если контент не помещается
        scroll = ctk.CTkScrollableFrame(tab, fg_color="transparent")
        scroll.grid(row=0, column=0, sticky="nsew")
        scroll.grid_columnconfigure(0, weight=1)

        pad = {"padx": 10, "sticky": "w"}

        # --- Типы ---
        ctk.CTkLabel(scroll, text="Типы обращений", font=ctk.CTkFont(weight="bold")).grid(
            row=0, column=0, pady=(8, 4), **pad
        )

        current_types = self.proj_conf.get("types") or []
        type_labels = self.provider.TYPE_LABELS
        self._type_vars = {}
        row = 1
        for key, label in type_labels.items():
            var = ctk.IntVar(value=1 if (not current_types or key in current_types) else 0)
            ctk.CTkCheckBox(scroll, text=label, variable=var).grid(row=row, column=0, **pad, pady=1)
            self._type_vars[key] = var
            row += 1

        # --- Каналы ---
        ctk.CTkLabel(scroll, text="Каналы", font=ctk.CTkFont(weight="bold")).grid(
            row=row, column=0, pady=(12, 4), **pad
        )
        row += 1

        self._channels_frame = ctk.CTkScrollableFrame(scroll, height=120)
        self._channels_frame.grid(row=row, column=0, sticky="ew", padx=10, pady=2)
        row += 1

        self._channel_vars = {}
        # Флаг «полный список каналов загружен из API» — определяет,
        # означает ли «все чекбоксы отмечены» снятие фильтра или нет.
        # Без него сохранение без загрузки каналов затирает сохранённый фильтр.
        self._channels_loaded_from_api = False
        current_channels = self.proj_conf.get("channels") or []

        # Кнопка загрузки каналов из API
        self._btn_load_channels = ctk.CTkButton(
            scroll, text="Загрузить каналы из API", width=200,
            command=self._load_channels,
        )
        self._btn_load_channels.grid(row=row, column=0, **pad, pady=4)
        row += 1

        # Если каналы уже настроены — показываем их
        if current_channels:
            for ch in current_channels:
                var = ctk.IntVar(value=1)
                ctk.CTkCheckBox(self._channels_frame, text=ch, variable=var).pack(anchor="w", pady=1)
                self._channel_vars[ch] = var

        # --- Статусы ---
        ctk.CTkLabel(scroll, text="Статусы", font=ctk.CTkFont(weight="bold")).grid(
            row=row, column=0, pady=(12, 4), **pad
        )
        row += 1

        current_statuses = self.proj_conf.get("statuses") or []
        self.entry_statuses = ctk.CTkEntry(scroll, width=400, placeholder_text="Лид, Целевой (через запятую, пусто = все)")
        self.entry_statuses.grid(row=row, column=0, padx=10, sticky="ew", pady=2)
        if current_statuses:
            self.entry_statuses.insert(0, ", ".join(current_statuses))
        row += 1

        # --- Формат ---
        ctk.CTkLabel(scroll, text="Формат выгрузки", font=ctk.CTkFont(weight="bold")).grid(
            row=row, column=0, pady=(12, 4), **pad
        )
        row += 1

        current_format = self.proj_conf.get("format", "xlsx")
        self._format_var = ctk.StringVar(value=current_format)
        ctk.CTkSegmentedButton(
            scroll, values=["xlsx", "csv"], variable=self._format_var, width=200,
        ).grid(row=row, column=0, **pad, pady=2)
        row += 1

        # --- Split by channel ---
        self._split_var = ctk.IntVar(value=1 if self.proj_conf.get("split_by_channel", False) else 0)
        ctk.CTkCheckBox(scroll, text="Разделять файлы по каналам", variable=self._split_var).grid(
            row=row, column=0, **pad, pady=(12, 4)
        )

    def _load_channels(self):
        """Загрузить каналы из API в отдельном потоке."""
        self._btn_load_channels.configure(state="disabled", text="Загружаем...")
        site_id = self.proj_conf.get("site_id")
        current_channels = self.proj_conf.get("channels") or []

        def _fetch():
            try:
                channel_names, statuses = self.provider.get_channels_and_statuses(
                    site_id, self.creds
                )
                self.after(0, lambda: self._show_channels(channel_names, current_channels, statuses))
            except Exception as e:
                self.after(0, lambda: self._btn_load_channels.configure(
                    state="normal", text=f"Ошибка: {e}"
                ))

        threading.Thread(target=_fetch, daemon=True).start()

    def _show_channels(self, channel_names, current_channels, statuses):
        self._btn_load_channels.configure(state="normal", text="Загрузить каналы из API")
        self._channels_loaded_from_api = True

        # Обновляем чекбоксы каналов
        for w in self._channels_frame.winfo_children():
            w.destroy()
        self._channel_vars.clear()

        for ch in channel_names:
            var = ctk.IntVar(value=1 if (not current_channels or ch in current_channels) else 0)
            ctk.CTkCheckBox(self._channels_frame, text=ch, variable=var).pack(anchor="w", pady=1)
            self._channel_vars[ch] = var

        # Подсказка по статусам
        if statuses and not self.entry_statuses.get().strip():
            self.entry_statuses.configure(placeholder_text=", ".join(statuses))

    # ── Вкладка «Google Sheets» ────────────────────────────────────────────

    def _build_gsheet_tab(self, tab):
        tab.grid_columnconfigure(0, weight=1)
        tab.grid_rowconfigure(0, weight=1)

        scroll = ctk.CTkScrollableFrame(tab, fg_color="transparent")
        scroll.grid(row=0, column=0, sticky="nsew")
        scroll.grid_columnconfigure(0, weight=1)
        pad = {"padx": 10, "sticky": "w"}

        gsheet_conf = self.proj_conf.get("gsheet") or {}

        # --- Включение ---
        self._gsheet_enabled_var = ctk.IntVar(value=1 if gsheet_conf.get("enabled") else 0)
        ctk.CTkCheckBox(
            scroll, text="Отправлять в Google Sheets", variable=self._gsheet_enabled_var,
            font=ctk.CTkFont(weight="bold"),
        ).grid(row=0, column=0, pady=(12, 8), **pad)

        # --- Таблица ---
        ctk.CTkLabel(scroll, text="Таблица (URL или ID):", font=ctk.CTkFont(size=12)).grid(
            row=1, column=0, pady=(4, 2), **pad
        )

        spreadsheet_frame = ctk.CTkFrame(scroll, fg_color="transparent")
        spreadsheet_frame.grid(row=2, column=0, sticky="ew", padx=10, pady=2)
        spreadsheet_frame.grid_columnconfigure(0, weight=1)

        self._gsheet_spreadsheet_entry = ctk.CTkEntry(
            spreadsheet_frame, placeholder_text="URL или ID таблицы"
        )
        self._gsheet_spreadsheet_entry.grid(row=0, column=0, sticky="ew", padx=(0, 6))

        sid = gsheet_conf.get("spreadsheet_id", "")
        if sid:
            self._gsheet_spreadsheet_entry.insert(0, sid)

        self._btn_load_sheets = ctk.CTkButton(
            spreadsheet_frame, text="Загрузить листы", width=130,
            command=self._on_load_sheets,
        )
        self._btn_load_sheets.grid(row=0, column=1)

        # Название таблицы (отображение)
        self._lbl_spreadsheet_title = ctk.CTkLabel(
            scroll, text="", text_color="gray", font=ctk.CTkFont(size=11),
        )
        self._lbl_spreadsheet_title.grid(row=3, column=0, **pad, pady=(0, 4))

        # --- Лист ---
        ctk.CTkLabel(scroll, text="Лист (вкладка):", font=ctk.CTkFont(size=12)).grid(
            row=4, column=0, pady=(8, 2), **pad
        )

        sheet_frame = ctk.CTkFrame(scroll, fg_color="transparent")
        sheet_frame.grid(row=5, column=0, sticky="ew", padx=10, pady=2)
        sheet_frame.grid_columnconfigure(0, weight=1)

        self._gsheet_sheet_var = ctk.StringVar(value=gsheet_conf.get("sheet_name", ""))
        self._gsheet_sheet_menu = ctk.CTkOptionMenu(
            sheet_frame, variable=self._gsheet_sheet_var,
            values=["(загрузите листы)"], width=200,
        )
        self._gsheet_sheet_menu.grid(row=0, column=0, sticky="w", padx=(0, 6))

        # Если sheet_name уже задан — показать его
        if gsheet_conf.get("sheet_name"):
            self._gsheet_sheet_menu.configure(values=[gsheet_conf["sheet_name"]])
            self._gsheet_sheet_var.set(gsheet_conf["sheet_name"])

        self._btn_create_sheet = ctk.CTkButton(
            sheet_frame, text="Создать новый", width=120,
            command=self._on_create_sheet,
        )
        self._btn_create_sheet.grid(row=0, column=1)

        # --- Режим ---
        ctk.CTkLabel(scroll, text="Режим записи:", font=ctk.CTkFont(size=12)).grid(
            row=6, column=0, pady=(12, 4), **pad
        )

        current_mode = gsheet_conf.get("mode", "append")
        self._gsheet_mode_var = ctk.StringVar(value=current_mode)
        ctk.CTkSegmentedButton(
            scroll, values=["append", "replace"],
            variable=self._gsheet_mode_var, width=250,
        ).grid(row=7, column=0, **pad, pady=2)

        mode_desc = ctk.CTkLabel(
            scroll, text="append = дополнить данные  |  replace = заменить все данные",
            text_color="gray", font=ctk.CTkFont(size=11),
        )
        mode_desc.grid(row=8, column=0, **pad, pady=(0, 4))

        # --- Файловый экспорт ---
        file_export = self.proj_conf.get("file_export", True)
        self._file_export_var = ctk.IntVar(value=1 if file_export else 0)
        ctk.CTkCheckBox(
            scroll, text="Сохранять в файл (XLSX/CSV)",
            variable=self._file_export_var,
        ).grid(row=9, column=0, pady=(12, 4), **pad)

        file_hint = ctk.CTkLabel(
            scroll, text="Если выключено — данные пойдут только в Google Sheets",
            text_color="gray", font=ctk.CTkFont(size=11),
        )
        file_hint.grid(row=10, column=0, **pad, pady=(0, 4))

        # --- Подсказка ---
        if not HAS_GSHEETS:
            hint = ctk.CTkLabel(
                scroll,
                text="Пакеты gspread/google-auth не установлены.\npip install gspread google-auth",
                text_color="red", font=ctk.CTkFont(size=11),
            )
            hint.grid(row=11, column=0, **pad, pady=(12, 4))
        elif not self.gsheet_credentials:
            hint = ctk.CTkLabel(
                scroll,
                text="Укажите путь к credentials.json в главном окне (секция Google Sheets)",
                text_color="orange", font=ctk.CTkFont(size=11),
            )
            hint.grid(row=11, column=0, **pad, pady=(12, 4))

    def _on_load_sheets(self):
        """Загрузить список листов из таблицы."""
        if not HAS_GSHEETS or not self.gsheet_credentials:
            return

        url_or_id = self._gsheet_spreadsheet_entry.get().strip()
        if not url_or_id:
            return

        self._btn_load_sheets.configure(state="disabled", text="Загрузка...")

        def _fetch():
            try:
                spreadsheet_id = gs.parse_spreadsheet_id(url_or_id)
                client = gs.authorize(self.gsheet_credentials)
                title, sheet_names = gs.get_spreadsheet_info(client, spreadsheet_id)
                self.after(0, lambda: self._show_sheets(spreadsheet_id, title, sheet_names))
            except Exception as e:
                self.after(0, lambda: self._sheets_error(str(e)))

        threading.Thread(target=_fetch, daemon=True).start()

    def _show_sheets(self, spreadsheet_id, title, sheet_names):
        self._btn_load_sheets.configure(state="normal", text="Загрузить листы")
        self._lbl_spreadsheet_title.configure(text=f'Таблица: "{title}"')

        # Обновляем entry — ставим чистый ID
        self._gsheet_spreadsheet_entry.delete(0, "end")
        self._gsheet_spreadsheet_entry.insert(0, spreadsheet_id)

        # Обновляем выпадающий список листов
        if sheet_names:
            self._gsheet_sheet_menu.configure(values=sheet_names)
            # Если текущий лист есть в списке — оставляем, иначе первый
            if self._gsheet_sheet_var.get() not in sheet_names:
                self._gsheet_sheet_var.set(sheet_names[0])

    def _sheets_error(self, error_msg):
        self._btn_load_sheets.configure(state="normal", text="Загрузить листы")
        self._lbl_spreadsheet_title.configure(
            text=f"Ошибка: {error_msg}", text_color="red"
        )

    def _on_create_sheet(self):
        """Создать новый лист в таблице."""
        if not HAS_GSHEETS or not self.gsheet_credentials:
            return

        spreadsheet_id = self._gsheet_spreadsheet_entry.get().strip()
        if not spreadsheet_id:
            return

        dialog = ctk.CTkInputDialog(
            text="Имя нового листа:", title="Создать лист"
        )
        sheet_name = dialog.get_input()
        if not sheet_name or not sheet_name.strip():
            return

        sheet_name = sheet_name.strip()

        try:
            sid = gs.parse_spreadsheet_id(spreadsheet_id)
            client = gs.authorize(self.gsheet_credentials)
            gs.create_sheet(client, sid, sheet_name)
            # Обновляем список
            _, sheet_names = gs.get_spreadsheet_info(client, sid)
            self._gsheet_sheet_menu.configure(values=sheet_names)
            self._gsheet_sheet_var.set(sheet_name)
        except Exception as e:
            self._lbl_spreadsheet_title.configure(
                text=f"Ошибка создания листа: {e}", text_color="red"
            )

    # ── Сохранение ────────────────────────────────────────────────────────

    def _on_save(self):
        self.result = {}

        # Поля
        fields = [f for f in self._selected_fields if f]
        if fields and fields != list(self.provider.DEFAULT_COLUMNS):
            self.result["fields"] = fields
        else:
            self.result["fields"] = None  # удалить ключ = использовать default

        # Типы
        selected_types = [k for k, v in self._type_vars.items() if v.get()]
        if len(selected_types) < len(self._type_vars):
            self.result["types"] = selected_types
        else:
            self.result["types"] = None

        # Каналы
        if self._channels_loaded_from_api:
            # Полный список загружен из API — отметка о реальном выборе.
            # «Все отмечены» или «ни один» → снять фильтр, иначе — список.
            selected_ch = [ch for ch, v in self._channel_vars.items() if v.get()]
            total_ch = len(self._channel_vars)
            if selected_ch and len(selected_ch) < total_ch:
                self.result["channels"] = selected_ch
            else:
                self.result["channels"] = None
        elif self._channel_vars:
            # Показаны только ранее сохранённые каналы — чекбоксы управляют
            # лишь снятием из фильтра, не полной заменой.
            selected_ch = [ch for ch, v in self._channel_vars.items() if v.get()]
            self.result["channels"] = selected_ch if selected_ch else None
        else:
            # Каналы вообще не трогали
            self.result["channels"] = self.proj_conf.get("channels")

        # Статусы
        statuses_str = self.entry_statuses.get().strip()
        if statuses_str:
            self.result["statuses"] = [s.strip() for s in statuses_str.split(",") if s.strip()]
        else:
            self.result["statuses"] = None

        # Формат
        self.result["format"] = self._format_var.get()

        # Split
        self.result["split_by_channel"] = bool(self._split_var.get())

        # Файловый экспорт
        self.result["file_export"] = bool(self._file_export_var.get())

        # Google Sheets
        gsheet_enabled = bool(self._gsheet_enabled_var.get())
        gsheet_sid = self._gsheet_spreadsheet_entry.get().strip()
        gsheet_sheet = self._gsheet_sheet_var.get().strip()
        gsheet_mode = self._gsheet_mode_var.get()

        if gsheet_enabled and gsheet_sid and gsheet_sheet:
            try:
                gsheet_sid = gs.parse_spreadsheet_id(gsheet_sid) if HAS_GSHEETS else gsheet_sid
            except Exception:
                pass
            self.result["gsheet"] = {
                "enabled": True,
                "spreadsheet_id": gsheet_sid,
                "sheet_name": gsheet_sheet,
                "mode": gsheet_mode,
            }
        elif gsheet_sid or gsheet_sheet:
            # Настроено частично — сохраняем но выключено
            self.result["gsheet"] = {
                "enabled": False,
                "spreadsheet_id": gsheet_sid,
                "sheet_name": gsheet_sheet,
                "mode": gsheet_mode,
            }
        else:
            self.result["gsheet"] = None  # удалить ключ

        self.destroy()


# ── Диалог выбора провайдера ────────────────────────────────────────────────

class ProviderChoiceDialog(ctk.CTkToplevel):
    """Выбор провайдера для нового проекта."""

    def __init__(self, parent, provider_names):
        super().__init__(parent)
        self.result = None
        self.title("Выбор провайдера")
        self.geometry("320x200")
        self.grab_set()

        ctk.CTkLabel(
            self, text="Какой API использовать?",
            font=ctk.CTkFont(size=14, weight="bold"),
        ).pack(pady=(16, 8))

        self._var = ctk.StringVar(value=provider_names[0])
        for name in provider_names:
            prov = providers.get_provider(name)
            ctk.CTkRadioButton(
                self, text=prov.LABEL, variable=self._var, value=name,
            ).pack(anchor="w", padx=40, pady=2)

        btn_frame = ctk.CTkFrame(self, fg_color="transparent")
        btn_frame.pack(pady=16)
        ctk.CTkButton(btn_frame, text="OK", width=100, command=self._ok).pack(side="left", padx=6)
        ctk.CTkButton(
            btn_frame, text="Отмена", width=100,
            fg_color="transparent", border_width=1,
            command=self.destroy,
        ).pack(side="left", padx=6)

    def _ok(self):
        self.result = self._var.get()
        self.destroy()


# ── Диалог добавления проекта ────────────────────────────────────────────────

class AddProjectDialog(ctk.CTkToplevel):
    """Диалог добавления нового проекта из списка API."""

    def __init__(self, parent, sites, existing_keys, provider_name="callibri"):
        super().__init__(parent)
        self.result = None
        self.provider_name = provider_name
        provider = providers.get_provider(provider_name)
        self.title(f"Добавить проект [{provider.LABEL}]")
        self.geometry("500x400")
        self.minsize(400, 300)
        self.grab_set()

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        ctk.CTkLabel(self, text="Выберите проект из API", font=ctk.CTkFont(size=14, weight="bold")).grid(
            row=0, column=0, sticky="w", padx=12, pady=(12, 4)
        )

        # Список проектов
        self.scroll = ctk.CTkScrollableFrame(self)
        self.scroll.grid(row=1, column=0, sticky="nsew", padx=12, pady=4)
        self.scroll.grid_columnconfigure(0, weight=1)

        self._site_var = ctk.IntVar(value=0)
        self._sites_map = {}  # radio_value -> site dict

        idx = 1
        for site in sites:
            sid = site.get("site_id")
            name = site.get("sitename", "?")
            domains = site.get("domains", "")
            if isinstance(domains, list):
                domains = ", ".join(domains)

            already = " (уже добавлен)" if (provider_name, sid) in existing_keys else ""
            label = f"{name} ({sid}) — {domains}{already}"

            rb = ctk.CTkRadioButton(
                self.scroll, text=label, variable=self._site_var, value=idx,
            )
            rb.pack(anchor="w", pady=2, padx=4)
            self._sites_map[idx] = site
            idx += 1

        # Папка
        folder_frame = ctk.CTkFrame(self, fg_color="transparent")
        folder_frame.grid(row=2, column=0, sticky="ew", padx=12, pady=4)
        folder_frame.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(folder_frame, text="Папка:").pack(side="left", padx=(0, 6))
        self.entry_folder = ctk.CTkEntry(folder_frame, placeholder_text="имя папки для output/")
        self.entry_folder.pack(side="left", fill="x", expand=True)

        # Кнопки
        btn_frame = ctk.CTkFrame(self, fg_color="transparent")
        btn_frame.grid(row=3, column=0, sticky="ew", padx=12, pady=(4, 12))

        ctk.CTkButton(btn_frame, text="Добавить", width=120, command=self._on_add).pack(side="left", padx=(0, 8))
        ctk.CTkButton(
            btn_frame, text="Отмена", width=100,
            fg_color="transparent", border_width=1, text_color=("gray10", "gray90"),
            command=self.destroy,
        ).pack(side="left")

        # Автозаполнение folder при выборе
        self._site_var.trace_add("write", self._on_site_selected)

    def _on_site_selected(self, *_):
        val = self._site_var.get()
        site = self._sites_map.get(val)
        if site:
            name = site.get("sitename", "").lower()
            # Простая транслитерация для папки
            safe = core.sanitize_filename(name).replace(" ", "-")
            self.entry_folder.delete(0, "end")
            self.entry_folder.insert(0, safe)

    def _on_add(self):
        val = self._site_var.get()
        site = self._sites_map.get(val)
        if not site:
            return

        folder = self.entry_folder.get().strip()
        if not folder:
            return

        self.result = {
            "provider": self.provider_name,
            "site_id": site.get("site_id"),
            "folder": folder,
            "split_by_channel": False,
            "enabled": True,
        }
        self.destroy()


# ── Диалог ручного добавления проекта (для Calltouch и др.) ─────────────────

class ManualAddProjectDialog(ctk.CTkToplevel):
    """Ручной ввод siteId + folder для провайдеров без /sites API."""

    def __init__(self, parent, provider_name, existing_keys):
        super().__init__(parent)
        self.result = None
        self.provider_name = provider_name
        self.existing_keys = existing_keys

        provider = providers.get_provider(provider_name)
        self.title(f"Добавить проект [{provider.LABEL}]")
        self.geometry("500x260")
        self.grab_set()

        ctk.CTkLabel(
            self, text=f"Ручной ввод — {provider.LABEL}",
            font=ctk.CTkFont(size=14, weight="bold"),
        ).pack(pady=(14, 4))

        ctk.CTkLabel(
            self,
            text=(
                "Публичного списка сайтов у Calltouch нет.\n"
                "siteId и название можно найти в ЛК Calltouch → Интеграции → API."
            ),
            text_color="gray", font=ctk.CTkFont(size=11), justify="center",
        ).pack(pady=(0, 10))

        form = ctk.CTkFrame(self, fg_color="transparent")
        form.pack(fill="x", padx=20)
        form.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(form, text="siteId:").grid(row=0, column=0, sticky="w", pady=4)
        self.entry_site_id = ctk.CTkEntry(form, placeholder_text="например 67890")
        self.entry_site_id.grid(row=0, column=1, sticky="ew", padx=(8, 0), pady=4)

        ctk.CTkLabel(form, text="Название:").grid(row=1, column=0, sticky="w", pady=4)
        self.entry_name = ctk.CTkEntry(form, placeholder_text="для отображения в логе")
        self.entry_name.grid(row=1, column=1, sticky="ew", padx=(8, 0), pady=4)

        ctk.CTkLabel(form, text="Папка:").grid(row=2, column=0, sticky="w", pady=4)
        self.entry_folder = ctk.CTkEntry(form, placeholder_text="имя папки в output/")
        self.entry_folder.grid(row=2, column=1, sticky="ew", padx=(8, 0), pady=4)

        self.lbl_error = ctk.CTkLabel(self, text="", text_color="red")
        self.lbl_error.pack(pady=(4, 0))

        btn_frame = ctk.CTkFrame(self, fg_color="transparent")
        btn_frame.pack(pady=12)
        ctk.CTkButton(btn_frame, text="Добавить", width=120, command=self._on_add).pack(side="left", padx=6)
        ctk.CTkButton(
            btn_frame, text="Отмена", width=100,
            fg_color="transparent", border_width=1,
            command=self.destroy,
        ).pack(side="left", padx=6)

    def _on_add(self):
        sid_raw = self.entry_site_id.get().strip()
        name = self.entry_name.get().strip()
        folder = self.entry_folder.get().strip()

        if not sid_raw:
            self.lbl_error.configure(text="Укажи siteId")
            return
        try:
            site_id = int(sid_raw)
        except ValueError:
            self.lbl_error.configure(text="siteId должен быть числом")
            return
        if not folder:
            self.lbl_error.configure(text="Укажи папку")
            return

        key = (self.provider_name, site_id)
        if key in self.existing_keys:
            self.lbl_error.configure(text="Такой проект уже добавлен")
            return

        self.result = {
            "provider": self.provider_name,
            "site_id": site_id,
            "folder": folder,
            "name": name or folder,
            "split_by_channel": False,
            "enabled": True,
        }
        self.destroy()


# ── Главное окно ─────────────────────────────────────────────────────────────

class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Callibri Export")
        self.geometry("680x780")
        self.minsize(560, 420)

        ctk.set_appearance_mode("system")
        ctk.set_default_color_theme("blue")

        self.msg_queue = queue.Queue()
        self._projects_config = []
        self._project_widgets = []  # list of dicts per project
        self._api_sites = None  # кэш get_sites()
        self._gsheet_credentials_path = ""  # путь к credentials.json

        self._build_ui()
        self._load_env()
        self._load_projects()
        self._poll_queue()

    # ── UI ────────────────────────────────────────────────────────────────

    def _build_ui(self):
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)  # скролл-контейнер растягивается

        # Весь основной контент живёт в скроллируемом фрейме,
        # чтобы окно можно было уменьшать без потери доступа к полям.
        scroll = ctk.CTkScrollableFrame(self, fg_color="transparent")
        scroll.grid(row=0, column=0, sticky="nsew", padx=0, pady=0)
        scroll.grid_columnconfigure(0, weight=1)

        pad = {"padx": 12, "pady": (6, 0)}

        # --- Подключение ---
        frame_conn = ctk.CTkFrame(scroll)
        frame_conn.grid(row=0, column=0, sticky="ew", **pad)
        frame_conn.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(frame_conn, text="Подключение", font=ctk.CTkFont(size=14, weight="bold")).grid(
            row=0, column=0, columnspan=3, sticky="w", padx=10, pady=(8, 4)
        )

        ctk.CTkLabel(frame_conn, text="Callibri Email:").grid(row=1, column=0, sticky="w", padx=(10, 4), pady=2)
        self.entry_email = ctk.CTkEntry(frame_conn, width=300)
        self.entry_email.grid(row=1, column=1, sticky="ew", padx=4, pady=2)

        ctk.CTkLabel(frame_conn, text="Callibri Token:").grid(row=2, column=0, sticky="w", padx=(10, 4), pady=2)
        self.entry_token = ctk.CTkEntry(frame_conn, width=300, show="*")
        self.entry_token.grid(row=2, column=1, sticky="ew", padx=4, pady=2)

        self.btn_check = ctk.CTkButton(frame_conn, text="Проверить", width=120, command=self._on_check_connection)
        self.btn_check.grid(row=1, column=2, rowspan=2, padx=(4, 10), pady=2)

        ctk.CTkLabel(frame_conn, text="Calltouch API ID:").grid(row=3, column=0, sticky="w", padx=(10, 4), pady=2)
        self.entry_calltouch = ctk.CTkEntry(frame_conn, width=300, show="*", placeholder_text="clientApiId")
        self.entry_calltouch.grid(row=3, column=1, sticky="ew", padx=4, pady=2)

        self.btn_check_calltouch = ctk.CTkButton(
            frame_conn, text="Проверить CT", width=120,
            command=self._on_check_calltouch,
        )
        self.btn_check_calltouch.grid(row=3, column=2, padx=(4, 10), pady=2)

        self.lbl_conn_status = ctk.CTkLabel(frame_conn, text="", text_color="gray")
        self.lbl_conn_status.grid(row=4, column=0, columnspan=3, sticky="w", padx=10, pady=(0, 8))

        # --- Google Sheets ---
        frame_gsheet = ctk.CTkFrame(scroll)
        frame_gsheet.grid(row=1, column=0, sticky="ew", **pad)
        frame_gsheet.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(frame_gsheet, text="Google Sheets", font=ctk.CTkFont(size=14, weight="bold")).grid(
            row=0, column=0, columnspan=3, sticky="w", padx=10, pady=(8, 4)
        )

        ctk.CTkLabel(frame_gsheet, text="Credentials:").grid(row=1, column=0, sticky="w", padx=(10, 4), pady=2)
        self.entry_gsheet_creds = ctk.CTkEntry(frame_gsheet, width=300, placeholder_text="credentials.json")
        self.entry_gsheet_creds.grid(row=1, column=1, sticky="ew", padx=4, pady=2)

        btn_frame_gs = ctk.CTkFrame(frame_gsheet, fg_color="transparent")
        btn_frame_gs.grid(row=1, column=2, padx=(4, 10), pady=2)

        ctk.CTkButton(
            btn_frame_gs, text="...", width=36,
            command=self._on_browse_credentials,
        ).pack(side="left", padx=(0, 4))

        self.btn_check_gsheet = ctk.CTkButton(
            btn_frame_gs, text="Проверить", width=90,
            command=self._on_check_gsheet,
        )
        self.btn_check_gsheet.pack(side="left")

        self.lbl_gsheet_status = ctk.CTkLabel(frame_gsheet, text="Не настроено", text_color="gray")
        self.lbl_gsheet_status.grid(row=2, column=0, columnspan=3, sticky="w", padx=10, pady=(0, 8))

        # --- Период ---
        frame_period = ctk.CTkFrame(scroll)
        frame_period.grid(row=2, column=0, sticky="ew", **pad)

        ctk.CTkLabel(frame_period, text="Период", font=ctk.CTkFont(size=14, weight="bold")).grid(
            row=0, column=0, columnspan=6, sticky="w", padx=10, pady=(8, 4)
        )

        ctk.CTkLabel(frame_period, text="С:").grid(row=1, column=0, sticky="w", padx=(10, 4), pady=2)
        self.entry_date1 = ctk.CTkEntry(frame_period, width=120, placeholder_text="dd.mm.yyyy")
        self.entry_date1.grid(row=1, column=1, padx=4, pady=2)
        ctk.CTkButton(
            frame_period, text="📅", width=32, height=28,
            fg_color="transparent", border_width=1, text_color=("gray10", "gray90"),
            command=lambda: self._pick_date(self.entry_date1),
        ).grid(row=1, column=2, padx=(0, 8), pady=2)

        ctk.CTkLabel(frame_period, text="По:").grid(row=1, column=3, sticky="w", padx=(12, 4), pady=2)
        self.entry_date2 = ctk.CTkEntry(frame_period, width=120, placeholder_text="dd.mm.yyyy")
        self.entry_date2.grid(row=1, column=4, padx=4, pady=2)
        ctk.CTkButton(
            frame_period, text="📅", width=32, height=28,
            fg_color="transparent", border_width=1, text_color=("gray10", "gray90"),
            command=lambda: self._pick_date(self.entry_date2),
        ).grid(row=1, column=5, padx=(0, 8), pady=2)

        btn_frame = ctk.CTkFrame(frame_period, fg_color="transparent")
        btn_frame.grid(row=2, column=0, columnspan=6, sticky="w", padx=10, pady=(4, 8))

        for days, text in [(7, "7 дней"), (14, "14 дней"), (30, "30 дней")]:
            ctk.CTkButton(
                btn_frame, text=text, width=80,
                fg_color="transparent", border_width=1, text_color=("gray10", "gray90"),
                command=lambda d=days: self._set_quick_period(d),
            ).pack(side="left", padx=(0, 6))

        self._set_quick_period(7)

        # --- Проекты ---
        frame_projects = ctk.CTkFrame(scroll)
        frame_projects.grid(row=3, column=0, sticky="ew", **pad)
        frame_projects.grid_columnconfigure(0, weight=1)

        projects_header = ctk.CTkFrame(frame_projects, fg_color="transparent")
        projects_header.grid(row=0, column=0, sticky="ew", padx=10, pady=(8, 4))

        ctk.CTkLabel(projects_header, text="Проекты", font=ctk.CTkFont(size=14, weight="bold")).pack(side="left")

        self.btn_add_project = ctk.CTkButton(
            projects_header, text="+ Добавить", width=100, height=28,
            font=ctk.CTkFont(size=12),
            command=self._on_add_project,
        )
        self.btn_add_project.pack(side="right")

        self.projects_scroll = ctk.CTkScrollableFrame(frame_projects, height=130)
        self.projects_scroll.grid(row=1, column=0, sticky="ew", padx=10, pady=(0, 8))
        self.projects_scroll.grid_columnconfigure(0, weight=1)

        # --- Прогресс ---
        frame_progress = ctk.CTkFrame(scroll)
        frame_progress.grid(row=4, column=0, sticky="ew", **pad)
        frame_progress.grid_columnconfigure(0, weight=1)

        self.progress_bar = ctk.CTkProgressBar(frame_progress)
        self.progress_bar.grid(row=0, column=0, sticky="ew", padx=10, pady=(8, 2))
        self.progress_bar.set(0)

        self.lbl_progress = ctk.CTkLabel(frame_progress, text="Ожидание", text_color="gray")
        self.lbl_progress.grid(row=1, column=0, sticky="w", padx=10, pady=(0, 8))

        # --- Лог ---
        frame_log = ctk.CTkFrame(scroll)
        frame_log.grid(row=5, column=0, sticky="ew", **pad)
        frame_log.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(frame_log, text="Лог", font=ctk.CTkFont(size=14, weight="bold")).grid(
            row=0, column=0, sticky="w", padx=10, pady=(8, 4)
        )

        self.log_box = ctk.CTkTextbox(frame_log, height=180, font=ctk.CTkFont(family="Consolas", size=12))
        self.log_box.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0, 8))
        self.log_box.configure(state="disabled")

        # --- Кнопки (всегда видны, вне скролла) ---
        frame_buttons = ctk.CTkFrame(self, fg_color="transparent")
        frame_buttons.grid(row=1, column=0, sticky="ew", padx=12, pady=(6, 12))

        self.btn_export = ctk.CTkButton(
            frame_buttons, text="Экспорт", width=160, height=38,
            font=ctk.CTkFont(size=14, weight="bold"),
            command=self._on_export,
        )
        self.btn_export.pack(side="left")

        self.btn_open = ctk.CTkButton(
            frame_buttons, text="Открыть output", width=160, height=38,
            fg_color="transparent", border_width=1, text_color=("gray10", "gray90"),
            command=self._on_open_output,
        )
        self.btn_open.pack(side="right")

    # ── Загрузка данных ───────────────────────────────────────────────────

    def _load_env(self):
        env_path = os.path.join(core.get_app_dir(), ".env")
        if os.path.exists(env_path):
            load_dotenv(env_path, override=True)
        email = os.getenv("CALLIBRI_EMAIL", "")
        token = os.getenv("CALLIBRI_TOKEN", "")
        calltouch_id = os.getenv("CALLTOUCH_API_ID", "")
        gsheet_creds = os.getenv("GSHEET_CREDENTIALS", "")
        if email:
            self.entry_email.insert(0, email)
        if token:
            self.entry_token.insert(0, token)
        if calltouch_id:
            self.entry_calltouch.insert(0, calltouch_id)
        if gsheet_creds:
            self.entry_gsheet_creds.insert(0, gsheet_creds)
            self._gsheet_credentials_path = gsheet_creds

    def _save_env(self):
        env_path = os.path.join(core.get_app_dir(), ".env")
        with open(env_path, "w", encoding="utf-8") as f:
            f.write(f"CALLIBRI_EMAIL={self.entry_email.get()}\n")
            f.write(f"CALLIBRI_TOKEN={self.entry_token.get()}\n")
            ct_id = self.entry_calltouch.get().strip()
            if ct_id:
                f.write(f"CALLTOUCH_API_ID={ct_id}\n")
            gsheet_creds = self.entry_gsheet_creds.get().strip()
            if gsheet_creds:
                f.write(f"GSHEET_CREDENTIALS={gsheet_creds}\n")

    # ── Google Sheets подключение ────────────────────────────────────────

    def _on_browse_credentials(self):
        path = filedialog.askopenfilename(
            title="Выберите credentials.json",
            filetypes=[("JSON", "*.json"), ("All", "*.*")],
        )
        if path:
            self.entry_gsheet_creds.delete(0, "end")
            self.entry_gsheet_creds.insert(0, path)
            self._gsheet_credentials_path = path

    def _on_check_gsheet(self):
        creds_path = self.entry_gsheet_creds.get().strip()
        if not creds_path:
            self.lbl_gsheet_status.configure(text="Укажите путь к credentials.json", text_color="orange")
            return

        if not HAS_GSHEETS:
            self.lbl_gsheet_status.configure(
                text="Пакеты не установлены: pip install gspread google-auth",
                text_color="red",
            )
            return

        self.btn_check_gsheet.configure(state="disabled", text="Проверяем...")
        self._gsheet_credentials_path = creds_path

        def _check():
            ok, sa_email, msg = gs.test_gsheet_connection(creds_path)
            self.msg_queue.put(("gsheet_conn_result", ok, msg))

        threading.Thread(target=_check, daemon=True).start()

    def _load_projects(self):
        for widget in self.projects_scroll.winfo_children():
            widget.destroy()
        self._project_widgets.clear()
        self._projects_config.clear()

        try:
            self._projects_config = core.load_projects()
        except (FileNotFoundError, ValueError) as e:
            self._append_log(f"Проекты: {e}")
            return

        for i, proj in enumerate(self._projects_config):
            self._add_project_row(i, proj)

    def _add_project_row(self, idx, proj):
        """Создаёт одну строку проекта в списке."""
        provider_name = proj.get("provider", "callibri")
        site_id = proj.get("site_id")
        folder = proj.get("folder", "")
        enabled = proj.get("enabled", True)
        channels = proj.get("channels")
        fmt = proj.get("format", "xlsx")
        fields = proj.get("fields")

        row_frame = ctk.CTkFrame(self.projects_scroll, fg_color="transparent")
        row_frame.pack(fill="x", pady=1)
        row_frame.grid_columnconfigure(1, weight=1)

        gsheet = proj.get("gsheet")

        var = ctk.IntVar(value=1 if enabled else 0)
        label = f"[{provider_name}] {folder} ({site_id})"
        if channels:
            label += f" — {', '.join(channels)}"
        label += f" | {fmt}"
        if fields:
            label += f" | {len(fields)} полей"
        if not proj.get("file_export", True):
            label += " | без файла"
        if gsheet and gsheet.get("enabled"):
            label += " | GSheets"

        cb = ctk.CTkCheckBox(row_frame, text=label, variable=var)
        cb.grid(row=0, column=0, columnspan=2, sticky="w", padx=4)

        btn_settings = ctk.CTkButton(
            row_frame, text="Настроить", width=80, height=24,
            font=ctk.CTkFont(size=11),
            command=lambda i=idx: self._on_project_settings(i),
        )
        btn_settings.grid(row=0, column=2, padx=(4, 2))

        btn_remove = ctk.CTkButton(
            row_frame, text="✕", width=30, height=24,
            font=ctk.CTkFont(size=11),
            fg_color="transparent", border_width=1,
            text_color=("red", "#ff6666"), hover_color=("gray90", "gray25"),
            command=lambda i=idx: self._on_remove_project(i),
        )
        btn_remove.grid(row=0, column=3, padx=(2, 4))

        self._project_widgets.append({
            "provider": provider_name,
            "site_id": site_id,
            "var": var,
            "frame": row_frame,
        })

    # ── Настройки проекта ─────────────────────────────────────────────────

    def _on_project_settings(self, idx):
        if idx >= len(self._projects_config):
            return
        proj = self._projects_config[idx]
        try:
            provider = core.get_project_provider(proj)
        except ValueError as e:
            self._append_log(f"ОШИБКА: {e}")
            return
        creds = self._get_creds(provider)
        gsheet_creds = self._gsheet_credentials_path

        dialog = ProjectSettingsDialog(
            self, proj, provider, creds, gsheet_credentials=gsheet_creds,
        )
        self.wait_window(dialog)

        if dialog.result is not None:
            # Применяем изменения
            r = dialog.result
            for key in ("fields", "types", "channels", "statuses"):
                if r[key] is not None:
                    proj[key] = r[key]
                elif key in proj:
                    del proj[key]

            proj["format"] = r["format"]
            proj["split_by_channel"] = r["split_by_channel"]
            proj["file_export"] = r.get("file_export", True)

            # Google Sheets
            if r.get("gsheet") is not None:
                proj["gsheet"] = r["gsheet"]
            elif "gsheet" in proj:
                del proj["gsheet"]

            # Сохраняем и обновляем UI
            core.save_projects(self._projects_config)
            self._load_projects()
            self._append_log(f"Настройки проекта {proj.get('folder')} обновлены")

    # ── Удаление проекта ──────────────────────────────────────────────────

    def _on_remove_project(self, idx):
        if idx >= len(self._projects_config):
            return
        proj = self._projects_config[idx]
        folder = proj.get("folder", "?")
        self._projects_config.pop(idx)
        core.save_projects(self._projects_config)
        self._load_projects()
        self._append_log(f"Проект {folder} удалён")

    # ── Добавление проекта ────────────────────────────────────────────────

    def _on_add_project(self):
        # Запрос провайдера у пользователя (если их > 1)
        avail = providers.provider_names()
        if len(avail) == 1:
            provider_name = avail[0]
        else:
            dialog = ProviderChoiceDialog(self, avail)
            self.wait_window(dialog)
            if not dialog.result:
                return
            provider_name = dialog.result

        provider = providers.get_provider(provider_name)
        creds = self._get_creds(provider)
        ok, msg = provider.check_credentials(creds)
        if not ok:
            self._append_log(f"Заполни учётные данные {provider.LABEL}: {msg}")
            return

        # Если у провайдера нет публичного /sites API — диалог ручного ввода
        if getattr(provider, "REQUIRES_MANUAL_SITE_ID", False):
            self._show_manual_add_dialog(provider.NAME)
            return

        self._append_log(f"Загружаем проекты из API [{provider.LABEL}]...")
        self.btn_add_project.configure(state="disabled")

        def _fetch():
            try:
                sites = provider.list_sites(creds)
                self.msg_queue.put(("sites_loaded", provider.NAME, sites))
            except Exception as e:
                self.msg_queue.put(("log", f"Ошибка загрузки проектов [{provider.LABEL}]: {e}"))
                self.msg_queue.put(("sites_loaded", provider.NAME, None))

        threading.Thread(target=_fetch, daemon=True).start()

    def _show_manual_add_dialog(self, provider_name):
        """Диалог ручного ввода siteId (для провайдеров без /sites API)."""
        existing = {
            (p.get("provider", "callibri"), p.get("site_id"))
            for p in self._projects_config
        }
        dialog = ManualAddProjectDialog(self, provider_name, existing)
        self.wait_window(dialog)

        if dialog.result is not None:
            key = (dialog.result["provider"], dialog.result["site_id"])
            if key in existing:
                self._append_log(f"Проект [{key[0]}] site_id={key[1]} уже есть в конфиге")
                return
            self._projects_config.append(dialog.result)
            core.save_projects(self._projects_config)
            self._load_projects()
            self._append_log(
                f"Добавлен проект [{key[0]}]: {dialog.result['folder']} ({key[1]})"
            )

    def _show_add_dialog(self, sites, provider_name="callibri"):
        existing = {
            (p.get("provider", "callibri"), p.get("site_id"))
            for p in self._projects_config
        }
        dialog = AddProjectDialog(self, sites, existing, provider_name)
        self.wait_window(dialog)

        if dialog.result is not None:
            new_sid = dialog.result["site_id"]
            key = (dialog.result.get("provider", "callibri"), new_sid)
            if key in existing:
                self._append_log(f"Проект [{key[0]}] site_id={new_sid} уже есть в конфиге")
                return

            self._projects_config.append(dialog.result)
            core.save_projects(self._projects_config)
            self._load_projects()
            self._append_log(f"Добавлен проект [{key[0]}]: {dialog.result['folder']} ({new_sid})")

    # ── Быстрый выбор периода ─────────────────────────────────────────────

    def _set_quick_period(self, days):
        date2 = datetime.now()
        date1 = date2 - timedelta(days=days - 1)
        self.entry_date1.delete(0, "end")
        self.entry_date1.insert(0, date1.strftime("%d.%m.%Y"))
        self.entry_date2.delete(0, "end")
        self.entry_date2.insert(0, date2.strftime("%d.%m.%Y"))

    def _pick_date(self, entry):
        current_value = entry.get().strip()
        initial = None
        if current_value:
            try:
                initial = datetime.strptime(current_value, "%d.%m.%Y")
            except ValueError:
                initial = None

        def _on_select(picked):
            entry.delete(0, "end")
            entry.insert(0, picked.strftime("%d.%m.%Y"))

        DatePickerDialog(self, initial_date=initial, on_select=_on_select)

    # ── Проверка соединения ───────────────────────────────────────────────

    def _on_check_connection(self):
        self.btn_check.configure(state="disabled", text="Проверяем...")
        self.lbl_conn_status.configure(text="", text_color="gray")
        self._api_sites = None  # сбрасываем кэш
        self._save_env()

        provider = providers.get_provider("callibri")
        creds = self._get_creds(provider)

        def _check():
            ok, count, msg = provider.test_connection(creds)
            self.msg_queue.put(("conn_result", ok, msg))

        threading.Thread(target=_check, daemon=True).start()

    def _on_check_calltouch(self):
        """Проверка подключения к Calltouch API."""
        self.btn_check_calltouch.configure(state="disabled", text="Проверяем...")
        self._api_sites = None
        # Сразу сохраняем токен в .env, чтобы он не терялся между запусками
        self._save_env()

        provider = providers.get_provider("calltouch")
        creds = self._get_creds(provider)

        def _check():
            ok, count, msg = provider.test_connection(creds)
            self.msg_queue.put(("conn_result", ok, f"[Calltouch] {msg}"))

        threading.Thread(target=_check, daemon=True).start()

    # ── Экспорт ───────────────────────────────────────────────────────────

    def _on_export(self):
        date1_str = self.entry_date1.get().strip()
        date2_str = self.entry_date2.get().strip()
        try:
            core.parse_date(date1_str, "С")
            core.parse_date(date2_str, "По")
        except ValueError as e:
            self._append_log(f"ОШИБКА: {e}")
            return

        enabled_keys = set()
        for pw in self._project_widgets:
            if pw["var"].get():
                enabled_keys.add((pw["provider"], pw["site_id"]))

        if not enabled_keys:
            self._append_log("Нет выбранных проектов для экспорта")
            return

        self._save_env()

        self.btn_export.configure(state="disabled", text="Экспорт...")
        self.progress_bar.set(0)
        self.lbl_progress.configure(text="Запуск...", text_color="gray")
        self.log_box.configure(state="normal")
        self.log_box.delete("1.0", "end")
        self.log_box.configure(state="disabled")
        self._write_log_file(f"{'=' * 60}\n=== Старт экспорта {date1_str} — {date2_str} ===")

        gsheet_creds = self._gsheet_credentials_path if self._gsheet_credentials_path else None
        credentials = self._all_credentials()

        params = dict(
            credentials=credentials,
            date1_str=date1_str,
            date2_str=date2_str,
            enabled_keys=enabled_keys,
            on_log=lambda m: self.msg_queue.put(("log", m)),
            on_progress=lambda pi, pt, ci, ct: self.msg_queue.put(("progress", pi, pt, ci, ct)),
            gsheet_credentials=gsheet_creds,
        )

        threading.Thread(target=self._run_export, args=(params,), daemon=True).start()

    # ── Учётные данные ────────────────────────────────────────────────────

    def _get_creds(self, provider):
        """Собрать creds dict для конкретного провайдера из UI."""
        if provider.NAME == "callibri":
            return {
                "email": self.entry_email.get().strip(),
                "token": self.entry_token.get().strip(),
            }
        if provider.NAME == "calltouch":
            return {
                "client_api_id": self.entry_calltouch.get().strip(),
            }
        return {}

    def _all_credentials(self):
        """Полный dict учётных данных для run_export."""
        creds = {}
        for prov in providers.all_providers():
            creds[prov.NAME] = self._get_creds(prov)
        return creds

    def _run_export(self, params):
        try:
            result = core.run_export(**params)
            self.msg_queue.put(("complete", result))
        except Exception as e:
            self.msg_queue.put(("log", f"ОШИБКА: {e}"))
            self.msg_queue.put(("complete", None))

    # ── Открыть output ────────────────────────────────────────────────────

    def _on_open_output(self):
        output_dir = os.path.join(core.get_app_dir(), "output")
        os.makedirs(output_dir, exist_ok=True)
        os.startfile(output_dir)

    # ── Очередь сообщений ─────────────────────────────────────────────────

    def _poll_queue(self):
        while True:
            try:
                msg = self.msg_queue.get_nowait()
            except queue.Empty:
                break

            kind = msg[0]

            if kind == "log":
                ts = datetime.now().strftime("%H:%M:%S")
                self._append_log(f"{ts}  {msg[1]}")

            elif kind == "progress":
                _, pi, pt, ci, ct = msg
                if pt > 0 and ct > 0:
                    frac = (pi + ci / ct) / pt
                    self.progress_bar.set(min(frac, 1.0))
                    self.lbl_progress.configure(
                        text=f"Проект {pi + 1}/{pt}, чанк {ci}/{ct}",
                        text_color="gray",
                    )

            elif kind == "complete":
                result = msg[1]
                self.btn_export.configure(state="normal", text="Экспорт")
                if result:
                    self.progress_bar.set(1.0)
                    total_rows = sum(c for _, c in result["report"])
                    self.lbl_progress.configure(
                        text=f"Готово! {result['processed']} проект(ов), {total_rows} строк",
                        text_color="green",
                    )
                else:
                    self.lbl_progress.configure(text="Ошибка", text_color="red")

            elif kind == "conn_result":
                _, ok, message = msg
                self.btn_check.configure(state="normal", text="Проверить")
                self.btn_check_calltouch.configure(state="normal", text="Проверить CT")
                self.lbl_conn_status.configure(
                    text=message,
                    text_color="green" if ok else "red",
                )

            elif kind == "gsheet_conn_result":
                _, ok, message = msg
                self.btn_check_gsheet.configure(state="normal", text="Проверить")
                self.lbl_gsheet_status.configure(
                    text=message,
                    text_color="green" if ok else "red",
                )

            elif kind == "sites_loaded":
                _, provider_name, sites = msg
                self.btn_add_project.configure(state="normal")
                if sites is not None:
                    self._api_sites = sites
                    self._show_add_dialog(sites, provider_name)

        self.after(100, self._poll_queue)

    # ── Лог ───────────────────────────────────────────────────────────────

    def _append_log(self, text):
        self.log_box.configure(state="normal")
        self.log_box.insert("end", text + "\n")
        self.log_box.see("end")
        self.log_box.configure(state="disabled")
        self._write_log_file(text)

    def _write_log_file(self, text):
        try:
            log_path = os.path.join(core.get_app_dir(), "output", "export.log")
            os.makedirs(os.path.dirname(log_path), exist_ok=True)
            stamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            with open(log_path, "a", encoding="utf-8") as f:
                f.write(f"{stamp}  {text}\n")
        except Exception:
            pass


if __name__ == "__main__":
    app = App()
    app.mainloop()
