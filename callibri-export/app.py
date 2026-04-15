"""
app.py — GUI-приложение Callibri Export на CustomTkinter.
Вся бизнес-логика — в core.py.
"""

import os
import sys
import queue
import threading
from datetime import datetime, timedelta

import customtkinter as ctk
from dotenv import load_dotenv

import core


# ── Диалог настроек проекта (поля + фильтры) ────────────────────────────────

class ProjectSettingsDialog(ctk.CTkToplevel):
    """Диалог редактирования полей и фильтров одного проекта."""

    def __init__(self, parent, proj_conf, email, token):
        super().__init__(parent)
        self.proj_conf = proj_conf
        self.email = email
        self.token = token
        self.result = None  # будет dict с обновлёнными настройками

        site_id = proj_conf.get("site_id")
        folder = proj_conf.get("folder", "")
        self.title(f"Настройки — {folder} ({site_id})")
        self.geometry("680x650")
        self.minsize(580, 520)
        self.grab_set()

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        # --- Вкладки ---
        self.tabview = ctk.CTkTabview(self)
        self.tabview.grid(row=0, column=0, sticky="nsew", padx=10, pady=(10, 0))
        self.tabview.grid_rowconfigure(0, weight=1)

        tab_fields = self.tabview.add("Поля")
        tab_filters = self.tabview.add("Фильтры")

        self._build_fields_tab(tab_fields)
        self._build_filters_tab(tab_filters)

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

        current_fields = list(self.proj_conf.get("fields") or core.DEFAULT_COLUMNS)
        available = [f for f in core.ALL_FIELDS if f not in current_fields]

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
        desc = core.FIELD_DESCRIPTIONS.get(field_name, "")
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
            desc = core.FIELD_DESCRIPTIONS.get(field, "Нет описания")
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
        self._selected_fields = list(core.DEFAULT_COLUMNS)
        self._available_fields = [f for f in core.ALL_FIELDS if f not in self._selected_fields]
        self._refresh_field_lists()

    # ── Вкладка «Фильтры» ────────────────────────────────────────────────

    def _build_filters_tab(self, tab):
        tab.grid_columnconfigure(0, weight=1)

        pad = {"padx": 10, "sticky": "w"}

        # --- Типы ---
        ctk.CTkLabel(tab, text="Типы обращений", font=ctk.CTkFont(weight="bold")).grid(
            row=0, column=0, pady=(8, 4), **pad
        )

        current_types = self.proj_conf.get("types") or []
        type_labels = {
            "calls": "Звонки",
            "feedbacks": "Заявки",
            "chats": "Чаты",
            "emails": "Email",
        }
        self._type_vars = {}
        row = 1
        for key, label in type_labels.items():
            var = ctk.IntVar(value=1 if (not current_types or key in current_types) else 0)
            ctk.CTkCheckBox(tab, text=label, variable=var).grid(row=row, column=0, **pad, pady=1)
            self._type_vars[key] = var
            row += 1

        # --- Каналы ---
        ctk.CTkLabel(tab, text="Каналы", font=ctk.CTkFont(weight="bold")).grid(
            row=row, column=0, pady=(12, 4), **pad
        )
        row += 1

        self._channels_frame = ctk.CTkScrollableFrame(tab, height=100)
        self._channels_frame.grid(row=row, column=0, sticky="ew", padx=10, pady=2)
        row += 1

        self._channel_vars = {}
        current_channels = self.proj_conf.get("channels") or []

        # Кнопка загрузки каналов из API
        self._btn_load_channels = ctk.CTkButton(
            tab, text="Загрузить каналы из API", width=200,
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
        ctk.CTkLabel(tab, text="Статусы", font=ctk.CTkFont(weight="bold")).grid(
            row=row, column=0, pady=(12, 4), **pad
        )
        row += 1

        current_statuses = self.proj_conf.get("statuses") or []
        self.entry_statuses = ctk.CTkEntry(tab, width=400, placeholder_text="Лид, Целевой (через запятую, пусто = все)")
        self.entry_statuses.grid(row=row, column=0, padx=10, sticky="ew", pady=2)
        if current_statuses:
            self.entry_statuses.insert(0, ", ".join(current_statuses))
        row += 1

        # --- Формат ---
        ctk.CTkLabel(tab, text="Формат выгрузки", font=ctk.CTkFont(weight="bold")).grid(
            row=row, column=0, pady=(12, 4), **pad
        )
        row += 1

        current_format = self.proj_conf.get("format", "xlsx")
        self._format_var = ctk.StringVar(value=current_format)
        ctk.CTkSegmentedButton(
            tab, values=["xlsx", "csv"], variable=self._format_var, width=200,
        ).grid(row=row, column=0, **pad, pady=2)
        row += 1

        # --- Split by channel ---
        self._split_var = ctk.IntVar(value=1 if self.proj_conf.get("split_by_channel", False) else 0)
        ctk.CTkCheckBox(tab, text="Разделять файлы по каналам", variable=self._split_var).grid(
            row=row, column=0, **pad, pady=(12, 4)
        )

    def _load_channels(self):
        """Загрузить каналы из API в отдельном потоке."""
        self._btn_load_channels.configure(state="disabled", text="Загружаем...")
        site_id = self.proj_conf.get("site_id")
        current_channels = self.proj_conf.get("channels") or []

        def _fetch():
            try:
                channel_names, statuses = core.get_channels_and_statuses(
                    site_id, self.email, self.token
                )
                self.after(0, lambda: self._show_channels(channel_names, current_channels, statuses))
            except Exception as e:
                self.after(0, lambda: self._btn_load_channels.configure(
                    state="normal", text=f"Ошибка: {e}"
                ))

        threading.Thread(target=_fetch, daemon=True).start()

    def _show_channels(self, channel_names, current_channels, statuses):
        self._btn_load_channels.configure(state="normal", text="Загрузить каналы из API")

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

    # ── Сохранение ────────────────────────────────────────────────────────

    def _on_save(self):
        self.result = {}

        # Поля
        fields = [f for f in self._selected_fields if f]
        if fields and fields != list(core.DEFAULT_COLUMNS):
            self.result["fields"] = fields
        else:
            self.result["fields"] = None  # удалить ключ = использовать default

        # Типы
        selected_types = [k for k, v in self._type_vars.items() if v.get()]
        if len(selected_types) < 4:
            self.result["types"] = selected_types
        else:
            self.result["types"] = None

        # Каналы
        if self._channel_vars:
            selected_ch = [ch for ch, v in self._channel_vars.items() if v.get()]
            total_ch = len(self._channel_vars)
            if selected_ch and len(selected_ch) < total_ch:
                self.result["channels"] = selected_ch
            else:
                self.result["channels"] = None
        else:
            # Не трогаем если каналы не загружались
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

        self.destroy()


# ── Диалог добавления проекта ────────────────────────────────────────────────

class AddProjectDialog(ctk.CTkToplevel):
    """Диалог добавления нового проекта из списка API."""

    def __init__(self, parent, sites, existing_site_ids):
        super().__init__(parent)
        self.result = None
        self.title("Добавить проект")
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

            already = " (уже добавлен)" if sid in existing_site_ids else ""
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
            "site_id": site.get("site_id"),
            "folder": folder,
            "split_by_channel": False,
            "enabled": True,
        }
        self.destroy()


# ── Главное окно ─────────────────────────────────────────────────────────────

class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Callibri Export")
        self.geometry("660x780")
        self.minsize(560, 620)

        ctk.set_appearance_mode("system")
        ctk.set_default_color_theme("blue")

        self.msg_queue = queue.Queue()
        self._projects_config = []
        self._project_widgets = []  # list of dicts per project
        self._api_sites = None  # кэш get_sites()

        self._build_ui()
        self._load_env()
        self._load_projects()
        self._poll_queue()

    # ── UI ────────────────────────────────────────────────────────────────

    def _build_ui(self):
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(4, weight=1)  # лог растягивается

        pad = {"padx": 12, "pady": (6, 0)}

        # --- Подключение ---
        frame_conn = ctk.CTkFrame(self)
        frame_conn.grid(row=0, column=0, sticky="ew", **pad)
        frame_conn.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(frame_conn, text="Подключение", font=ctk.CTkFont(size=14, weight="bold")).grid(
            row=0, column=0, columnspan=3, sticky="w", padx=10, pady=(8, 4)
        )

        ctk.CTkLabel(frame_conn, text="Email:").grid(row=1, column=0, sticky="w", padx=(10, 4), pady=2)
        self.entry_email = ctk.CTkEntry(frame_conn, width=300)
        self.entry_email.grid(row=1, column=1, sticky="ew", padx=4, pady=2)

        ctk.CTkLabel(frame_conn, text="Token:").grid(row=2, column=0, sticky="w", padx=(10, 4), pady=2)
        self.entry_token = ctk.CTkEntry(frame_conn, width=300, show="*")
        self.entry_token.grid(row=2, column=1, sticky="ew", padx=4, pady=2)

        self.btn_check = ctk.CTkButton(frame_conn, text="Проверить", width=120, command=self._on_check_connection)
        self.btn_check.grid(row=1, column=2, rowspan=2, padx=(4, 10), pady=2)

        self.lbl_conn_status = ctk.CTkLabel(frame_conn, text="", text_color="gray")
        self.lbl_conn_status.grid(row=3, column=0, columnspan=3, sticky="w", padx=10, pady=(0, 8))

        # --- Период ---
        frame_period = ctk.CTkFrame(self)
        frame_period.grid(row=1, column=0, sticky="ew", **pad)

        ctk.CTkLabel(frame_period, text="Период", font=ctk.CTkFont(size=14, weight="bold")).grid(
            row=0, column=0, columnspan=6, sticky="w", padx=10, pady=(8, 4)
        )

        ctk.CTkLabel(frame_period, text="С:").grid(row=1, column=0, sticky="w", padx=(10, 4), pady=2)
        self.entry_date1 = ctk.CTkEntry(frame_period, width=120, placeholder_text="dd.mm.yyyy")
        self.entry_date1.grid(row=1, column=1, padx=4, pady=2)

        ctk.CTkLabel(frame_period, text="По:").grid(row=1, column=2, sticky="w", padx=(12, 4), pady=2)
        self.entry_date2 = ctk.CTkEntry(frame_period, width=120, placeholder_text="dd.mm.yyyy")
        self.entry_date2.grid(row=1, column=3, padx=4, pady=2)

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
        frame_projects = ctk.CTkFrame(self)
        frame_projects.grid(row=2, column=0, sticky="ew", **pad)
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
        frame_progress = ctk.CTkFrame(self)
        frame_progress.grid(row=3, column=0, sticky="ew", **pad)
        frame_progress.grid_columnconfigure(0, weight=1)

        self.progress_bar = ctk.CTkProgressBar(frame_progress)
        self.progress_bar.grid(row=0, column=0, sticky="ew", padx=10, pady=(8, 2))
        self.progress_bar.set(0)

        self.lbl_progress = ctk.CTkLabel(frame_progress, text="Ожидание", text_color="gray")
        self.lbl_progress.grid(row=1, column=0, sticky="w", padx=10, pady=(0, 8))

        # --- Лог ---
        frame_log = ctk.CTkFrame(self)
        frame_log.grid(row=4, column=0, sticky="nsew", **pad)
        frame_log.grid_columnconfigure(0, weight=1)
        frame_log.grid_rowconfigure(1, weight=1)

        ctk.CTkLabel(frame_log, text="Лог", font=ctk.CTkFont(size=14, weight="bold")).grid(
            row=0, column=0, sticky="w", padx=10, pady=(8, 4)
        )

        self.log_box = ctk.CTkTextbox(frame_log, height=180, font=ctk.CTkFont(family="Consolas", size=12))
        self.log_box.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0, 8))
        self.log_box.configure(state="disabled")

        # --- Кнопки ---
        frame_buttons = ctk.CTkFrame(self, fg_color="transparent")
        frame_buttons.grid(row=5, column=0, sticky="ew", padx=12, pady=(6, 12))

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
        if email:
            self.entry_email.insert(0, email)
        if token:
            self.entry_token.insert(0, token)

    def _save_env(self):
        env_path = os.path.join(core.get_app_dir(), ".env")
        with open(env_path, "w", encoding="utf-8") as f:
            f.write(f"CALLIBRI_EMAIL={self.entry_email.get()}\n")
            f.write(f"CALLIBRI_TOKEN={self.entry_token.get()}\n")

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
        site_id = proj.get("site_id")
        folder = proj.get("folder", "")
        enabled = proj.get("enabled", True)
        channels = proj.get("channels")
        fmt = proj.get("format", "xlsx")
        fields = proj.get("fields")

        row_frame = ctk.CTkFrame(self.projects_scroll, fg_color="transparent")
        row_frame.pack(fill="x", pady=1)
        row_frame.grid_columnconfigure(1, weight=1)

        var = ctk.IntVar(value=1 if enabled else 0)
        label = f"{folder} ({site_id})"
        if channels:
            label += f" — {', '.join(channels)}"
        label += f" | {fmt}"
        if fields:
            label += f" | {len(fields)} полей"

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
            "site_id": site_id,
            "var": var,
            "frame": row_frame,
        })

    # ── Настройки проекта ─────────────────────────────────────────────────

    def _on_project_settings(self, idx):
        if idx >= len(self._projects_config):
            return
        proj = self._projects_config[idx]
        email = self.entry_email.get().strip()
        token = self.entry_token.get().strip()

        dialog = ProjectSettingsDialog(self, proj, email, token)
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
        email = self.entry_email.get().strip()
        token = self.entry_token.get().strip()

        if not email or not token:
            self._append_log("Заполни email и token для загрузки списка проектов")
            return

        # Загрузка списка проектов из API (с кэшем)
        if self._api_sites is None:
            self._append_log("Загружаем проекты из API...")
            self.btn_add_project.configure(state="disabled")

            def _fetch():
                try:
                    sites = core.get_sites(email, token)
                    self.msg_queue.put(("sites_loaded", sites))
                except Exception as e:
                    self.msg_queue.put(("log", f"Ошибка загрузки проектов: {e}"))
                    self.msg_queue.put(("sites_loaded", None))

            threading.Thread(target=_fetch, daemon=True).start()
        else:
            self._show_add_dialog(self._api_sites)

    def _show_add_dialog(self, sites):
        existing = {p.get("site_id") for p in self._projects_config}
        dialog = AddProjectDialog(self, sites, existing)
        self.wait_window(dialog)

        if dialog.result is not None:
            # Проверяем дубликат
            new_sid = dialog.result["site_id"]
            if any(p.get("site_id") == new_sid for p in self._projects_config):
                self._append_log(f"Проект site_id={new_sid} уже есть в конфиге")
                return

            self._projects_config.append(dialog.result)
            core.save_projects(self._projects_config)
            self._load_projects()
            self._append_log(f"Добавлен проект: {dialog.result['folder']} ({new_sid})")

    # ── Быстрый выбор периода ─────────────────────────────────────────────

    def _set_quick_period(self, days):
        date2 = datetime.now()
        date1 = date2 - timedelta(days=days - 1)
        self.entry_date1.delete(0, "end")
        self.entry_date1.insert(0, date1.strftime("%d.%m.%Y"))
        self.entry_date2.delete(0, "end")
        self.entry_date2.insert(0, date2.strftime("%d.%m.%Y"))

    # ── Проверка соединения ───────────────────────────────────────────────

    def _on_check_connection(self):
        self.btn_check.configure(state="disabled", text="Проверяем...")
        self.lbl_conn_status.configure(text="", text_color="gray")
        self._api_sites = None  # сбрасываем кэш

        email = self.entry_email.get().strip()
        token = self.entry_token.get().strip()

        def _check():
            ok, count, msg = core.test_connection(email, token)
            self.msg_queue.put(("conn_result", ok, msg))

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

        email = self.entry_email.get().strip()
        token = self.entry_token.get().strip()

        enabled_ids = set()
        for pw in self._project_widgets:
            if pw["var"].get():
                enabled_ids.add(pw["site_id"])

        if not enabled_ids:
            self._append_log("Нет выбранных проектов для экспорта")
            return

        self._save_env()

        self.btn_export.configure(state="disabled", text="Экспорт...")
        self.progress_bar.set(0)
        self.lbl_progress.configure(text="Запуск...", text_color="gray")
        self.log_box.configure(state="normal")
        self.log_box.delete("1.0", "end")
        self.log_box.configure(state="disabled")

        params = dict(
            email=email,
            token=token,
            date1_str=date1_str,
            date2_str=date2_str,
            enabled_site_ids=enabled_ids,
            on_log=lambda m: self.msg_queue.put(("log", m)),
            on_progress=lambda pi, pt, ci, ct: self.msg_queue.put(("progress", pi, pt, ci, ct)),
        )

        threading.Thread(target=self._run_export, args=(params,), daemon=True).start()

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
                self.lbl_conn_status.configure(
                    text=message,
                    text_color="green" if ok else "red",
                )

            elif kind == "sites_loaded":
                sites = msg[1]
                self.btn_add_project.configure(state="normal")
                if sites is not None:
                    self._api_sites = sites
                    self._show_add_dialog(sites)

        self.after(100, self._poll_queue)

    # ── Лог ───────────────────────────────────────────────────────────────

    def _append_log(self, text):
        self.log_box.configure(state="normal")
        self.log_box.insert("end", text + "\n")
        self.log_box.see("end")
        self.log_box.configure(state="disabled")


if __name__ == "__main__":
    app = App()
    app.mainloop()
