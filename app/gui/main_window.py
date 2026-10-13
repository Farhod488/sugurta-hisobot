from __future__ import annotations

import threading
import traceback
from pathlib import Path
from tkinter import filedialog, messagebox

import customtkinter as ctk

from app.tushum import Inspector, TushumInputs, generate_tushum_report
from app.tushum.assignments import PERSONAL_INSPECTORS, default_selected_names
from app.tushum.parsers import parse_kuzatuv_xlsx, parse_osago_xlsx


PROJECT_ROOT = Path(__file__).resolve().parents[2]
DEFAULT_OUTPUT = PROJECT_ROOT / "output"
DEFAULT_TEMPLATES = PROJECT_ROOT / "templates"
DEFAULT_SOURCE = PROJECT_ROOT / "templates" / "source"


class App(ctk.CTk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Sug'urta hisobot — Тушум")
        self.geometry("920x820")
        self.minsize(820, 720)

        ctk.set_appearance_mode("System")
        ctk.set_default_color_theme("blue")

        self.kuzatuv_pdf = ctk.StringVar()
        self.kuzatuv_xlsx = ctk.StringVar()
        self.osago_pdf = ctk.StringVar()
        self.osago_xlsx = ctk.StringVar()
        self.output_dir = ctk.StringVar(value=str(DEFAULT_OUTPUT))
        self.period_text = ctk.StringVar(value="24-31-avgust")
        self.inspector = ctk.StringVar(value=Inspector.ARIPOV.value)
        self.status_text = ctk.StringVar(
            value=(
                "Укажите период и файлы. После загрузки Excel появится выбор компаний "
                "(по умолчанию — текущие назначения)."
            )
        )

        self.kuzatuv_companies: list[str] = []
        self.osago_companies: list[str] = []
        # inspector -> company name -> BooleanVar
        self.kuz_vars: dict[Inspector, dict[str, ctk.BooleanVar]] = {
            item: {} for item in PERSONAL_INSPECTORS
        }
        self.osa_vars: dict[Inspector, dict[str, ctk.BooleanVar]] = {
            item: {} for item in PERSONAL_INSPECTORS
        }

        self._autoload_source_files()
        self._build_ui()
        self._reload_kuzatuv_companies()
        self._reload_osago_companies()

        self.kuzatuv_xlsx.trace_add("write", lambda *_: self._reload_kuzatuv_companies())
        self.osago_xlsx.trace_add("write", lambda *_: self._reload_osago_companies())

    def _autoload_source_files(self) -> None:
        folder = DEFAULT_SOURCE if DEFAULT_SOURCE.exists() else DEFAULT_TEMPLATES
        mapping = [
            (self.kuzatuv_pdf, "KUZATUV", ".pdf"),
            (self.kuzatuv_xlsx, "KUZATUV", ".xlsx"),
            (self.osago_pdf, "OSAGO", ".pdf"),
            (self.osago_xlsx, "OSAGO", ".xlsx"),
        ]
        if not folder.exists():
            return
        for var, prefix, suffix in mapping:
            matches = sorted(
                p
                for p in folder.iterdir()
                if p.is_file()
                and p.suffix.lower() == suffix
                and p.name.upper().startswith(prefix)
            )
            if matches:
                var.set(str(matches[0]))

    def _build_ui(self) -> None:
        outer = ctk.CTkScrollableFrame(self, corner_radius=12)
        outer.pack(fill="both", expand=True, padx=14, pady=14)

        ctk.CTkLabel(
            outer,
            text="Sug'urta hisobot — отчёты Тушум",
            font=ctk.CTkFont(size=22, weight="bold"),
        ).pack(anchor="w", padx=8, pady=(8, 4))

        ctk.CTkLabel(
            outer,
            text="PDF/Excel необязательны по отдельности. Компании — после загрузки Excel.",
            text_color=("gray30", "gray70"),
        ).pack(anchor="w", padx=8, pady=(0, 10))

        period_row = ctk.CTkFrame(outer, fg_color="transparent")
        period_row.pack(fill="x", padx=8, pady=4)
        ctk.CTkLabel(
            period_row, text="Период", font=ctk.CTkFont(weight="bold")
        ).pack(anchor="w")
        ctk.CTkEntry(period_row, textvariable=self.period_text).pack(fill="x", pady=(4, 0))

        self._row_file(
            outer, "KUZATUV PDF (абз. 4)", self.kuzatuv_pdf, self._browse_kuz_pdf
        )
        self._row_file(
            outer, "KUZATUV Excel (абз. 5)", self.kuzatuv_xlsx, self._browse_kuz_xlsx
        )
        self._row_file(outer, "OSAGO PDF (абз. 8)", self.osago_pdf, self._browse_osa_pdf)
        self._row_file(
            outer, "OSAGO Excel (абз. 9)", self.osago_xlsx, self._browse_osa_xlsx
        )
        self._row_file(outer, "Папка для отчётов", self.output_dir, self._browse_output)

        kind_frame = ctk.CTkFrame(outer, fg_color="transparent")
        kind_frame.pack(fill="x", padx=8, pady=(10, 4))
        ctk.CTkLabel(kind_frame, text="Тип отчёта", font=ctk.CTkFont(weight="bold")).pack(
            anchor="w"
        )
        kinds = ctk.CTkFrame(kind_frame, fg_color="transparent")
        kinds.pack(fill="x", pady=(6, 0))
        for item in Inspector:
            ctk.CTkRadioButton(
                kinds,
                text=item.label_ru,
                variable=self.inspector,
                value=item.value,
            ).pack(side="left", padx=(0, 14))

        # Company assignment panel
        self.assign_frame = ctk.CTkFrame(outer, corner_radius=10)
        self.assign_frame.pack(fill="x", padx=8, pady=(12, 8))
        ctk.CTkLabel(
            self.assign_frame,
            text="Назначение компаний",
            font=ctk.CTkFont(size=15, weight="bold"),
        ).pack(anchor="w", padx=12, pady=(10, 4))
        ctk.CTkLabel(
            self.assign_frame,
            text="Хамракулов всегда берёт итог по всем компаниям Excel. "
            "Ниже — только Арипов и Абдимуродов.",
            text_color=("gray30", "gray70"),
            wraplength=820,
            justify="left",
        ).pack(anchor="w", padx=12, pady=(0, 8))

        self.assign_hint = ctk.CTkLabel(
            self.assign_frame,
            text="Загрузите KUZATUV.xlsx и/или OSAGO.xlsx — здесь появятся списки.",
            text_color=("gray40", "gray60"),
        )
        self.assign_hint.pack(anchor="w", padx=12, pady=(0, 8))

        self.tabs = ctk.CTkTabview(self.assign_frame, height=260)
        self.tabs.pack(fill="both", expand=True, padx=12, pady=(0, 12))
        self.tab_frames: dict[Inspector, ctk.CTkFrame] = {}
        self.kuz_boxes_frames: dict[Inspector, ctk.CTkScrollableFrame] = {}
        self.osa_boxes_frames: dict[Inspector, ctk.CTkScrollableFrame] = {}

        for item in PERSONAL_INSPECTORS:
            short = {
                Inspector.ARIPOV: "Ш.Арипов",
                Inspector.ABDIMURODOV: "Ш.Абдимуродов",
            }[item]
            tab = self.tabs.add(short)
            self.tab_frames[item] = tab

            cols = ctk.CTkFrame(tab, fg_color="transparent")
            cols.pack(fill="both", expand=True, padx=4, pady=4)
            cols.grid_columnconfigure(0, weight=1)
            cols.grid_columnconfigure(1, weight=1)

            left = ctk.CTkFrame(cols, corner_radius=8)
            left.grid(row=0, column=0, sticky="nsew", padx=(0, 6))
            ctk.CTkLabel(
                left, text="KUZATUV", font=ctk.CTkFont(weight="bold")
            ).pack(anchor="w", padx=8, pady=(8, 4))
            kuz_scroll = ctk.CTkScrollableFrame(left, height=160)
            kuz_scroll.pack(fill="both", expand=True, padx=6, pady=(0, 8))
            self.kuz_boxes_frames[item] = kuz_scroll

            right = ctk.CTkFrame(cols, corner_radius=8)
            right.grid(row=0, column=1, sticky="nsew", padx=(6, 0))
            ctk.CTkLabel(
                right, text="OSAGO", font=ctk.CTkFont(weight="bold")
            ).pack(anchor="w", padx=8, pady=(8, 4))
            osa_scroll = ctk.CTkScrollableFrame(right, height=160)
            osa_scroll.pack(fill="both", expand=True, padx=6, pady=(0, 8))
            self.osa_boxes_frames[item] = osa_scroll

        actions = ctk.CTkFrame(outer, fg_color="transparent")
        actions.pack(fill="x", padx=8, pady=(8, 4))

        self.generate_btn = ctk.CTkButton(
            actions, text="Сформировать отчёт", height=40, command=self._on_generate
        )
        self.generate_btn.pack(side="left")

        ctk.CTkButton(
            actions,
            text="Сформировать все 3",
            height=40,
            command=self._on_generate_all,
        ).pack(side="left", padx=(10, 0))

        ctk.CTkButton(
            actions,
            text="Открыть папку",
            height=40,
            fg_color=("gray75", "gray35"),
            text_color=("gray10", "gray90"),
            command=self._open_output_folder,
        ).pack(side="left", padx=(10, 0))

        status_box = ctk.CTkFrame(outer, corner_radius=10)
        status_box.pack(fill="x", padx=8, pady=(10, 8))
        ctk.CTkLabel(status_box, text="Статус", font=ctk.CTkFont(weight="bold")).pack(
            anchor="w", padx=12, pady=(10, 4)
        )
        ctk.CTkLabel(
            status_box,
            textvariable=self.status_text,
            justify="left",
            wraplength=820,
            anchor="nw",
        ).pack(fill="x", padx=12, pady=(0, 12))

    def _row_file(self, parent, label: str, variable: ctk.StringVar, browse) -> None:
        frame = ctk.CTkFrame(parent, fg_color="transparent")
        frame.pack(fill="x", padx=8, pady=3)
        ctk.CTkLabel(frame, text=label, font=ctk.CTkFont(weight="bold")).pack(anchor="w")
        row = ctk.CTkFrame(frame, fg_color="transparent")
        row.pack(fill="x", pady=(4, 0))
        ctk.CTkEntry(row, textvariable=variable).pack(side="left", fill="x", expand=True)
        ctk.CTkButton(row, text="Обзор…", width=90, command=browse).pack(
            side="left", padx=(8, 0)
        )
        ctk.CTkButton(
            row,
            text="×",
            width=36,
            fg_color=("gray75", "gray35"),
            text_color=("gray10", "gray90"),
            command=lambda: variable.set(""),
        ).pack(side="left", padx=(6, 0))

    def _browse_kuz_pdf(self) -> None:
        self._browse_open(self.kuzatuv_pdf, [("PDF", "*.pdf")])

    def _browse_kuz_xlsx(self) -> None:
        self._browse_open(self.kuzatuv_xlsx, [("Excel", "*.xlsx")])

    def _browse_osa_pdf(self) -> None:
        self._browse_open(self.osago_pdf, [("PDF", "*.pdf")])

    def _browse_osa_xlsx(self) -> None:
        self._browse_open(self.osago_xlsx, [("Excel", "*.xlsx")])

    def _browse_open(self, variable: ctk.StringVar, filetypes) -> None:
        path = filedialog.askopenfilename(filetypes=filetypes + [("Все файлы", "*.*")])
        if path:
            variable.set(path)

    def _browse_output(self) -> None:
        path = filedialog.askdirectory()
        if path:
            self.output_dir.set(path)

    def _open_output_folder(self) -> None:
        folder = Path(self.output_dir.get().strip() or DEFAULT_OUTPUT)
        folder.mkdir(parents=True, exist_ok=True)
        try:
            import os

            os.startfile(folder)  # noqa: S606
        except OSError as exc:
            messagebox.showerror("Ошибка", str(exc))

    def _clear_frame(self, frame) -> None:
        for child in frame.winfo_children():
            child.destroy()

    def _update_assign_hint(self) -> None:
        parts = []
        if self.kuzatuv_companies:
            parts.append(f"KUZATUV: {len(self.kuzatuv_companies)} комп.")
        if self.osago_companies:
            parts.append(f"OSAGO: {len(self.osago_companies)} комп.")
        if parts:
            self.assign_hint.configure(text="Загружено — " + ", ".join(parts))
        else:
            self.assign_hint.configure(
                text="Загрузите KUZATUV.xlsx и/или OSAGO.xlsx — здесь появятся списки."
            )

    def _rebuild_checkboxes(
        self,
        *,
        source: str,
        companies: list[str],
        vars_map: dict[Inspector, dict[str, ctk.BooleanVar]],
        frames: dict[Inspector, ctk.CTkScrollableFrame],
    ) -> None:
        for inspector in PERSONAL_INSPECTORS:
            frame = frames[inspector]
            self._clear_frame(frame)
            vars_map[inspector] = {}
            if not companies:
                ctk.CTkLabel(
                    frame,
                    text="Файл не загружен",
                    text_color=("gray45", "gray55"),
                ).pack(anchor="w", padx=4, pady=4)
                continue

            defaults = set(
                default_selected_names(companies, inspector, source=source)
            )
            for name in companies:
                var = ctk.BooleanVar(value=name in defaults)
                vars_map[inspector][name] = var
                ctk.CTkCheckBox(frame, text=name, variable=var).pack(
                    anchor="w", padx=4, pady=2
                )

    def _reload_kuzatuv_companies(self) -> None:
        path_text = self.kuzatuv_xlsx.get().strip()
        self.kuzatuv_companies = []
        if path_text:
            path = Path(path_text)
            if path.exists():
                try:
                    data = parse_kuzatuv_xlsx(path)
                    self.kuzatuv_companies = [row.name for row in data.rows]
                except Exception as exc:  # noqa: BLE001
                    self.status_text.set(f"Не удалось прочитать KUZATUV Excel: {exc}")
                    self.kuzatuv_companies = []
        self._rebuild_checkboxes(
            source="kuzatuv",
            companies=self.kuzatuv_companies,
            vars_map=self.kuz_vars,
            frames=self.kuz_boxes_frames,
        )
        self._update_assign_hint()

    def _reload_osago_companies(self) -> None:
        path_text = self.osago_xlsx.get().strip()
        self.osago_companies = []
        if path_text:
            path = Path(path_text)
            if path.exists():
                try:
                    data = parse_osago_xlsx(path)
                    self.osago_companies = [row.name for row in data.rows]
                except Exception as exc:  # noqa: BLE001
                    self.status_text.set(f"Не удалось прочитать OSAGO Excel: {exc}")
                    self.osago_companies = []
        self._rebuild_checkboxes(
            source="osago",
            companies=self.osago_companies,
            vars_map=self.osa_vars,
            frames=self.osa_boxes_frames,
        )
        self._update_assign_hint()

    def _selected_map(
        self, vars_map: dict[Inspector, dict[str, ctk.BooleanVar]]
    ) -> dict[Inspector, list[str]]:
        result: dict[Inspector, list[str]] = {}
        for inspector, mapping in vars_map.items():
            selected = [name for name, var in mapping.items() if var.get()]
            result[inspector] = selected
        return result

    def _path_or_none(self, value: str) -> Path | None:
        text = value.strip()
        return Path(text) if text else None

    def _inputs(self) -> TushumInputs:
        kuz_pdf = self._path_or_none(self.kuzatuv_pdf.get())
        kuz_xlsx = self._path_or_none(self.kuzatuv_xlsx.get())
        osa_pdf = self._path_or_none(self.osago_pdf.get())
        osa_xlsx = self._path_or_none(self.osago_xlsx.get())
        if not any([kuz_pdf, kuz_xlsx, osa_pdf, osa_xlsx]):
            raise ValueError(
                "Выберите хотя бы один файл. Без файла его абзац останется из шаблона."
            )

        kuz_assign = self._selected_map(self.kuz_vars) if kuz_xlsx else None
        osa_assign = self._selected_map(self.osa_vars) if osa_xlsx else None

        return TushumInputs(
            kuzatuv_pdf=kuz_pdf,
            kuzatuv_xlsx=kuz_xlsx,
            osago_pdf=osa_pdf,
            osago_xlsx=osa_xlsx,
            period_text=self.period_text.get().strip(),
            templates_dir=DEFAULT_TEMPLATES,
            output_dir=Path(self.output_dir.get().strip() or DEFAULT_OUTPUT),
            kuzatuv_assignments=kuz_assign,
            osago_assignments=osa_assign,
        )

    def _validate_assignments(
        self, inputs: TushumInputs, inspectors: list[Inspector]
    ) -> None:
        for inspector in inspectors:
            if inspector not in PERSONAL_INSPECTORS:
                continue
            if inputs.kuzatuv_xlsx and inputs.kuzatuv_assignments is not None:
                selected = inputs.kuzatuv_assignments.get(inspector) or []
                if not selected:
                    raise ValueError(
                        f"{inspector.label_ru}: выберите хотя бы одну компанию KUZATUV."
                    )
            if inputs.osago_xlsx and inputs.osago_assignments is not None:
                selected = inputs.osago_assignments.get(inspector) or []
                if not selected:
                    raise ValueError(
                        f"{inspector.label_ru}: выберите хотя бы одну компанию OSAGO."
                    )

    def _on_generate(self) -> None:
        self._start_worker([Inspector(self.inspector.get())])

    def _on_generate_all(self) -> None:
        self._start_worker(list(Inspector))

    def _start_worker(self, inspectors: list[Inspector]) -> None:
        try:
            inputs = self._inputs()
            self._validate_assignments(inputs, inspectors)
        except ValueError as exc:
            messagebox.showwarning("Проверка", str(exc))
            return

        self.generate_btn.configure(state="disabled")
        self.status_text.set("Формирование отчёта…")
        threading.Thread(
            target=self._worker, args=(inspectors, inputs), daemon=True
        ).start()

    def _worker(self, inspectors: list[Inspector], inputs: TushumInputs) -> None:
        try:
            results = [generate_tushum_report(item, inputs) for item in inspectors]
            self.after(0, lambda r=results: self._on_success(r))
        except Exception as exc:  # noqa: BLE001
            details = traceback.format_exc()
            self.after(0, lambda e=exc, d=details: self._on_error(e, d))

    def _on_success(self, results: list[Path]) -> None:
        self.generate_btn.configure(state="normal")
        joined = "\n".join(str(path) for path in results)
        self.status_text.set(f"Готово:\n{joined}")
        messagebox.showinfo("Готово", f"Сохранено:\n{joined}")

    def _on_error(self, exc: Exception, details: str) -> None:
        self.generate_btn.configure(state="normal")
        self.status_text.set(f"Ошибка: {exc}")
        messagebox.showerror("Ошибка", f"{exc}\n\n{details[-1500:]}")


def run() -> None:
    DEFAULT_OUTPUT.mkdir(parents=True, exist_ok=True)
    app = App()
    app.mainloop()
