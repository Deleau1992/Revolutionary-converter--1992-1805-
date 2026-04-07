import csv
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import date, timedelta


# =========================================================
# Franse revolutionaire kalender (historische periode)
# Officieel gebruikt van 22-09-1792 t.e.m. 31-12-1805
# =========================================================

GREGORIAN_START = date(1792, 9, 22)
GREGORIAN_END = date(1805, 12, 31)

MONTHS = [
    "Vendémiaire",
    "Brumaire",
    "Frimaire",
    "Nivôse",
    "Pluviôse",
    "Ventôse",
    "Germinal",
    "Floréal",
    "Prairial",
    "Messidor",
    "Thermidor",
    "Fructidor",
]

GREGORIAN_MONTHS = [
    "Januari", "Februari", "Maart", "April", "Mei", "Juni",
    "Juli", "Augustus", "September", "Oktober", "November", "December"
]

DECADE_DAYS = [
    "Primidi",
    "Duodi",
    "Tridi",
    "Quartidi",
    "Quintidi",
    "Sextidi",
    "Septidi",
    "Octidi",
    "Nonidi",
    "Décadi",
]

SANSCULOTTIDES = [
    "La Fête de la Vertu",
    "La Fête du Génie",
    "La Fête du Travail",
    "La Fête de l'Opinion",
    "La Fête des Récompenses",
    "La Fête de la Révolution",  # alleen in schrikkeljaar
]

LEAP_YEARS = {3, 7, 11}


def is_republican_leap_year(year: int) -> bool:
    return year in LEAP_YEARS


def roman_numeral(n: int) -> str:
    values = [
        (1000, "M"), (900, "CM"), (500, "D"), (400, "CD"),
        (100, "C"), (90, "XC"), (50, "L"), (40, "XL"),
        (10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I")
    ]
    result = []
    for value, numeral in values:
        while n >= value:
            result.append(numeral)
            n -= value
    return "".join(result)


def start_of_republican_year(year: int) -> date:
    if year < 1:
        raise ValueError("Republikeins jaar moet minstens 1 zijn.")

    current = GREGORIAN_START
    for y in range(1, year):
        current += timedelta(days=366 if is_republican_leap_year(y) else 365)
    return current


def republican_year_length(year: int) -> int:
    return 366 if is_republican_leap_year(year) else 365


def gregorian_to_republican(g_date: date) -> dict:
    if not (GREGORIAN_START <= g_date <= GREGORIAN_END):
        raise ValueError(
            "Datum valt buiten de officiële Franse revolutionaire kalender "
            "(22/09/1792 t.e.m. 31/12/1805)."
        )

    year = 1
    while True:
        start = start_of_republican_year(year)
        length = republican_year_length(year)
        end = start + timedelta(days=length - 1)
        if start <= g_date <= end:
            break
        year += 1

    day_of_year = (g_date - start).days + 1

    if day_of_year <= 360:
        month_index = (day_of_year - 1) // 30
        day_in_month = ((day_of_year - 1) % 30) + 1
        decade_day = DECADE_DAYS[(day_in_month - 1) % 10]

        return {
            "type": "month_day",
            "year": year,
            "year_roman": roman_numeral(year),
            "month": MONTHS[month_index],
            "month_number": month_index + 1,
            "day": day_in_month,
            "decade_day": decade_day,
            "day_of_year": day_of_year,
        }
    else:
        complementary_index = day_of_year - 361
        max_comp = 6 if is_republican_leap_year(year) else 5
        if complementary_index > max_comp:
            raise ValueError("Ongeldige complementaire dag.")

        return {
            "type": "complementary",
            "year": year,
            "year_roman": roman_numeral(year),
            "festival_day_number": complementary_index,
            "festival_name": SANSCULOTTIDES[complementary_index - 1],
            "day_of_year": day_of_year,
        }


def republican_to_gregorian(year: int, month: int = None, day: int = None, festival_day: int = None) -> date:
    if year < 1:
        raise ValueError("Republikeins jaar moet minstens 1 zijn.")

    start = start_of_republican_year(year)
    year_start = start
    year_end = start + timedelta(days=republican_year_length(year) - 1)

    if year_end < GREGORIAN_START or year_start > GREGORIAN_END:
        raise ValueError("Republikeins jaar valt buiten de officiële periode.")

    if festival_day is not None:
        max_comp = 6 if is_republican_leap_year(year) else 5
        if not (1 <= festival_day <= max_comp):
            raise ValueError(f"Complementaire dag moet tussen 1 en {max_comp} liggen voor jaar {year}.")
        result = start + timedelta(days=360 + (festival_day - 1))
    else:
        if month is None or day is None:
            raise ValueError("Geef maand en dag op.")
        if not (1 <= month <= 12):
            raise ValueError("Maand moet tussen 1 en 12 liggen.")
        if not (1 <= day <= 30):
            raise ValueError("Dag moet tussen 1 en 30 liggen.")
        result = start + timedelta(days=(month - 1) * 30 + (day - 1))

    if not (GREGORIAN_START <= result <= GREGORIAN_END):
        raise ValueError("Omgezette datum valt buiten de officiële periode.")
    return result


class SearchableTree(ttk.Frame):
    def __init__(self, parent, columns, headings, rows, widths=None):
        super().__init__(parent)
        self.columns_config = columns
        self.all_rows = rows[:]

        top = ttk.Frame(self)
        top.pack(fill="x", pady=(0, 8))

        ttk.Label(top, text="Zoeken:").pack(side="left")
        self.search_var = tk.StringVar()
        entry = ttk.Entry(top, textvariable=self.search_var, width=30)
        entry.pack(side="left", padx=(6, 6))
        entry.bind("<KeyRelease>", lambda e: self.apply_filter())

        ttk.Button(top, text="Wis", command=self.clear_search).pack(side="left")

        self.tree = ttk.Treeview(self, columns=columns, show="headings", height=14)
        self.tree.pack(fill="both", expand=True)

        for i, col in enumerate(columns):
            self.tree.heading(col, text=headings[i])
            width = widths[i] if widths and i < len(widths) else 120
            anchor = "center" if i == 0 else "w"
            self.tree.column(col, width=width, anchor=anchor)

        self.populate(self.all_rows)

    def populate(self, rows):
        for item in self.tree.get_children():
            self.tree.delete(item)
        for row in rows:
            self.tree.insert("", "end", values=row)

    def apply_filter(self):
        needle = self.search_var.get().strip().lower()
        if not needle:
            self.populate(self.all_rows)
            return

        filtered = []
        for row in self.all_rows:
            joined = " | ".join(str(x).lower() for x in row)
            if needle in joined:
                filtered.append(row)
        self.populate(filtered)

    def clear_search(self):
        self.search_var.set("")
        self.populate(self.all_rows)

    def get_visible_rows(self):
        rows = []
        for item in self.tree.get_children():
            rows.append(self.tree.item(item, "values"))
        return rows


class ConverterApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Franse / Napoleontische Datum Converter")
        self.root.geometry("1220x820")
        self.root.minsize(1080, 720)

        self.last_result_text = "Resultaten verschijnen hier..."

        self.style = ttk.Style()
        try:
            self.style.theme_use("clam")
        except Exception:
            pass

        self.configure_styles()
        self.build_ui()

    def configure_styles(self):
        self.root.configure(bg="#eef3f8")

        self.style.configure("TFrame", background="#eef3f8")
        self.style.configure("TLabelframe", background="#eef3f8")
        self.style.configure("TLabelframe.Label", background="#eef3f8")
        self.style.configure("Title.TLabel", font=("Segoe UI", 20, "bold"), background="#eef3f8")
        self.style.configure("SubTitle.TLabel", font=("Segoe UI", 11), background="#eef3f8")
        self.style.configure("Section.TLabelframe.Label", font=("Segoe UI", 11, "bold"))
        self.style.configure("Treeview", rowheight=26, font=("Segoe UI", 10))
        self.style.configure("Treeview.Heading", font=("Segoe UI", 10, "bold"))
        self.style.configure("TButton", font=("Segoe UI", 10), padding=6)
        self.style.configure("TLabel", font=("Segoe UI", 10), background="#eef3f8")
        self.style.configure("TEntry", font=("Segoe UI", 10))
        self.style.configure("TCombobox", font=("Segoe UI", 10))
        self.style.configure("TNotebook", background="#eef3f8")
        self.style.configure("TNotebook.Tab", font=("Segoe UI", 10, "bold"), padding=(12, 8))

    def build_ui(self):
        main = ttk.Frame(self.root, padding=14)
        main.pack(fill="both", expand=True)

        header = ttk.Frame(main)
        header.pack(fill="x", pady=(0, 10))

        ttk.Label(header, text="Franse / Napoleontische Datum Converter", style="Title.TLabel").pack(anchor="w")
        ttk.Label(
            header,
            text="Converteer tussen gewone datums en de Franse revolutionaire kalender (officiële periode 1792–1805).",
            style="SubTitle.TLabel"
        ).pack(anchor="w", pady=(4, 0))

        content = ttk.Panedwindow(main, orient="horizontal")
        content.pack(fill="both", expand=True)

        left = ttk.Frame(content, padding=(0, 0, 8, 0))
        right = ttk.Frame(content)

        content.add(left, weight=3)
        content.add(right, weight=2)

        self.build_left_panel(left)
        self.build_right_panel(right)

    def build_left_panel(self, parent):
        frame1 = ttk.LabelFrame(parent, text="Gewone datum → Revolutionaire datum", style="Section.TLabelframe")
        frame1.pack(fill="x", pady=(0, 10))

        grid1 = ttk.Frame(frame1, padding=10)
        grid1.pack(fill="x")

        ttk.Label(grid1, text="Dag").grid(row=0, column=0, sticky="w", padx=(0, 6), pady=4)
        ttk.Label(grid1, text="Maand").grid(row=0, column=2, sticky="w", padx=(14, 6), pady=4)
        ttk.Label(grid1, text="Jaar").grid(row=0, column=4, sticky="w", padx=(14, 6), pady=4)

        self.g_day_var = tk.StringVar(value="22")
        self.g_month_var = tk.StringVar(value="September")
        self.g_year_var = tk.StringVar(value="1792")

        self.g_day_combo = ttk.Combobox(grid1, textvariable=self.g_day_var, values=[str(i) for i in range(1, 32)], width=8, state="readonly")
        self.g_day_combo.grid(row=0, column=1, sticky="w")

        self.g_month_combo = ttk.Combobox(grid1, textvariable=self.g_month_var, values=GREGORIAN_MONTHS, width=14, state="readonly")
        self.g_month_combo.grid(row=0, column=3, sticky="w")

        self.g_year_combo = ttk.Combobox(grid1, textvariable=self.g_year_var, values=[str(i) for i in range(1792, 1806)], width=10, state="readonly")
        self.g_year_combo.grid(row=0, column=5, sticky="w")

        ttk.Button(grid1, text="Converteer →", command=self.convert_to_republican).grid(row=0, column=6, padx=(18, 0), sticky="w")

        frame2 = ttk.LabelFrame(parent, text="Revolutionaire datum → Gewone datum", style="Section.TLabelframe")
        frame2.pack(fill="x", pady=(0, 10))

        grid2 = ttk.Frame(frame2, padding=10)
        grid2.pack(fill="x")

        ttk.Label(grid2, text="Jaar").grid(row=0, column=0, sticky="w", padx=(0, 6), pady=4)
        self.r_year_var = tk.StringVar(value="1")
        ttk.Combobox(grid2, textvariable=self.r_year_var, values=[str(i) for i in range(1, 15)], width=10, state="readonly").grid(row=0, column=1, sticky="w")

        ttk.Label(grid2, text="Type").grid(row=0, column=2, sticky="w", padx=(14, 6), pady=4)
        self.rep_type_var = tk.StringVar(value="Normale maanddag")
        type_combo = ttk.Combobox(
            grid2,
            textvariable=self.rep_type_var,
            values=["Normale maanddag", "Complementaire dag"],
            width=22,
            state="readonly"
        )
        type_combo.grid(row=0, column=3, sticky="w")
        type_combo.bind("<<ComboboxSelected>>", lambda e: self.update_rep_input_mode())

        self.month_day_frame = ttk.Frame(grid2)
        self.month_day_frame.grid(row=1, column=0, columnspan=6, sticky="w", pady=(10, 0))

        ttk.Label(self.month_day_frame, text="Maand").grid(row=0, column=0, sticky="w", padx=(0, 6), pady=4)
        self.r_month_var = tk.StringVar(value=MONTHS[0])
        self.month_combo = ttk.Combobox(self.month_day_frame, textvariable=self.r_month_var, values=MONTHS, width=18, state="readonly")
        self.month_combo.grid(row=0, column=1, sticky="w")

        ttk.Label(self.month_day_frame, text="Dag").grid(row=0, column=2, sticky="w", padx=(14, 6), pady=4)
        self.r_day_var = tk.StringVar(value="1")
        ttk.Combobox(self.month_day_frame, textvariable=self.r_day_var, values=[str(i) for i in range(1, 31)], width=8, state="readonly").grid(row=0, column=3, sticky="w")

        self.comp_day_frame = ttk.Frame(grid2)
        ttk.Label(self.comp_day_frame, text="Complementaire dag").grid(row=0, column=0, sticky="w", padx=(0, 6), pady=4)
        self.comp_day_var = tk.StringVar(value="1")
        self.comp_day_combo = ttk.Combobox(self.comp_day_frame, textvariable=self.comp_day_var, values=["1", "2", "3", "4", "5", "6"], width=8, state="readonly")
        self.comp_day_combo.grid(row=0, column=1, sticky="w")

        ttk.Button(grid2, text="Converteer ←", command=self.convert_to_gregorian).grid(row=2, column=0, pady=(12, 0), sticky="w")

        self.update_rep_input_mode()

        result_frame = ttk.LabelFrame(parent, text="Resultaat", style="Section.TLabelframe")
        result_frame.pack(fill="both", expand=True)

        toolbar = ttk.Frame(result_frame)
        toolbar.pack(fill="x", padx=10, pady=(10, 0))

        ttk.Button(toolbar, text="Kopieer resultaat", command=self.copy_result).pack(side="left")
        ttk.Button(toolbar, text="Export TXT", command=self.export_txt).pack(side="left", padx=(8, 0))
        ttk.Button(toolbar, text="Export CSV", command=self.export_csv).pack(side="left", padx=(8, 0))

        self.result_text = tk.Text(
            result_frame,
            height=14,
            wrap="word",
            font=("Consolas", 11),
            padx=10,
            pady=10
        )
        self.result_text.pack(fill="both", expand=True, padx=10, pady=10)
        self.result_text.insert("1.0", self.last_result_text)
        self.result_text.config(state="disabled")

    def build_right_panel(self, parent):
        notebook = ttk.Notebook(parent)
        notebook.pack(fill="both", expand=True)

        tab_days = ttk.Frame(notebook)
        tab_months = ttk.Frame(notebook)
        tab_years = ttk.Frame(notebook)

        notebook.add(tab_days, text="Dagen")
        notebook.add(tab_months, text="Maanden")
        notebook.add(tab_years, text="Jaren")

        self.build_days_tab(tab_days)
        self.build_months_tab(tab_months)
        self.build_years_tab(tab_years)

    def build_days_tab(self, parent):
        container = ttk.Frame(parent, padding=10)
        container.pack(fill="both", expand=True)

        ttk.Label(container, text="Dagentabel van de 10-daagse week (décade)", font=("Segoe UI", 11, "bold")).pack(anchor="w", pady=(0, 8))

        meanings = [
            "Eerste dag",
            "Tweede dag",
            "Derde dag",
            "Vierde dag",
            "Vijfde dag",
            "Zesde dag",
            "Zevende dag",
            "Achtste dag",
            "Negende dag",
            "Tiende dag",
        ]
        rows = [(i, name, meaning) for i, (name, meaning) in enumerate(zip(DECADE_DAYS, meanings), start=1)]
        self.days_table = SearchableTree(container, ("num", "name", "meaning"), ("Nr", "Naam", "Betekenis"), rows, [60, 140, 260])
        self.days_table.pack(fill="both", expand=True)

    def build_months_tab(self, parent):
        container = ttk.Frame(parent, padding=10)
        container.pack(fill="both", expand=True)

        ttk.Label(container, text="Maandtabel", font=("Segoe UI", 11, "bold")).pack(anchor="w", pady=(0, 8))

        month_info = [
            ("Vendémiaire", "Wijn-/druivenoogst", "sep–okt"),
            ("Brumaire", "Mist", "okt–nov"),
            ("Frimaire", "Rijp / koude", "nov–dec"),
            ("Nivôse", "Sneeuw", "dec–jan"),
            ("Pluviôse", "Regen", "jan–feb"),
            ("Ventôse", "Wind", "feb–mrt"),
            ("Germinal", "Ontkieming", "mrt–apr"),
            ("Floréal", "Bloei", "apr–mei"),
            ("Prairial", "Weiden", "mei–jun"),
            ("Messidor", "Oogst", "jun–jul"),
            ("Thermidor", "Hitte", "jul–aug"),
            ("Fructidor", "Vruchten", "aug–sep"),
        ]
        rows = [(i, month, season, approx) for i, (month, season, approx) in enumerate(month_info, start=1)]
        self.months_table = SearchableTree(container, ("num", "month", "season", "approx"), ("Nr", "Maand", "Seizoen / idee", "Ongeveer"), rows, [50, 150, 200, 120])
        self.months_table.pack(fill="both", expand=True)

        info = ttk.Label(container, text="Na maand 12 volgen 5 of 6 complementaire feestdagen (Sansculottides).")
        info.pack(anchor="w", pady=(8, 0))

    def build_years_tab(self, parent):
        container = ttk.Frame(parent, padding=10)
        container.pack(fill="both", expand=True)

        ttk.Label(container, text="Jaaroverzicht (officiële periode)", font=("Segoe UI", 11, "bold")).pack(anchor="w", pady=(0, 8))

        rows = []
        for y in range(1, 15):
            start = start_of_republican_year(y)
            length = republican_year_length(y)
            end = start + timedelta(days=length - 1)

            if start > GREGORIAN_END:
                break

            shown_end = min(end, GREGORIAN_END)
            rows.append((
                y,
                roman_numeral(y),
                start.strftime("%d/%m/%Y"),
                shown_end.strftime("%d/%m/%Y"),
                length if end <= GREGORIAN_END else "—",
                "Ja" if is_republican_leap_year(y) else "Nee"
            ))

        self.years_table = SearchableTree(
            container,
            ("year", "roman", "start", "end", "days", "leap"),
            ("Jaar", "Romeins", "Begint op", "Eindigt op", "Dagen", "Schrikkeljaar"),
            rows,
            [60, 90, 110, 110, 70, 90]
        )
        self.years_table.pack(fill="both", expand=True)

    def update_rep_input_mode(self):
        mode = self.rep_type_var.get()
        if mode == "Complementaire dag":
            self.month_day_frame.grid_remove()
            self.comp_day_frame.grid(row=1, column=0, columnspan=6, sticky="w", pady=(10, 0))
        else:
            self.comp_day_frame.grid_remove()
            self.month_day_frame.grid()

    def month_name_to_number(self, month_name):
        return GREGORIAN_MONTHS.index(month_name) + 1

    def set_result(self, text: str):
        self.last_result_text = text
        self.result_text.config(state="normal")
        self.result_text.delete("1.0", "end")
        self.result_text.insert("1.0", text)
        self.result_text.config(state="disabled")

    def copy_result(self):
        self.root.clipboard_clear()
        self.root.clipboard_append(self.last_result_text)
        self.root.update()
        messagebox.showinfo("Gekopieerd", "Resultaat staat nu op je klembord.")

    def export_txt(self):
        path = filedialog.asksaveasfilename(
            title="Exporteer resultaat naar TXT",
            defaultextension=".txt",
            filetypes=[("Tekstbestand", "*.txt"), ("Alle bestanden", "*.*")]
        )
        if not path:
            return
        try:
            with open(path, "w", encoding="utf-8") as f:
                f.write(self.last_result_text)
            messagebox.showinfo("Export gelukt", f"TXT opgeslagen:\n{path}")
        except Exception as e:
            messagebox.showerror("Export fout", str(e))

    def export_csv(self):
        path = filedialog.asksaveasfilename(
            title="Exporteer tabellen naar CSV",
            defaultextension=".csv",
            filetypes=[("CSV bestand", "*.csv"), ("Alle bestanden", "*.*")]
        )
        if not path:
            return
        try:
            with open(path, "w", newline="", encoding="utf-8-sig") as f:
                writer = csv.writer(f, delimiter=";")
                writer.writerow(["=== RESULTAAT ==="])
                for line in self.last_result_text.splitlines():
                    writer.writerow([line])

                writer.writerow([])
                writer.writerow(["=== DAGEN TABEL ==="])
                writer.writerow(["Nr", "Naam", "Betekenis"])
                for row in self.days_table.get_visible_rows():
                    writer.writerow(row)

                writer.writerow([])
                writer.writerow(["=== MAANDEN TABEL ==="])
                writer.writerow(["Nr", "Maand", "Seizoen / idee", "Ongeveer"])
                for row in self.months_table.get_visible_rows():
                    writer.writerow(row)

                writer.writerow([])
                writer.writerow(["=== JAREN TABEL ==="])
                writer.writerow(["Jaar", "Romeins", "Begint op", "Eindigt op", "Dagen", "Schrikkeljaar"])
                for row in self.years_table.get_visible_rows():
                    writer.writerow(row)

            messagebox.showinfo("Export gelukt", f"CSV opgeslagen:\n{path}")
        except Exception as e:
            messagebox.showerror("Export fout", str(e))

    def convert_to_republican(self):
        try:
            d = int(self.g_day_var.get())
            m = self.month_name_to_number(self.g_month_var.get())
            y = int(self.g_year_var.get())

            g = date(y, m, d)
            result = gregorian_to_republican(g)

            if result["type"] == "month_day":
                text = (
                    f"GEWONE DATUM\n"
                    f"------------------------------\n"
                    f"{g.strftime('%d/%m/%Y')}\n\n"
                    f"REVOLUTIONAIRE DATUM\n"
                    f"------------------------------\n"
                    f"Dag: {result['day']}\n"
                    f"Maand: {result['month']} (#{result['month_number']})\n"
                    f"Jaar: {result['year']} (Jaar {result['year_roman']})\n"
                    f"Dag van de décade: {result['decade_day']}\n"
                    f"Dag van het jaar: {result['day_of_year']}\n\n"
                    f"Volledige notatie:\n"
                    f"{result['day']} {result['month']} an {result['year_roman']}"
                )
            else:
                text = (
                    f"GEWONE DATUM\n"
                    f"------------------------------\n"
                    f"{g.strftime('%d/%m/%Y')}\n\n"
                    f"REVOLUTIONAIRE DATUM\n"
                    f"------------------------------\n"
                    f"Complementaire dag: {result['festival_day_number']}\n"
                    f"Naam: {result['festival_name']}\n"
                    f"Jaar: {result['year']} (Jaar {result['year_roman']})\n"
                    f"Dag van het jaar: {result['day_of_year']}\n\n"
                    f"Volledige notatie:\n"
                    f"Jour complémentaire {result['festival_day_number']} - an {result['year_roman']}"
                )

            self.set_result(text)

        except Exception as e:
            messagebox.showerror("Fout", str(e))

    def convert_to_gregorian(self):
        try:
            year = int(self.r_year_var.get())
            mode = self.rep_type_var.get()

            if mode == "Complementaire dag":
                festival_day = int(self.comp_day_var.get())
                g = republican_to_gregorian(year=year, festival_day=festival_day)
                name = SANSCULOTTIDES[festival_day - 1]

                text = (
                    f"REVOLUTIONAIRE DATUM\n"
                    f"------------------------------\n"
                    f"Complementaire dag: {festival_day}\n"
                    f"Naam: {name}\n"
                    f"Jaar: {year} (Jaar {roman_numeral(year)})\n\n"
                    f"GEWONE DATUM\n"
                    f"------------------------------\n"
                    f"{g.strftime('%d/%m/%Y')}"
                )
            else:
                month_name = self.r_month_var.get()
                day = int(self.r_day_var.get())
                month = MONTHS.index(month_name) + 1

                g = republican_to_gregorian(year=year, month=month, day=day)
                decade_day = DECADE_DAYS[(day - 1) % 10]

                text = (
                    f"REVOLUTIONAIRE DATUM\n"
                    f"------------------------------\n"
                    f"Dag: {day}\n"
                    f"Maand: {month_name} (#{month})\n"
                    f"Jaar: {year} (Jaar {roman_numeral(year)})\n"
                    f"Décade-dag: {decade_day}\n\n"
                    f"GEWONE DATUM\n"
                    f"------------------------------\n"
                    f"{g.strftime('%d/%m/%Y')}"
                )

            self.set_result(text)

        except Exception as e:
            messagebox.showerror("Fout", str(e))


def main():
    root = tk.Tk()
    app = ConverterApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()