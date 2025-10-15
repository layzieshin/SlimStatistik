import configparser, datetime, os
from typing import List, Optional, Iterable, Tuple
from models.repository import Repository
from logic.stats_service import StatsService
from logic.housekeeping_service import HousekeepingService

class MainController:
    def __init__(self, settings_path: str, mapping_path: str):
        self.cfg = configparser.ConfigParser()
        self.cfg.read(settings_path, encoding="utf-8")

        # Sicherstellen, dass Sektion [paths] existiert + Defaults setzen
        if "paths" not in self.cfg:
            self.cfg["paths"] = {}
        self.paths = self.cfg["paths"]

        base = os.getcwd()
        self.paths.setdefault("database_path", "")  # kann der/die Nutzer:in später setzen
        self.paths.setdefault("excel_file", os.path.join(base, "exports", "monthly.xlsx"))
        self.paths.setdefault("export_dir", os.path.join(base, "exports"))
        self.paths.setdefault("analytes_file", os.path.join(base, "analytes.txt"))

        # Config speichern, falls neu/ergänzt
        os.makedirs(os.path.dirname(settings_path), exist_ok=True)
        with open(settings_path, "w", encoding="utf-8") as f:
            self.cfg.write(f)

        # Services
        self.repo = Repository(self.paths.get("database_path"), mapping_path)
        self.stats = StatsService(self.repo)
        self.housekeeping = HousekeepingService(self.repo, self.paths.get("excel_file"))
        self.analytes = self._load_analytes(self.paths.get("analytes_file"))

    def _load_analytes(self, path: str) -> List[str]:
        import os
        if not os.path.exists(path):
            return []
        out = []
        with open(path, "r", encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if not line or line.startswith("#"):
                    continue
                code = line.split(";")[0].strip()
                if code:
                    out.append(code)
        return out

    # ---------------- Business Facade (UI-frei) ----------------
    def build_counts_rows(
        self, start: datetime.datetime, end: datetime.datetime, analyte: Optional[str], only_open: bool
    ) -> List[Tuple[str, str, str, str]]:
        rows: List[Tuple[str, str, str, str]] = []
        if analyte:
            try:
                n = self.stats.analyte_requests(analyte, start, end)
                rows.append(("Anforderungen (Analyt)", str(n), "", f"Analyt={analyte}"))
            except Exception as ex:
                rows.append(("Fehler Task 1", "-", str(ex), ""))
        try:
            cnt = self.stats.counts_by_status(start, end)
            rows.append(("Befunde gesamt", str(cnt.get("all", 0)), "", ""))
            rows.append(("Befunde offen", str(cnt.get("open", 0)), "", ""))
            rows.append(("Befunde fertig", str(cnt.get("done", 0)), "", ""))
        except Exception as ex:
            rows.append(("Fehler Task 2", "-", str(ex), ""))
        try:
            status_filter = "open" if only_open else "all"
            wd = self.stats.weekday_stats(start, end, status_filter, analyte)
            for i, name in enumerate(["Mo","Di","Mi","Do","Fr","Sa","So"]):
                rows.append((f"Wochentag {name}", str(wd[i]["count"]), f"Ø {wd[i]['avg']:.2f} / Tag",
                             f"Filter: Status={status_filter}, Analyt={analyte or '—'}"))
        except Exception as ex:
            rows.append(("Fehler Task 3/4", "-", str(ex), ""))
        try:
            excl = self.stats.suspected_missing_blood_draw()
            rows.append(("Ausgeschlossene Proben (nicht entnommen?)", str(len(excl)),
                         "Automatisch aus allen Zählungen ausgeschlossen", "Regel: >24h, >1 Anforderung, 0 Ergebnisse"))
        except Exception as ex:
            rows.append(("Hinweis Ausschlüsse (Fehler)", "-", str(ex), ""))
        return rows

    def build_open_counts_since(self, analytes: Iterable[str], since: datetime.datetime):
        return self.stats.open_requirements_counts_since(analytes, since)

    # ---------------- Direct calls ----------------
    def suspected_missing_blood_draw(self):
        return self.stats.suspected_missing_blood_draw()

    def delete_sample(self, proben_nr: str) -> int:
        return self.repo.delete_sample(proben_nr)

    def monthly_export_if_first(self) -> str:
        return self.housekeeping.ensure_monthly_export(self.analytes)
