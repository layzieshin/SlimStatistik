from typing import Dict, Iterable, Optional, List, Tuple
from models.repository import Repository
import datetime

class StatsService:
    def __init__(self, repo: Repository):
        self.repo = repo

    def _excluded(self) -> set:
        # Global: „nicht entnommen?“-Proben ausschließen
        return self.repo.suspected_missing_blood_draw_proben()

    # --- bestehende Auswertungen (unverändert) ---
    def analyte_requests(self, analyte_code: str, start: datetime.datetime, end: datetime.datetime) -> int:
        excl = self._excluded()
        return self.repo.count_analyte_requests(
            analyte_code,
            start.strftime("%Y-%m-%d %H:%M:%S"),
            end.strftime("%Y-%m-%d %H:%M:%S"),
            excl
        )

    def counts_by_status(self, start: datetime.datetime, end: datetime.datetime) -> Dict[str, int]:
        excl = self._excluded()
        per = self.repo.per_sample_completion(
            start.strftime("%Y-%m-%d %H:%M:%S"),
            end.strftime("%Y-%m-%d %H:%M:%S"),
            excl
        )
        out = {"open": 0, "done": 0, "all": 0}
        for st in per.values():
            out[st] += 1
        out["all"] = len(per)
        return out

    def weekday_stats(self, start: datetime.datetime, end: datetime.datetime, status: Optional[str], analyte: Optional[str]) -> Dict[int, Dict[str, float]]:
        excl = self._excluded()
        counts = self.repo.weekday_counts(
            status, analyte,
            start.strftime("%Y-%m-%d %H:%M:%S"),
            end.strftime("%Y-%m-%d %H:%M:%S"),
            excl
        )
        # Durchschnitt pro vorkommendem Wochen-Tag im Intervall
        day = start.date()
        end_date = end.date()
        days_per_weekday = {i: 0 for i in range(7)}
        while day <= end_date:
            days_per_weekday[day.weekday()] += 1
            day += datetime.timedelta(days=1)
        out = {}
        for wd, c in counts.items():
            denom = max(1, days_per_weekday.get(wd, 0))
            out[wd] = {"count": c, "avg": c / denom}
        return out

    # --- NEU: Offene Anforderungen je Analyt (ab Datum X) ---
    def open_requirements_counts_since(self, analytes: Iterable[str], since: datetime.datetime) -> List[Tuple[str, int]]:
        excl = self._excluded()
        rows = self.repo.open_requirements_counts_since(
            list(analytes),
            since.strftime("%Y-%m-%d %H:%M:%S"),
            excl
        )
        return [(r["code"], int(r["c"] or 0)) for r in rows]

    def suspected_missing_blood_draw(self) -> List[dict]:
        return self.repo.list_suspected_missing_blood_draw()
