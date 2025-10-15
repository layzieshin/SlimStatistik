from typing import Dict, Iterable, Optional, List, Tuple
from collections import Counter
from itertools import combinations
import datetime

from models.repository import Repository

class StatsService:
    def __init__(self, repo: Repository):
        self.repo = repo

    def suspected_missing_blood_draw(self, now: Optional[datetime.datetime] = None):
        return self.repo.suspected_missing_blood_draw_proben(now)

    def analyte_requests(self, analyte_code: str, start: datetime.datetime, end: datetime.datetime) -> int:
        excl = self.suspected_missing_blood_draw()
        return self.repo.count_analyte_requests(
            analyte_code,
            start.strftime("%Y-%m-%d %H:%M:%S"),
            end.strftime("%Y-%m-%d %H:%M:%S"),
            excl,
        )

    def counts_by_status(self, start: datetime.datetime, end: datetime.datetime) -> Dict[str, int]:
        excl = self.suspected_missing_blood_draw()
        per = self.repo.per_sample_completion(
            start.strftime("%Y-%m-%d %H:%M:%S"),
            end.strftime("%Y-%m-%d %H:%M:%S"),
            excl,
        )
        out = {"open": 0, "done": 0, "all": 0}
        for st in per.values():
            out[st] += 1
        out["all"] = len(per)
        return out

    def weekday_stats_multi(
        self,
        start: datetime.datetime,
        end: datetime.datetime,
        status: Optional[str],
        analytes: Iterable[str],
    ) -> Dict[int, Dict[str, float]]:
        excl = self.suspected_missing_blood_draw()
        counts = self.repo.weekday_counts_multi(
            status or "all",
            list(analytes),
            start.strftime("%Y-%m-%d %H:%M:%S"),
            end.strftime("%Y-%m-%d %H:%M:%S"),
            excl,
        )
        # Mittelwerte auf Basis der Tage im Intervall
        day = start.date()
        end_date = end.date()
        days_per_weekday = {i: 0 for i in range(7)}
        while day <= end_date:
            days_per_weekday[day.weekday()] += 1
            day += datetime.timedelta(days=1)
        return {wd: {"count": c, "avg": c / max(1, days_per_weekday.get(wd, 0))} for wd, c in counts.items()}

    def get_deleted_sample_details(self, proben_nr: str):
        return self.repo.get_deleted_sample_details(proben_nr)

    def combo_stats_since(self, since: datetime.datetime):
        excl = self.suspected_missing_blood_draw()
        per = self.repo.open_analytes_per_sample_since(since.strftime("%Y-%m-%d %H:%M:%S"), excl)
        sing = Counter()
        pairs = Counter()
        trips = Counter()
        for codes in per:
            if len(codes) == 1:
                sing[codes[0]] += 1
            if len(codes) >= 2:
                for a, b in combinations(codes, 2):
                    pairs[tuple(sorted((a, b)))] += 1
            if len(codes) >= 3:
                for a, b, c in combinations(codes, 3):
                    trips[tuple(sorted((a, b, c)))] += 1
        sing_list = sorted(sing.items(), key=lambda x: (-x[1], x[0]))
        pair_list = [(" + ".join(k), v) for k, v in sorted(pairs.items(), key=lambda x: (-x[1], x[0]))]
        trip_list = [(" + ".join(k), v) for k, v in sorted(trips.items(), key=lambda x: (-x[1], x[0]))]
        return sing_list, pair_list, trip_list
