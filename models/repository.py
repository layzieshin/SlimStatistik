import os
import sqlite3
from typing import List, Tuple, Dict, Optional
import itertools


def _best_effort_decode(b):
    if b is None or isinstance(b, str):
        return b
    for enc in ("utf-8", "cp1252", "latin-1"):
        try:
            return b.decode(enc)
        except Exception:
            pass
    return b.decode("utf-8", errors="replace")


class Repository:
    """
    Dünne DB-Schicht (reine SQL-Queries).
    Tabellen:
      Befund(ProbenNr, TimeStamp, AbnahmeDatum, Name, Vname, GebDat, PatID, AuftragsNr,
             EinsenderInfo, EinsenderKennung, ...)
      BefTag(ProbenNr, TestKB, Ergebnis, ...)
    """
    def __init__(self, db_path: str):
        self.db_path = db_path

    def _available(self) -> bool:
        return bool(self.db_path) and os.path.exists(self.db_path)

    def _conn(self) -> sqlite3.Connection:
        con = sqlite3.connect(self.db_path)
        con.row_factory = sqlite3.Row
        con.text_factory = _best_effort_decode
        return con

    # --------- Analyten
    def list_all_analytes(self) -> List[str]:
        if not self._available():
            return []
        q = "SELECT DISTINCT TestKB FROM BefTag WHERE TestKB IS NOT NULL AND TRIM(TestKB) <> '' ORDER BY TestKB"
        with self._conn() as con:
            return [r[0] for r in con.execute(q).fetchall()]

    # --------- Zählungen
    def count_requirements_per_analyte(self, analytes: List[str], start: str, end: str) -> List[Tuple[str, int]]:
        if not self._available() or not analytes:
            return []
        placeholders = ",".join("?" for _ in analytes)
        q = f"""
        SELECT t.TestKB, COUNT(*) AS cnt
        FROM BefTag t
        JOIN Befund b ON b.ProbenNr = t.ProbenNr
        WHERE t.TestKB IN ({placeholders})
          AND COALESCE(b.AbnahmeDatum, b.TimeStamp) >= ?
          AND COALESCE(b.AbnahmeDatum, b.TimeStamp) <= ?
        GROUP BY t.TestKB
        ORDER BY t.TestKB
        """
        params = analytes + [start, end]
        with self._conn() as con:
            rows = con.execute(q, params).fetchall()
            return [(r["TestKB"], int(r["cnt"])) for r in rows]

    def count_befund_status(self, start: str, end: str) -> Tuple[int, int, int]:
        if not self._available():
            return 0, 0, 0
        q_total = """
        SELECT COUNT(*) AS c
        FROM Befund b
        WHERE COALESCE(b.AbnahmeDatum, b.TimeStamp) >= ?
          AND COALESCE(b.AbnahmeDatum, b.TimeStamp) <= ?
        """
        q_open = """
        SELECT COUNT(DISTINCT b.ProbenNr) AS c
        FROM Befund b
        JOIN BefTag t ON t.ProbenNr = b.ProbenNr
        WHERE COALESCE(b.AbnahmeDatum, b.TimeStamp) >= ?
          AND COALESCE(b.AbnahmeDatum, b.TimeStamp) <= ?
          AND t.Ergebnis IS NULL
        """
        q_done = """
        SELECT COUNT(*) AS c FROM (
            SELECT b.ProbenNr
            FROM Befund b
            JOIN BefTag t ON t.ProbenNr = b.ProbenNr
            WHERE COALESCE(b.AbnahmeDatum, b.TimeStamp) >= ?
              AND COALESCE(b.AbnahmeDatum, b.TimeStamp) <= ?
            GROUP BY b.ProbenNr
            HAVING SUM(CASE WHEN t.Ergebnis IS NULL THEN 1 ELSE 0 END) = 0
        ) x
        """
        with self._conn() as con:
            total = con.execute(q_total, (start, end)).fetchone()["c"]
            opened = con.execute(q_open, (start, end)).fetchone()["c"]
            done = con.execute(q_done, (start, end)).fetchone()["c"]
        return int(total), int(opened), int(done)

    def count_befunde_per_weekday(self, start: str, end: str, only_open: bool, analytes=None) -> Dict[str, int]:
        if not self._available():
            return {"Mo": 0, "Di": 0, "Mi": 0, "Do": 0, "Fr": 0, "Sa": 0, "So": 0}
        if only_open:
            q = """
            SELECT STRFTIME('%w', COALESCE(b.AbnahmeDatum, b.TimeStamp)) AS wd,
                   COUNT(DISTINCT b.ProbenNr) AS c
            FROM Befund b
            JOIN BefTag t ON t.ProbenNr = b.ProbenNr
            WHERE COALESCE(b.AbnahmeDatum, b.TimeStamp) >= ?
              AND COALESCE(b.AbnahmeDatum, b.TimeStamp) <= ?
              AND t.Ergebnis IS NULL
            GROUP BY wd
            """
            params = (start, end)
        else:
            q = """
            SELECT STRFTIME('%w', COALESCE(b.AbnahmeDatum, b.TimeStamp)) AS wd,
                   COUNT(*) AS c
            FROM Befund b
            WHERE COALESCE(b.AbnahmeDatum, b.TimeStamp) >= ?
              AND COALESCE(b.AbnahmeDatum, b.TimeStamp) <= ?
            GROUP BY wd
            """
            params = (start, end)

        wd_map = {"0": "So", "1": "Mo", "2": "Di", "3": "Mi", "4": "Do", "5": "Fr", "6": "Sa"}
        out: Dict[str, int] = {"Mo": 0, "Di": 0, "Mi": 0, "Do": 0, "Fr": 0, "Sa": 0, "So": 0}
        with self._conn() as con:
            for r in con.execute(q, params):
                out[wd_map.get(r["wd"], "?")] = int(r["c"])
        return out

    # --------- Offene Anforderungen
    def count_open_requirements_per_analyte(self, analytes: List[str], since: str) -> List[Tuple[str, int]]:
        if not self._available() or not analytes:
            return []
        placeholders = ",".join("?" for _ in analytes)
        q = f"""
        SELECT t.TestKB, COUNT(*) AS cnt
        FROM BefTag t
        JOIN Befund b ON b.ProbenNr = t.ProbenNr
        WHERE t.TestKB IN ({placeholders})
          AND t.Ergebnis IS NULL
          AND COALESCE(b.AbnahmeDatum, b.TimeStamp) >= ?
        GROUP BY t.TestKB
        ORDER BY t.TestKB
        """
        params = analytes + [since]
        with self._conn() as con:
            rows = con.execute(q, params).fetchall()
            return [(r["TestKB"], int(r["cnt"])) for r in rows]

    # --------- Nicht entnommen?
    def list_suspected_missing_draw(self, older_than_hours: int = 24) -> List[Dict]:
        if not self._available():
            return []
        q = f"""
        SELECT b.ProbenNr,
               COALESCE(b.AbnahmeDatum, b.TimeStamp) AS ts,
               COUNT(t.TestKB) AS num_req,
               REPLACE(GROUP_CONCAT(DISTINCT t.TestKB), ',', ', ') AS analytes
        FROM Befund b
        JOIN BefTag t ON t.ProbenNr = b.ProbenNr
        GROUP BY b.ProbenNr
        HAVING num_req >= 1
           AND SUM(CASE WHEN t.Ergebnis IS NULL THEN 0 ELSE 1 END) = 0
           AND ts <= DATETIME('now', '-{older_than_hours} hours')
        ORDER BY ts ASC
        """
        with self._conn() as con:
            rows = con.execute(q).fetchall()
            return [{
                "ProbenNr": r["ProbenNr"],
                "OrderTime": r["ts"],
                "NumReq": int(r["num_req"]),
                "Analytes": r["analytes"] or "",
            } for r in rows]

    def get_sample_audit_info(self, proben_nr: str) -> Optional[Dict]:
        if not self._available():
            return None
        q = """
        SELECT b.ProbenNr,
               COALESCE(b.AbnahmeDatum, b.TimeStamp) AS Abnahme,
               b.AuftragsNr,
               COALESCE(NULLIF(b.EinsenderInfo, ''), b.EinsenderKennung) AS Einsender,
               b.Name, b.Vname, b.GebDat, b.PatID,
               REPLACE(
                   (SELECT GROUP_CONCAT(DISTINCT t.TestKB)
                    FROM BefTag t WHERE t.ProbenNr=b.ProbenNr),
                   ',', ', '
               ) AS Analyte
        FROM Befund b
        WHERE b.ProbenNr = ?
        """
        with self._conn() as con:
            r = con.execute(q, (proben_nr,)).fetchone()
            if not r:
                return None
            return {
                "ProbenNr": r["ProbenNr"],
                "Abnahme": r["Abnahme"],
                "AuftragsNr": r["AuftragsNr"],
                "Einsender": r["Einsender"],
                "Name": r["Name"],
                "Vname": r["Vname"],
                "GebDat": r["GebDat"],
                "PatID": r["PatID"],
                "Analyte": r["Analyte"] or "",
            }

    def delete_sample(self, proben_nr: str) -> int:
        return self.delete_samples([proben_nr])

    def delete_samples(self, proben_nrs: List[str]) -> int:
        if not self._available() or not proben_nrs:
            return 0
        placeholders = ",".join("?" for _ in proben_nrs)
        with self._conn() as con:
            c1 = con.execute(f"DELETE FROM BefTag WHERE ProbenNr IN ({placeholders})", proben_nrs).rowcount
            c2 = con.execute(f"DELETE FROM Befund WHERE ProbenNr IN ({placeholders})", proben_nrs).rowcount
            con.commit()
            return int(c1 + c2)

    # --------- Singlets / Kombinationen (1–4)
    import itertools
    # ... Rest unverändert ...

    def open_combo_stats(self, since: str, excluded: Optional[set] = None, max_k: int = 4):
        """
        EXAKT-Größen-Logik:
          - 1er: nur Proben mit GENAU 1 offenem Analyt
          - 2er/3er/4er: nur Proben mit GENAU 2/3/4 offenen Analyten
        WICHTIG: 'excluded' wird HIER NICHT angewendet, um die reale offene Matrix
                 der Probe nicht zu verfälschen (keine Teilmengenbildung).
        """
        if not self._available():
            return {}, {}, {}, {}

        q = """
        SELECT b.ProbenNr, GROUP_CONCAT(DISTINCT t.TestKB) AS ks
        FROM Befund b
        JOIN BefTag t ON t.ProbenNr = b.ProbenNr
        WHERE t.Ergebnis IS NULL
          AND COALESCE(b.AbnahmeDatum, b.TimeStamp) >= ?
        GROUP BY b.ProbenNr
        """

        singles: Dict[str, int] = {}
        pairs: Dict[str, int] = {}
        trips: Dict[str, int] = {}
        quads: Dict[str, int] = {}

        max_k = max(1, min(4, int(max_k)))

        with self._conn() as con:
            for r in con.execute(q, (since,)):
                raw = (r["ks"] or "")
                # deduplizieren + sortieren
                ks = sorted({s.strip() for s in raw.split(",") if s and s.strip()})
                n = len(ks)
                if n == 0:
                    continue

                if n == 1 and max_k >= 1:
                    singles[ks[0]] = singles.get(ks[0], 0) + 1
                elif n == 2 and max_k >= 2:
                    key = " + ".join(ks)
                    pairs[key] = pairs.get(key, 0) + 1
                elif n == 3 and max_k >= 3:
                    key = " + ".join(ks)
                    trips[key] = trips.get(key, 0) + 1
                elif n == 4 and max_k >= 4:
                    key = " + ".join(ks)
                    quads[key] = quads.get(key, 0) + 1

        return singles, pairs, trips, quads
