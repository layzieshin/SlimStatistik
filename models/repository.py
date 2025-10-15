import sqlite3
from typing import Iterable, List, Dict, Any, Optional, Set
import json, datetime

class Repository:
    """
    SQLite Repository mit konfigurierbarem Feld-Mapping (config/mapping.json).

    Datenmodell:
      - Kopf:   Befund (1)  --ProbenNr-->  (n) BefTag : Zeilen/Anforderungen
      - Zeitraum-/Orderzeitfilter laufen über Spalten aus Befund (AbnahmeDatum etc.)
      - Status nur über BefTag.Ergebnis:
          * done = alle Zeilen haben Ergebnis (nicht NULL/leer)
          * open = mindestens eine Zeile hat kein Ergebnis
      - 'Nicht entnommen?': Proben mit >1 Anforderungen, in keiner Zeile ein Ergebnis, >24h alt
        → globaler Ausschluss + separate Liste/Löschfunktion
    """

    def __init__(self, db_path: str, mapping_path: str):
        self.db_path = db_path
        self.mapping = self._load_mapping(mapping_path)

    @staticmethod
    def _load_mapping(path: str) -> Dict[str, Any]:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)

    def _conn(self) -> sqlite3.Connection:
        con = sqlite3.connect(self.db_path)
        con.row_factory = sqlite3.Row
        return con

    # -------------------- Date/Time Helpers --------------------

    def _order_ts_candidates(self) -> List[str]:
        return list(self.mapping["header_fields"]["order_timestamp"])

    def _normalize_dt_sql(self, col: str) -> str:
        """
        Normalisiert h."col" robust nach datetime.

        Unterstützt u. a.:
        - ISO 'YYYY-MM-DD HH:MM:SS' (+ Millisekunden, egal ob '.' oder ',' als Separator)
        - ISO mit 'T'
        - 14-stellig 'YYYYMMDDHHMMSS'
        - 8-stellig  'YYYYMMDD'
        - Deutsch 'DD.MM.YYYY[ HH:MM:SS]'
        - Unix s/ms, Excel-Serienzahl, FILETIME, .NET-Ticks
        """
        # Vorbehandlung: 'T' → ' ', Komma in Millis → Punkt
        cleaned = f"replace(replace(h.\"{col}\",'T',' '), ',', '.')"

        # 1) Vollformat mit evtl. Millisekunden (SQLite kann fractional seconds -> datetime() + julianday())
        # 2) Truncate auf 19 Zeichen (falls exotische Suffixe vorhanden sind)
        iso_full   = f"datetime({cleaned})"
        iso_19     = f"datetime(substr({cleaned},1,19))"

        # 3) 14-stellig YYYYMMDDHHMMSS
        ymdhms_14 = (
            f"substr(h.\"{col}\",1,4) || '-' || substr(h.\"{col}\",5,2) || '-' || substr(h.\"{col}\",7,2) || "
            f"' ' || substr(h.\"{col}\",9,2) || ':' || substr(h.\"{col}\",11,2) || ':' || substr(h.\"{col}\",13,2)"
        )
        # 4) 8-stellig YYYYMMDD
        ymd_8 = f"substr(h.\"{col}\",1,4) || '-' || substr(h.\"{col}\",5,2) || '-' || substr(h.\"{col}\",7,2) || ' 00:00:00'"

        # 5) Deutsch dd.mm.yyyy[ HH:MM:SS]
        de_date = f"substr(h.\"{col}\",7,4) || '-' || substr(h.\"{col}\",4,2) || '-' || substr(h.\"{col}\",1,2)"
        de_time = f"CASE WHEN length(h.\"{col}\")>=19 THEN ' '||substr(h.\"{col}\",12,8) ELSE ' 00:00:00' END"

        # Numerische Varianten
        unix_s = (
            f"CASE WHEN typeof(h.\"{col}\") IN ('integer','real') "
            f"     AND h.\"{col}\" BETWEEN 0 AND 32503680000 "
            f"THEN datetime(h.\"{col}\",'unixepoch') END"
        )
        unix_ms = (
            f"CASE WHEN typeof(h.\"{col}\") IN ('integer','real') "
            f"     AND h.\"{col}\" BETWEEN 100000000000 AND 32503680000000 "
            f"THEN datetime(h.\"{col}\"/1000,'unixepoch') END"
        )
        excel_serial = (
            f"CASE WHEN typeof(h.\"{col}\") IN ('integer','real') "
            f"     AND h.\"{col}\" BETWEEN 20000 AND 60000 "
            f"THEN datetime( (h.\"{col}\"-25569)*86400, 'unixepoch') END"
        )
        filetime = (
            f"CASE WHEN typeof(h.\"{col}\") IN ('integer','real') "
            f"     AND h.\"{col}\" > 11644473600*10000000 "
            f"     AND h.\"{col}\" < 32503680000*10000000 "
            f"THEN datetime( (h.\"{col}\"/10000000)-11644473600, 'unixepoch') END"
        )
        dotnet_ticks = (
            f"CASE WHEN typeof(h.\"{col}\") IN ('integer','real') "
            f"     AND h.\"{col}\" > 62135596800*10000000 "
            f"     AND h.\"{col}\" < 32503680000*10000000 "
            f"THEN datetime( (h.\"{col}\"/10000000)-62135596800, 'unixepoch') END"
        )

        return (
            "COALESCE("
            f"  {iso_full},"                                # z. B. 2024-03-14 11:46:14.163
            f"  {iso_19},"                                  # auf 19 Zeichen gekappt
            f"  CASE WHEN length(h.\"{col}\")=14 AND h.\"{col}\" NOT LIKE '%-%' THEN datetime({ymdhms_14}) END,"
            f"  CASE WHEN length(h.\"{col}\")=8  AND h.\"{col}\" NOT LIKE '%-%' THEN datetime({ymd_8}) END,"
            f"  CASE WHEN instr(h.\"{col}\",'.')>0 AND length(h.\"{col}\")>=10 THEN datetime({de_date} || {de_time}) END,"
            f"  {unix_s}, {unix_ms}, {excel_serial}, {filetime}, {dotnet_ticks}"
            ")"
        )

    def _order_ts_sql(self) -> str:
        frags = [self._normalize_dt_sql(col) for col in self._order_ts_candidates()]
        return f"COALESCE({', '.join(frags)})"

    def _order_ts_jd(self) -> str:
        return f"julianday({self._order_ts_sql()})"

    # -------------------- Business Queries --------------------

    def suspected_missing_blood_draw_proben(self, now: Optional[datetime.datetime] = None) -> Set[str]:
        """
        Proben älter als 24h, >1 Anforderungen, KEIN Ergebnis in irgendeiner Zeile.
        → Werden global aus allen Statistiken ausgeschlossen.
        """
        now = now or datetime.datetime.now()
        header = self.mapping["tables"]["header"]
        lines = self.mapping["tables"]["lines"]
        key = self.mapping["keys"]["sample_id"]
        result_val = self.mapping["line_fields"]["result_value"]
        order_jd = self._order_ts_jd()

        sql = f"""
        SELECT h."{key}" AS ProbenNr
        FROM "{lines}" l
        JOIN "{header}" h ON h."{key}" = l."{key}"
        WHERE {order_jd} IS NOT NULL
        GROUP BY h."{key}"
        HAVING COUNT(*) > 1
           AND SUM(CASE WHEN l."{result_val}" IS NOT NULL AND TRIM(l."{result_val}") != '' THEN 1 ELSE 0 END) = 0
           AND (julianday(?) - {order_jd}) > 1.0
        """
        with self._conn() as con:
            res = con.execute(sql, (now.strftime("%Y-%m-%d %H:%M:%S"),)).fetchall()
        return {row["ProbenNr"] for row in res}

    def count_analyte_requests(self, analyte_code: str, start: str, end: str, exclude_proben: Iterable[str]) -> int:
        header = self.mapping["tables"]["header"]
        lines = self.mapping["tables"]["lines"]
        key = self.mapping["keys"]["sample_id"]
        analyte_col = self.mapping["line_fields"]["analyte_code"]
        order_jd = self._order_ts_jd()
        excl = tuple(exclude_proben)
        excl_clause = f"AND h.\"{key}\" NOT IN ({','.join(['?']*len(excl))})" if excl else ""
        sql = f"""
        SELECT COUNT(*)
        FROM "{lines}" l
        JOIN "{header}" h ON h."{key}" = l."{key}"
        WHERE l."{analyte_col}" = ?
          AND {order_jd} BETWEEN julianday(?) AND julianday(?)
          {excl_clause}
        """
        params = [analyte_code, start, end] + list(excl)
        with self._conn() as con:
            row = con.execute(sql, params).fetchone()
            return int(row[0] or 0)

    def per_sample_completion(self, start: str, end: str, exclude_proben: Iterable[str]) -> Dict[str, str]:
        """Status je Probe (open/done) – nur BefTag.Ergebnis zählt."""
        header = self.mapping["tables"]["header"]
        lines = self.mapping["tables"]["lines"]
        key = self.mapping["keys"]["sample_id"]
        result_val = self.mapping["line_fields"]["result_value"]
        order_jd = self._order_ts_jd()
        excl = tuple(exclude_proben)
        excl_clause = f"AND h.\"{key}\" NOT IN ({','.join(['?']*len(excl))})" if excl else ""
        sql = f"""
        SELECT h."{key}" AS ProbenNr,
               SUM(CASE WHEN (l."{result_val}" IS NOT NULL AND TRIM(l."{result_val}") != '') THEN 1 ELSE 0 END) AS with_result,
               COUNT(*) AS total
        FROM "{lines}" l
        JOIN "{header}" h ON h."{key}" = l."{key}"
        WHERE {order_jd} BETWEEN julianday(?) AND julianday(?)
          {excl_clause}
        GROUP BY h."{key}"
        """
        params = [start, end] + list(excl)
        out: Dict[str, str] = {}
        with self._conn() as con:
            for r in con.execute(sql, params):
                out[r["ProbenNr"]] = "done" if int(r["with_result"]) == int(r["total"]) else "open"
        return out

    def weekday_counts(
        self,
        status_filter: Optional[str],
        analyte: Optional[str],
        start: str,
        end: str,
        exclude_proben: Iterable[str]
    ) -> Dict[int, int]:
        """Aggregiert pro Wochentag (0=Mo..6=So) die Anzahl Proben, optional Analyt/Status-Filter."""
        header = self.mapping["tables"]["header"]
        lines = self.mapping["tables"]["lines"]
        key = self.mapping["keys"]["sample_id"]
        result_val = self.mapping["line_fields"]["result_value"]
        analyte_col = self.mapping["line_fields"]["analyte_code"]
        order_jd = self._order_ts_jd()
        order_date = f"date({self._order_ts_sql()})"
        excl = tuple(exclude_proben)
        excl_clause = f"AND h.\"{key}\" NOT IN ({','.join(['?']*len(excl))})" if excl else ""
        analyte_clause = f"AND l.\"{analyte_col}\" = ?" if analyte else ""
        sql = f"""
        WITH sample AS (
            SELECT h."{key}" AS ProbenNr,
                   {order_date} AS d,
                   SUM(CASE WHEN (l."{result_val}" IS NOT NULL AND TRIM(l."{result_val}") != '') THEN 1 ELSE 0 END) AS with_result,
                   COUNT(*) AS total
            FROM "{lines}" l
            JOIN "{header}" h ON h."{key}" = l."{key}"
            WHERE {order_jd} BETWEEN julianday(?) AND julianday(?)
              {excl_clause}
              {analyte_clause}
            GROUP BY h."{key}", {order_date}
        )
        SELECT strftime('%w', d) AS weekday, COUNT(*) AS c
        FROM sample
        WHERE (? IS NULL) OR
              (? = 'open' AND with_result < total) OR
              (? = 'done' AND with_result = total) OR
              (? = 'all')
        GROUP BY weekday
        """
        params = [start, end] + list(excl)
        if analyte:
            params += [analyte]
        params += [status_filter, status_filter, status_filter, status_filter]
        out = {i: 0 for i in range(7)}
        with self._conn() as con:
            for r in con.execute(sql, params):
                wd_sql = int(r["weekday"])   # 0=So..6=Sa
                wd = (wd_sql + 6) % 7        # -> 0=Mo..6=So
                out[wd] = int(r["c"] or 0)
        return out

    # ---------- NEU: Offene Anforderungen je Analyt (ab Datum X) ----------

    def open_requirements_counts_since(
        self, analytes: Iterable[str], since_inclusive: str, exclude_proben: Iterable[str]
    ) -> List[Dict[str, Any]]:
        """
        Gibt pro Analyt (TestKB) die Anzahl offener Anforderungen (ohne Ergebnis)
        für alle Proben mit Orderzeit >= since_inclusive zurück.
        """
        header = self.mapping["tables"]["header"]
        lines = self.mapping["tables"]["lines"]
        key = self.mapping["keys"]["sample_id"]
        result_val = self.mapping["line_fields"]["result_value"]
        analyte_col = self.mapping["line_fields"]["analyte_code"]
        order_jd = self._order_ts_jd()

        excl = tuple(exclude_proben)
        excl_clause = f"AND h.\"{key}\" NOT IN ({','.join(['?']*len(excl))})" if excl else ""

        analytes = list(analytes)
        placeholders = ",".join(["?"] * len(analytes)) if analytes else "NULL"

        sql = f"""
        SELECT l."{analyte_col}" AS code, COUNT(*) AS c
        FROM "{lines}" l
        JOIN "{header}" h ON h."{key}" = l."{key}"
        WHERE l."{analyte_col}" IN ({placeholders})
          AND (l."{result_val}" IS NULL OR TRIM(l."{result_val}") = '')
          AND {order_jd} >= julianday(?)
          {excl_clause}
        GROUP BY l."{analyte_col}"
        ORDER BY l."{analyte_col}"
        """
        params = list(analytes) + [since_inclusive] + list(excl)
        with self._conn() as con:
            rows = con.execute(sql, params).fetchall()
            return [dict(r) for r in rows]

    # -----------------------------------------------------------------------

    def list_suspected_missing_blood_draw(self) -> List[Dict[str, Any]]:
        """Details zu verdächtigen Proben (Definition s. suspected_missing_blood_draw_proben)."""
        header = self.mapping["tables"]["header"]
        lines = self.mapping["tables"]["lines"]
        key = self.mapping["keys"]["sample_id"]
        result_val = self.mapping["line_fields"]["result_value"]
        analyte_col = self.mapping["line_fields"]["analyte_code"]
        order_jd = self._order_ts_jd()
        order_ts = self._order_ts_sql()

        sql = f"""
        SELECT h."{key}" AS ProbenNr,
               {order_ts} AS OrderTime,
               COUNT(*) AS NumReq,
               GROUP_CONCAT(l."{analyte_col}", ',') AS Analytes
        FROM "{lines}" l
        JOIN "{header}" h ON h."{key}" = l."{key}"
        WHERE {order_jd} IS NOT NULL
        GROUP BY h."{key}"
        HAVING NumReq > 1
           AND SUM(CASE WHEN l."{result_val}" IS NOT NULL AND TRIM(l."{result_val}") != '' THEN 1 ELSE 0 END) = 0
           AND (julianday('now') - {order_jd}) > 1.0
        ORDER BY OrderTime ASC
        """
        with self._conn() as con:
            rows = con.execute(sql).fetchall()
            return [dict(r) for r in rows]

    def delete_sample(self, proben_nr: str) -> int:
        """Löscht `Befund` + zugehörige `BefTag`-Zeilen. Rückgabe: Anzahl gelöschter Zeilen gesamt."""
        header = self.mapping["tables"]["header"]
        lines = self.mapping["tables"]["lines"]
        key = self.mapping["keys"]["sample_id"]
        with self._conn() as con:
            cur = con.cursor()
            n1 = cur.execute(f'DELETE FROM "{lines}" WHERE "{key}" = ?', (proben_nr,)).rowcount
            n2 = cur.execute(f'DELETE FROM "{header}" WHERE "{key}" = ?', (proben_nr,)).rowcount
            con.commit()
        return int((n1 or 0) + (n2 or 0))
