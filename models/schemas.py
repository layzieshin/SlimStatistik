from dataclasses import dataclass
from typing import Optional

@dataclass(frozen=True)
class SampleHeader:
    proben_nr: str
    order_ts: Optional[str]  # raw text from DB
    report_ts: Optional[str]

@dataclass(frozen=True)
class SampleLine:
    proben_nr: str
    analyte_code: str
    analyte_name: Optional[str]
    result_value: Optional[str]
    result_ts: Optional[str]
