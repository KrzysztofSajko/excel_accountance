from __future__ import annotations

from enum import Enum, auto
from typing import Optional

FEMALE_DESC = ["Kobieta"]
MALE_DESC = ["Mężczyzna"]


class Gender(Enum):
    male = auto()
    female = auto()

    @classmethod
    def from_string(cls, gender_desc: str) -> Optional[Gender]:
        if gender_desc in FEMALE_DESC:
            return cls.female
        if gender_desc in MALE_DESC:
            return cls.male

    def __str__(self):
        if self == Gender.male:
            return "Mężczyzna"
        if self == Gender.female:
            return "Kobieta"
