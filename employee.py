from __future__ import annotations

import os
from dataclasses import dataclass, field
from typing import Optional

from gender import Gender
from summary import MonthlySummary


@dataclass
class Employee:
    name: str
    last_name: str
    gender: Gender
    monthly_summaries: dict[str, MonthlySummary] = field(default_factory=dict)

    def find_summary(self, directory: str) -> Optional[str]:
        for filename in os.listdir(directory):
            if filename.endswith(".xlsx"):
                *last_name, name = tuple(map(lambda part: part.capitalize(), filename.split(".")[0].split()))
                last_name = " ".join(last_name)
                if last_name == self.last_name and name == self.name:
                    return os.path.join(directory, filename)
