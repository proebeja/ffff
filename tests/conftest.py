import os
import sys

import pytest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
TESTDATA = os.path.join(ROOT, "testdata")

from fdd.core.hausconvention import Hausconvention  # noqa: E402


def datei(name: str) -> str:
    return os.path.join(TESTDATA, name)


ECKART = "Testdaten_Eckart_SuSa_2022-2025-03.xlsx"
SAP = "Testdaten_SAP_BK4756_HGB_BilGuV_2024.xlsx"
HUCHTEMEIER = "Testdaten_Huchtemeier_Kontennachweis_2024.pdf"
NAMUR = "Testdaten_Namur_Databook.xlsx"
ALLE = [ECKART, SAP, HUCHTEMEIER, NAMUR]


@pytest.fixture(scope="session")
def hc() -> Hausconvention:
    return Hausconvention.laden()
