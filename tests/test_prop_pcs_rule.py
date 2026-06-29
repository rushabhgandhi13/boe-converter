"""Property-based test for the ``pcs`` rule (task 2.3).

Property 6: pcs is QTY×12 exactly when the unit is DOZ, otherwise blank.

**Validates: Requirements 5.8, 5.9**

For any line item, if the unit string with leading/trailing whitespace trimmed
and case ignored equals ``"DOZ"`` then ``pcs == qty * 12``; otherwise the
``pcs`` cell is blank (``None``).

The generators below produce both kinds of unit strings:

* "DOZ-equivalent" units: the three letters ``DOZ`` with arbitrary letter
  casing and arbitrary surrounding whitespace (spaces, tabs, newlines). Every
  such string must normalize to ``"DOZ"`` and yield ``pcs == qty * 12``.
* Non-DOZ units: arbitrary text (plus a few realistic UQC codes such as
  ``KGS``, ``NOS``, ``PCS``) explicitly filtered so they never normalize to
  ``"DOZ"``. Every such string must yield ``pcs is None``.

``qty * 12`` is asserted with exact equality because the calculator performs
exactly that multiplication at full floating-point precision.
"""

from __future__ import annotations

from hypothesis import assume, given
from hypothesis import strategies as st

from boe_converter.calculator import ValueCalculator
from boe_converter.models import LineItem, RawValue

# Whitespace characters allowed to surround the unit token; all must be removed
# by ``str.strip()`` so the DOZ rule still fires.
_WHITESPACE = " \t\n\r\f\v"


def _raw_num(value: float) -> RawValue:
    """A RawValue carrying a numeric ``parsed`` interpretation, as the parser emits."""

    return RawValue(raw_text=repr(value), parsed=value)


def _line_item(unit: RawValue, qty: float) -> LineItem:
    """Build a minimal LineItem; only ``unit`` and ``quantity`` drive ``pcs``."""

    return LineItem(
        item_serial=1,
        cth_hsn=RawValue(raw_text="39264049", parsed="39264049"),
        description=RawValue(raw_text="ITEM", parsed="ITEM"),
        unit_price_usd=_raw_num(1.0),
        quantity=_raw_num(qty),
        unit=unit,
        assessable_value=_raw_num(1.0),
        bcd_rate=_raw_num(0.0),
        bcd_amount=_raw_num(0.0),
        igst_rate=_raw_num(0.0),
        total_duty=_raw_num(0.0),
    )


# Finite, well-bounded quantities (negatives, zero, and fractions included) so
# ``qty * 12`` never overflows while still exercising the input space.
quantities = st.floats(
    min_value=-1e5,
    max_value=1e5,
    allow_nan=False,
    allow_infinity=False,
)

_surrounding_ws = st.text(alphabet=_WHITESPACE, max_size=4)


@st.composite
def doz_unit_strings(draw) -> str:
    """A string that normalizes to ``"DOZ"``: the letters DOZ in any casing,
    wrapped in arbitrary leading/trailing whitespace."""

    letters = [draw(st.sampled_from([lo, lo.upper()])) for lo in "doz"]
    core = "".join(letters)
    return draw(_surrounding_ws) + core + draw(_surrounding_ws)


# Realistic non-DOZ UQC codes plus arbitrary text, all filtered so none can
# normalize to "DOZ".
_OTHER_UQCS = ["KGS", "NOS", "PCS", "SET", "MTR", "LTR", "U", "BAG", "DOZEN", "DZ", ""]


@st.composite
def non_doz_unit_strings(draw) -> str:
    """A string that does NOT normalize to ``"DOZ"``."""

    text = draw(
        st.one_of(
            st.sampled_from(_OTHER_UQCS),
            st.text(max_size=8),
        )
    )
    assume(text.strip().upper() != "DOZ")
    return text


@given(unit_text=doz_unit_strings(), qty=quantities)
def test_pcs_is_qty_times_12_when_unit_is_doz(unit_text: str, qty: float):
    """Unit normalizing to DOZ => pcs == qty * 12 (Req 5.8)."""

    # Guard the generator's own invariant.
    assert unit_text.strip().upper() == "DOZ"

    unit = RawValue(raw_text=unit_text, parsed=unit_text)
    line = ValueCalculator().compute_line(_line_item(unit, qty), usd_rate=1.0)

    assert line.pcs == qty * 12


@given(unit_text=non_doz_unit_strings(), qty=quantities)
def test_pcs_is_blank_when_unit_is_not_doz(unit_text: str, qty: float):
    """Any unit not normalizing to DOZ => pcs is blank/None (Req 5.9)."""

    assert unit_text.strip().upper() != "DOZ"

    unit = RawValue(raw_text=unit_text, parsed=unit_text)
    line = ValueCalculator().compute_line(_line_item(unit, qty), usd_rate=1.0)

    assert line.pcs is None
