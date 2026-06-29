"""Property 15: Unparseable fields retain raw text, are flagged, and are not removed.

*For any* line-item field whose printed text cannot be parsed into its expected
data type, the raw extracted text is preserved (never dropped or rewritten), the
field is marked for User review, and neither the field nor the line is removed.

**Validates: Requirements 9.2, 9.3**

The property is checked from two complementary angles, mirroring how the parser
threads an unparseable value through to the output:

1. **Capture level** -- ``PdfParser._capture(text, numeric=True)`` on a non-numeric
   string yields ``is_unparseable=True`` with ``raw_text`` preserved byte-for-byte
   and ``parsed=None`` (no inferred/substituted value).
2. **Merge level** -- ``PdfParser._merge_items`` keeps any unparseable ``RawValue``
   on the merged ``LineItem`` (the field is still present, carrying its raw text)
   *and* emits an ``UNPARSEABLE`` ``ReviewFlag`` for that serial/field carrying the
   same raw text. The line item itself is never dropped.
"""

from __future__ import annotations

import string

import hypothesis.strategies as st
from hypothesis import given

from boe_converter.models import RawValue
from boe_converter.parser import DutyItemRow, InvoiceItemRow, PdfParser

# Required line-item fields sourced from Part II (invoice) and Part III (duty),
# matching the join in ``_merge_items`` (Req 3.4-3.12). ``item_serial`` is the
# always-present int join key and so is not a flaggable field.
_INVOICE_FIELDS = ("cth_hsn", "description", "unit_price_usd", "quantity", "unit")
_DUTY_FIELDS = (
    "assessable_value",
    "bcd_rate",
    "bcd_amount",
    "igst_rate",
    "total_duty",
)
_ALL_FIELDS = _INVOICE_FIELDS + _DUTY_FIELDS


@st.composite
def _unparseable_text(draw) -> str:
    """Draw a non-blank string that ``_capture(numeric=True)`` cannot parse.

    ``_parse_number`` strips grouping commas, currency symbols and a trailing
    percent sign before matching ``[+-]?\\d*\\.?\\d+``. An embedded ASCII *letter*
    survives that cleaning and can never be part of a number, so appending one
    guarantees the text is located-but-unparseable while still exercising a wide
    variety of surrounding characters (digits, symbols, unicode, whitespace).
    """
    base = draw(st.text(min_size=0, max_size=20))
    letter = draw(st.sampled_from(string.ascii_letters))
    # Place the guaranteeing letter somewhere inside the string for variety.
    pos = draw(st.integers(min_value=0, max_value=len(base)))
    return base[:pos] + letter + base[pos:]


# ---------------------------------------------------------------------------
# Angle 1: capture level
# ---------------------------------------------------------------------------


@given(text=_unparseable_text())
def test_unparseable_capture_retains_raw_text(text):
    """A non-numeric value captured numerically is flagged, raw text intact.

    ``raw_text`` equals the source characters exactly, ``parsed`` is ``None`` (no
    substituted value), ``is_unparseable`` is set, and the field is *not* treated
    as missing -- it was located, just not resolvable as a number.

    **Validates: Requirements 9.2, 9.3**
    """
    rv = PdfParser()._capture(text, numeric=True)

    assert rv.is_unparseable is True
    assert rv.is_missing is False
    assert rv.raw_text == text          # verbatim, no truncation/reformatting
    assert len(rv.raw_text) == len(text)
    assert rv.parsed is None            # nothing inferred or substituted


# ---------------------------------------------------------------------------
# Angle 2: merge level
# ---------------------------------------------------------------------------


@st.composite
def _merge_scenario(draw):
    """Build (inv, duty, expected) where some fields carry unparseable RawValues.

    Every serial is present in *both* sources so all ten required fields exist on
    the merged ``LineItem``; an arbitrary non-empty subset of those fields is made
    unparseable (the rest are valid). ``expected`` maps each serial to the
    ``{field_name: raw_text}`` of its unparseable fields, so the test can assert
    both retention on the ``LineItem`` and the emitted ``UNPARSEABLE`` flags.
    """
    serials = draw(
        st.lists(
            st.integers(min_value=1, max_value=300),
            min_size=1,
            max_size=6,
            unique=True,
        )
    )

    inv: dict[int, InvoiceItemRow] = {}
    duty: dict[int, DutyItemRow] = {}
    expected: dict[int, dict[str, str]] = {}

    for serial in serials:
        # Choose a non-empty subset of fields to be unparseable for this serial.
        unparseable_fields = draw(
            st.lists(st.sampled_from(_ALL_FIELDS), min_size=1, max_size=len(_ALL_FIELDS), unique=True)
        )
        field_values: dict[str, RawValue] = {}
        expected_for_serial: dict[str, str] = {}

        for name in _ALL_FIELDS:
            if name in unparseable_fields:
                raw = draw(_unparseable_text())
                field_values[name] = RawValue.unparseable(raw)
                expected_for_serial[name] = raw
            else:
                # A clean, parseable value (no flag expected).
                field_values[name] = RawValue(raw_text="1", parsed=1.0)

        expected[serial] = expected_for_serial

        inv[serial] = InvoiceItemRow(
            item_serial=serial,
            cth_hsn=field_values["cth_hsn"],
            description=field_values["description"],
            unit_price_usd=field_values["unit_price_usd"],
            quantity=field_values["quantity"],
            unit=field_values["unit"],
            amount=RawValue(raw_text="1", parsed=1.0),  # not a merged required field
        )
        duty[serial] = DutyItemRow(
            item_serial=serial,
            assessable_value=field_values["assessable_value"],
            bcd_rate=field_values["bcd_rate"],
            bcd_amount=field_values["bcd_amount"],
            sws_amount=RawValue(raw_text="0", parsed=0.0),  # not a merged required field
            igst_rate=field_values["igst_rate"],
            total_duty=field_values["total_duty"],
        )

    return inv, duty, expected


@given(scenario=_merge_scenario())
def test_unparseable_fields_retained_flagged_and_not_removed(scenario):
    """Merged line items keep unparseable raw text and emit UNPARSEABLE flags.

    For each serial and each unparseable field:
    - the ``LineItem`` still carries that field, with its ``RawValue`` retaining
      the verbatim ``raw_text`` and ``is_unparseable=True`` (field not removed);
    - exactly one ``UNPARSEABLE`` ``ReviewFlag`` is emitted for that serial/field,
      carrying the same raw text;
    and no line item is dropped.

    **Validates: Requirements 9.2, 9.3**
    """
    inv, duty, expected = scenario
    parser = PdfParser()

    line_items, flags = parser._merge_items(inv, duty, declared_count=None)

    # No line item is removed: every serial appears exactly once.
    by_serial = {li.item_serial: li for li in line_items}
    assert set(by_serial) == set(expected)
    assert len(line_items) == len(expected)

    # Unparseable raw text is retained on each LineItem field (not removed).
    for serial, fields in expected.items():
        line = by_serial[serial]
        for field_name, raw in fields.items():
            rv = getattr(line, field_name)
            assert rv.is_unparseable is True
            assert rv.raw_text == raw            # verbatim, retained
            assert rv.parsed is None             # no substituted value

    # Exactly the expected UNPARSEABLE flags are emitted, each carrying raw text.
    unparseable_flags = [
        f for f in flags if f.scope == "line_item" and f.reason == "UNPARSEABLE"
    ]
    actual = {(f.item_serial, f.field_name, f.raw_text) for f in unparseable_flags}
    expected_flagset = {
        (serial, field_name, raw)
        for serial, fields in expected.items()
        for field_name, raw in fields.items()
    }
    assert actual == expected_flagset
    assert len(unparseable_flags) == len(expected_flagset), "each flag emitted once"
