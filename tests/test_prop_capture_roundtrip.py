"""Property-based test for verbatim capture round-trip (task 4.2).

Property 3: Extraction preserves printed values verbatim (capture round-trip).

**Validates: Requirements 2.11, 3.6**

For any printed field value, the captured ``RawValue.raw_text`` equals the
source characters exactly -- no truncation, reformatting, rounding, or
inference. The non-destructive numeric parse populates ``parsed`` *alongside*
the raw text but never alters the stored characters, regardless of whether the
parse succeeds. Blank or ``None`` input yields a "missing" RawValue.

The generators below cover the input space called out by the task:

* numeric-looking strings with grouping commas, currency symbols, and trailing
  percent signs (e.g. ``"1,234.56"``, ``"$1000"``, ``"18%"``);
* plain integer/decimal numbers (including signed and fractional);
* blank / whitespace-only strings and ``None`` (the "missing" cases);
* arbitrary non-numeric text (descriptions, codes, wrapped multi-line text).

``PdfParser._capture`` is exercised in both modes -- ``numeric=False`` (the
default text capture) and ``numeric=True`` (numeric fields) -- because the
verbatim round-trip invariant must hold identically in both.
"""

from __future__ import annotations

from hypothesis import given
from hypothesis import strategies as st

from boe_converter.parser import PdfParser

# ---------------------------------------------------------------------------
# Generators for arbitrary printed text values
# ---------------------------------------------------------------------------

# Plain numbers as they might be printed: integers and decimals, signed too.
_plain_numbers = st.one_of(
    st.integers(min_value=-10_000_000, max_value=10_000_000).map(str),
    st.floats(
        min_value=-1e7,
        max_value=1e7,
        allow_nan=False,
        allow_infinity=False,
    ).map(lambda f: repr(f)),
)


@st.composite
def _grouped_numbers(draw) -> str:
    """A number printed with thousands-grouping commas, e.g. ``"1,234,567.89"``."""

    whole = draw(st.integers(min_value=0, max_value=999_999_999))
    grouped = f"{whole:,}"
    if draw(st.booleans()):
        frac = draw(st.integers(min_value=0, max_value=99))
        grouped = f"{grouped}.{frac:02d}"
    return grouped


# Numbers decorated with currency symbols and/or a trailing percent sign.
_currency = st.sampled_from(["", "₹", "$", "€", "£"])
_percent = st.sampled_from(["", "%"])


@st.composite
def _decorated_numbers(draw) -> str:
    """A numeric string wrapped with an optional currency symbol / percent sign."""

    base = draw(st.one_of(_plain_numbers, _grouped_numbers()))
    return draw(_currency) + base + draw(_percent)


# Blank-ish values that must be treated as "missing".
_blanks = st.sampled_from(["", " ", "   ", "\t", "\n", "  \t \n "])

# Arbitrary non-numeric text (descriptions, wrapped text, codes, unicode).
# Whitespace-only strings are excluded: ``present_values`` represents *present*
# (non-blank) values, and the parser deliberately classifies a blank/whitespace
# string as "missing" (raw_text -> None), so such strings belong to ``_blanks``.
_free_text = st.text(min_size=1, max_size=60).filter(lambda s: s.strip() != "")

# The union used for "present" (non-blank) values. Free text may itself happen
# to be blank, so the test classifies by ``raw_text.strip()`` rather than by
# which generator produced the value.
present_values = st.one_of(
    _plain_numbers,
    _grouped_numbers(),
    _decorated_numbers(),
    _free_text,
)

# Everything, including None and blanks, for the broad round-trip invariant.
any_values = st.one_of(present_values, _blanks, st.none())


# ---------------------------------------------------------------------------
# Properties
# ---------------------------------------------------------------------------


@given(text=any_values, numeric=st.booleans())
def test_capture_preserves_raw_text_verbatim(text, numeric):
    """raw_text is preserved exactly (or marked missing for blank/None).

    The single universal invariant: the parser never alters the printed
    characters. Either the field is missing (blank/None) or ``raw_text`` is
    byte-for-byte identical to the input.

    **Validates: Requirements 2.11, 3.6**
    """

    rv = PdfParser()._capture(text, numeric=numeric)

    if text is None or text.strip() == "":
        # Blank/None -> missing, with no fabricated characters.
        assert rv.is_missing is True
        assert rv.raw_text is None
        assert rv.parsed is None
    else:
        # Present -> verbatim round-trip, never flagged missing.
        assert rv.is_missing is False
        assert rv.raw_text == text
        assert len(rv.raw_text) == len(text)


@given(text=present_values)
def test_capture_text_mode_never_unparseable_and_keeps_value(text):
    """Text mode preserves the value verbatim and is never unparseable.

    Text is always resolvable as text, so ``numeric=False`` never flags a
    present value unparseable; the raw characters survive unchanged.

    **Validates: Requirements 2.11, 3.6**
    """

    rv = PdfParser()._capture(text, numeric=False)

    assert rv.raw_text == text
    assert rv.is_unparseable is False
    # parsed is a non-destructive interpretation: either a number or, when the
    # text is not numeric, the verbatim string itself. Either way raw is intact.
    assert rv.parsed is not None
    if isinstance(rv.parsed, str):
        assert rv.parsed == text


@given(text=present_values)
def test_numeric_parse_is_non_destructive(text):
    """Numeric parsing never mutates raw_text, regardless of parse outcome.

    Whether the value parses as a number (``parsed`` set, not unparseable) or
    cannot be parsed (``is_unparseable`` True), the stored ``raw_text`` always
    equals the source characters exactly.

    **Validates: Requirements 2.11, 3.6**
    """

    rv = PdfParser()._capture(text, numeric=True)

    # Raw text is verbatim in every outcome.
    assert rv.raw_text == text

    if rv.is_unparseable:
        # Located but not numeric -> raw retained, no parsed value, not missing.
        assert rv.parsed is None
        assert rv.is_missing is False
    else:
        # Parsed successfully -> a numeric interpretation sits alongside raw.
        assert rv.parsed is not None
        assert isinstance(rv.parsed, float)


@given(blank=_blanks | st.none())
def test_blank_or_none_yields_missing(blank):
    """Blank or None input yields a missing RawValue with no characters.

    **Validates: Requirements 2.11, 3.6**
    """

    rv = PdfParser()._capture(blank, numeric=False)

    assert rv.is_missing is True
    assert rv.raw_text is None
    assert rv.parsed is None
    assert rv.is_unparseable is False
