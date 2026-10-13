"""Row-level behaviour of the invoice carton parser, on synthetic pages.

The parser is positional, so each test lays out a page as extractor words
(``text``, ``x0``, ``x1``, ``top``) around the layout the real supplier invoices
use: SR NO | DESCRIPTION | TOTAL CTNS | QTY | UNIT | UNIT PRICE | AMOUNT.

Three groups:

- the behaviours PR #2 added (``CTN`` as a unit and as a header spelling, split
  digits, ``THD``, dotted serials, fused ``number+unit`` tokens, split-row
  stitching), pinned here because they shipped without tests;
- the behaviour PR #3 added: a unit that arrives punctuated (``THD.``) or merged
  with its right-hand neighbour (``THD$0.12``) still marks a line item;
- the guards that keep that added tolerance from misreading a page: numbered
  notes are not line items, and the carton header is never a data row.
"""

from __future__ import annotations

import pytest

from boe_converter import invoice_parser
from boe_converter.invoice_parser import InvoicePackingListParser


def W(text, x0, x1, top) -> dict:
    return {"text": text, "x0": float(x0), "x1": float(x1), "top": float(top)}


class _Page:
    def __init__(self, words):
        self._words = words

    def extract_text(self) -> str:
        return "SR NO DESCRIPTION TOTAL CTNS QTY UNIT PRICE TOTAL AMOUNT"

    def extract_words(self):
        return [dict(w) for w in self._words]


class _Doc:
    def __init__(self, *pages):
        self.pages = list(pages)


def header(top=100, label="CTNS", ctns_top=None):
    """Column header; ``ctns_top`` puts the carton label on its own row."""
    ctns_top = top if ctns_top is None else ctns_top
    return [
        W("SR", 30, 40, top), W("NO", 42, 55, top), W("DESCRIPTION", 100, 170, top),
        W("TOTAL", 300, 325, top), W(label, 328, 352, ctns_top),
        W("QTY", 400, 420, top), W("UNIT", 440, 460, top),
        W("PRICE", 480, 510, top), W("AMOUNT", 540, 580, top),
    ]


def row(top, serial, name_words, carton_words, qty, unit, serial_text=None):
    words = [W(serial_text or str(serial), 30, 38, top)]
    x = 100
    for word in name_words:
        words.append(W(word, x, x + 6 * len(word), top))
        x += 6 * len(word) + 4
    words.extend(carton_words)
    words += [
        W(str(qty), 395, 420, top), W(unit, 440, 458, top),
        W("0.5", 480, 495, top), W("2000", 540, 565, top),
    ]
    return words


def ctn(text, top):
    """A carton value centred under the CTNS header."""
    half = 3 * len(text)
    return W(text, 340 - half, 340 + half, top)


def parse(*pages) -> dict[int, tuple]:
    details = InvoicePackingListParser().parse_line_details(
        _Doc(*[_Page(words) for words in pages])
    )
    return {
        serial: (
            None if d["cartons"] is None else d["cartons"].parsed,
            None if d["description"] is None else d["description"].parsed,
        )
        for serial, d in sorted(details.items())
    }


# ---------------------------------------------------------------------------
# Behaviour added by PR #2
# ---------------------------------------------------------------------------
def test_plain_row():
    assert parse(header() + row(120, 1, ["SLIDERS"], [ctn("131", 120)], 4000, "PCS")) == {
        1: (131, "SLIDERS")
    }


def test_blank_carton_cell_stays_blank():
    assert parse(header() + row(120, 37, ["KEYCHAIN"], [], 10, "PCS")) == {37: (None, "KEYCHAIN")}


def test_row_whose_unit_is_ctn_is_kept():
    """The whole row used to be dropped: ``CTN`` was not a recognised unit."""
    page = header() + row(120, 16, ["SILICONE", "SEALANT"], [ctn("1", 120)], 20, "CTN")
    assert parse(page) == {16: (1, "SILICONE SEALANT")}


def test_header_spelled_ctn():
    page = header(label="CTN") + row(120, 1, ["SLIDERS"], [ctn("7", 120)], 10, "PCS")
    assert parse(page) == {1: (7, "SLIDERS")}


def test_digits_split_by_the_extractor_are_joined():
    """``20`` read as ``2`` + ``0`` was booked as 2 cartons."""
    split = [W("2", 334, 339, 120), W("0", 340, 345, 120)]
    assert parse(header() + row(120, 1, ["SLIDERS"], split, 10, "PCS")) == {1: (20, "SLIDERS")}


def test_thd_unit():
    page = header() + row(120, 8, ["METAL", "JOINT", "CLIP"], [ctn("40", 120)], 4000, "THD")
    assert parse(page) == {8: (40, "METAL JOINT CLIP")}


def test_dotted_serial():
    page = header() + row(120, 1, ["SLIDERS"], [ctn("5", 120)], 10, "PCS", serial_text="1.")
    assert parse(page) == {1: (5, "SLIDERS")}


def test_fused_number_and_unit_token():
    page = header() + [
        W("8", 30, 38, 120), W("METAL", 100, 130, 120), W("CLIP", 134, 158, 120),
        ctn("40", 120), W("4000THD", 395, 440, 120),
        W("0.5", 480, 495, 120), W("2000", 540, 565, 120),
    ]
    assert parse(page) == {8: (40, "METAL CLIP")}


def test_fused_token_inside_a_name_stays_verbatim():
    page = header() + row(120, 3, ["126PCS", "STATIONERY", "SET"], [ctn("9", 120)], 10, "SET")
    assert parse(page) == {3: (9, "126PCS STATIONERY SET")}


def test_line_split_across_two_geometric_rows_is_stitched():
    """Name half and numeric half 5pt apart - further than the row tolerance."""
    page = header() + [
        W("8", 30, 38, 120), W("METAL", 100, 130, 120), W("CLIP", 134, 158, 120),
        ctn("40", 125), W("4000", 395, 420, 125), W("THD", 440, 458, 125),
        W("0.5", 480, 495, 125), W("2000", 540, 565, 125),
    ]
    assert parse(page) == {8: (40, "METAL CLIP")}


def test_footer_total_is_never_stitched_onto_a_line():
    page = header() + [
        W("8", 30, 38, 120), W("METAL", 100, 130, 120), W("CLIP", 134, 158, 120),
        W("TOTAL", 100, 130, 126), W("CTN", 300, 318, 126), ctn("1159", 126),
    ]
    assert parse(page) == {}


# ---------------------------------------------------------------------------
# Behaviour added by PR #3: punctuated and merged unit tokens
# ---------------------------------------------------------------------------
def _row_with_unit_token(unit_token, top=120):
    """Serial 8, ``METAL CLIP``, 40 cartons, 4000 of ``unit_token``."""
    return [
        W("8", 30, 38, top), W("METAL", 100, 130, top), W("CLIP", 134, 158, top),
        ctn("40", top), W("4000", 395, 420, top), W(unit_token, 440, 480, top),
        W("480", 540, 565, top),
    ]


@pytest.mark.parametrize("token", ["THD.", "THD,", "CTN:", "thd."])
def test_punctuated_unit_token(token):
    assert parse(header() + _row_with_unit_token(token)) == {8: (40, "METAL CLIP")}


@pytest.mark.parametrize("token", ["THD$0.12", "PCS$1.5", "THD0.12", "KGS/"])
def test_unit_merged_with_its_neighbour(token):
    """A tight layout prints ``4000 THD $0.12`` with no gap the extractor can
    see, so the unit arrives fused to the price. The row is still a line item:
    it must pass the row gate and count as one when the header is located."""
    assert parse(header() + _row_with_unit_token(token)) == {8: (40, "METAL CLIP")}


@pytest.mark.parametrize("token", ["BOXES.", "THDX$1", "12CM", "CM.", "$0.12"])
def test_token_that_only_resembles_a_unit_does_not_qualify_a_row(token):
    assert parse(header() + _row_with_unit_token(token)) == {}


def test_row_with_a_merged_unit_is_complete_and_absorbs_nothing():
    """If a merged unit went unrecognised the row would look like a fragment
    missing its numeric half and stitch the wrapped name line below onto itself."""
    wrapped = [W("GIFT", 100, 124, 126), W("SET", 128, 146, 126)]
    page = header() + _row_with_unit_token("THD$0.12") + wrapped
    assert parse(page) == {8: (40, "METAL CLIP")}


@pytest.mark.parametrize(
    "text,unit",
    [
        ("THD", "THD"), ("thd", "THD"), (" THD. ", "THD"), ("CTN:", "CTN"),
        ("THD$0.12", "THD"), ("PAIR$2", "PAIR"),
    ],
)
def test_unit_token_recognises(text, unit):
    assert invoice_parser._unit_token(text) == unit


@pytest.mark.parametrize(
    "text", ["2PCS", "126PCS", "12CM", "BOXES.", "TOTAL", "DESCRIPTION", "$0.12", ""]
)
def test_unit_token_rejects(text):
    """Digit-led fragments are not this function's business (a fused
    ``4000THD`` is handled separately); longer words are not units."""
    assert invoice_parser._unit_token(text) is None


# ---------------------------------------------------------------------------
# Guards: numbered notes are not line items
# ---------------------------------------------------------------------------
def _two_lines():
    return (
        header()
        + row(120, 1, ["WHITE", "TAPE"], [ctn("131", 120)], 4000, "PCS")
        + row(135, 2, ["KEYCHAIN"], [], 500, "PCS")
    )


def test_numbered_note_does_not_replace_a_line_description():
    """``1.`` reads as serial 1 and ``PCS``/``CTN`` as a unit."""
    note = [
        W("1.", 30, 38, 400), W("PACKING:", 100, 148, 400), W("20", 152, 164, 400),
        W("PCS", 168, 186, 400), W("PER", 190, 208, 400), W("CTN", 212, 230, 400),
    ]
    assert parse(_two_lines() + note) == {1: (131, "WHITE TAPE"), 2: (None, "KEYCHAIN")}


def test_numbered_note_does_not_fill_a_blank_carton_cell():
    note = [
        W("2.", 30, 38, 400), W("NET", 100, 118, 400), W("WEIGHT", 122, 158, 400),
        ctn("1200", 400), W("KG", 370, 382, 400),
    ]
    assert parse(_two_lines() + note) == {1: (131, "WHITE TAPE"), 2: (None, "KEYCHAIN")}


def test_note_on_a_later_page_does_not_replace_an_earlier_line():
    last_page = (
        header()
        + row(120, 3, ["NAIL", "CUTTER"], [ctn("11", 120)], 200, "PCS")
        + [
            W("1.", 30, 38, 400), W("PACKING:", 100, 148, 400), W("20", 152, 164, 400),
            W("PCS", 168, 186, 400),
        ]
    )
    assert parse(_two_lines(), last_page) == {
        1: (131, "WHITE TAPE"), 2: (None, "KEYCHAIN"), 3: (11, "NAIL CUTTER"),
    }


def test_text_above_the_column_header_is_not_a_line():
    """An address line such as ``5 FLOOR, BOX 12`` has a number and a unit word."""
    address = [W("5", 30, 38, 40), W("FLOOR,", 100, 136, 40), W("BOX", 140, 158, 40), W("12", 162, 174, 40)]
    page = address + header() + row(120, 5, ["NAIL", "CUTTER"], [ctn("11", 120)], 200, "PCS")
    assert parse(page) == {5: (11, "NAIL CUTTER")}


# ---------------------------------------------------------------------------
# Guards: locating the carton column
# ---------------------------------------------------------------------------
def test_footer_cannot_displace_a_header_printed_on_its_own_row():
    """The footer shares its row with ``QTY``, which looks like a header hint."""
    page = (
        header(top=100, ctns_top=110)
        + row(130, 1, ["WHITE", "TAPE"], [ctn("131", 130)], 4000, "PCS")
        + row(145, 2, ["KEYCHAIN"], [ctn("15", 145)], 500, "PCS")
        + [
            W("TOTAL", 100, 130, 300), W("CTNS:", 134, 164, 300), W("1159", 168, 192, 300),
            W("TOTAL", 380, 398, 300), W("QTY", 400, 420, 300), W("4500", 424, 448, 300),
        ]
    )
    assert parse(page) == {1: (131, "WHITE TAPE"), 2: (15, "KEYCHAIN")}


def test_data_row_with_unit_ctn_is_never_the_header():
    """``CTN`` is both a header spelling and a unit; ``NO`` is a header hint."""
    page = (
        header(top=100, ctns_top=110)
        + row(130, 1, ["WHITE", "TAPE"], [ctn("131", 130)], 4000, "PCS")
        + row(145, 2, ["MODEL", "NO", "A12"], [ctn("15", 145)], 20, "CTN")
    )
    assert parse(page) == {1: (131, "WHITE TAPE"), 2: (15, "MODEL NO A12")}


def test_title_line_does_not_displace_the_real_header():
    title = [W("PACKED", 100, 136, 40), W("IN", 140, 152, 40), W("500", 156, 174, 40), W("CTNS", 178, 202, 40)]
    page = title + header() + row(120, 1, ["SLIDERS"], [ctn("131", 120)], 4000, "PCS")
    assert parse(page) == {1: (131, "SLIDERS")}


def test_page_without_a_header_yields_nothing():
    """Without a header there is no carton column to read; a ``CTN`` unit in a
    data row must not be mistaken for one (it would misplace every column)."""
    page = row(120, 1, ["SILICONE", "SEALANT"], [ctn("1", 120)], 20, "CTN")
    assert parse(page) == {}


# ---------------------------------------------------------------------------
# Invoice value columns (TOTAL QTY / UNIT PRICE / TOTAL AMOUNT)
# ---------------------------------------------------------------------------
def _details(*pages) -> dict:
    return InvoicePackingListParser().parse_line_details(
        _Doc(*[_Page(words) for words in pages])
    )


def _val(detail, key):
    v = detail.get(key)
    return None if v is None else v.parsed


def test_value_columns_are_captured_per_row():
    page = header() + row(120, 1, ["SLIDERS"], [ctn("131", 120)], 4000, "PCS")
    d = _details(page)[1]
    assert _val(d, "quantity") == 4000
    assert _val(d, "unit_price") == 0.5
    assert _val(d, "amount") == 2000


def test_value_cells_with_currency_and_separators():
    words = (
        header()
        + [W("1", 30, 38, 120), W("KEYCHAIN", 100, 160, 120), W("85", 334, 346, 120)]
        + [W("2913.5", 395, 420, 120), W("DOZ", 440, 458, 120)]
        + [W("$0.12", 480, 505, 120), W("$1,500.00", 540, 580, 120)]
    )
    d = _details(words)[1]
    assert _val(d, "quantity") == 2913.5
    assert _val(d, "unit_price") == 0.12
    assert _val(d, "amount") == 1500.0


def test_blank_value_cells_yield_none():
    words = (
        header()
        + [W("1", 30, 38, 120), W("WIRE", 100, 140, 120)]
        + [W("PCS", 440, 458, 120)]
    )
    d = _details(words)[1]
    assert _val(d, "quantity") is None
    assert _val(d, "unit_price") is None
    assert _val(d, "amount") is None
