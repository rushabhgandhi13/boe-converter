"""Tally master (``Master.json``) loading, matching, and updating.

The Tally *master* is the source of truth for canonical entity names. Before a
Purchase voucher can be imported, every ledger it references (the supplier
"party" ledger, the grouped purchase ledgers, the IGST purchase/payable
ledgers, ``Custom Duty Payable`` ...) must already exist in the company master,
otherwise Tally rejects the import.

This module:

- loads ``Master.json`` (which is **UTF-16 encoded** - always decoded as such);
- indexes every ``Ledger`` / ``Group`` master by name;
- fuzzy-matches a requested name against the existing ledgers and, on a close
  match (default >= 0.9), returns the master's **exact canonical name** so the
  emitted JSON uses the master's spelling/casing (Req: master is source of
  truth for names);
- reports names that have **no** acceptable match so the UI can prompt the user
  to add them;
- appends new minimal ``Ledger`` masters and re-serialises an updated
  ``Master.json`` (UTF-16) the user can import into Tally first.

Nothing about a specific company/party is hard-coded here; callers pass the
names they need matched. Only pure structural constants (the default parent
group for a new supplier ledger) live here.
"""

from __future__ import annotations

import difflib
import io
import json
import uuid
from dataclasses import dataclass

# Encoding of Tally's exported master file. This is NOT optional - the file is
# UTF-16 and decoding it as UTF-8 raises/garbles data.
MASTER_ENCODING = "utf-16"

# Default parent group for a newly-created supplier (party) ledger. Chosen to
# match how existing import suppliers are filed in the sample master
# ("Sundry Creditor Purchase Import").
DEFAULT_SUPPLIER_PARENT = "Sundry Creditor Purchase Import"

# Similarity threshold above which a fuzzy match is accepted as "the same"
# entity (so the master's canonical name is used verbatim).
DEFAULT_CUTOFF = 0.9


@dataclass(frozen=True)
class MatchResult:
    """Outcome of matching a requested name against the master ledgers."""

    query: str
    canonical: str | None          # the master's exact name when matched, else None
    score: float                   # best similarity ratio in [0, 1]
    matched: bool                  # True when score >= cutoff

    @property
    def name(self) -> str:
        """The name to emit: the canonical master name, or the query verbatim."""
        return self.canonical if (self.matched and self.canonical) else self.query


def _norm(s: str) -> str:
    """Normalise a name for comparison (case/space-insensitive)."""
    return " ".join(s.strip().lower().split())


class TallyMaster:
    """An in-memory view over a Tally ``Master.json`` document."""

    def __init__(self, document: dict) -> None:
        self._doc = document
        messages = document.get("tallymessage", [])
        # Preserve every ledger name in first-seen order for stable matching.
        self._ledger_names: list[str] = []
        self._group_names: list[str] = []
        for msg in messages:
            meta = msg.get("metadata") or {}
            mtype = (meta.get("type") or "").strip()
            name = meta.get("name")
            if not name:
                continue
            if mtype == "Ledger":
                self._ledger_names.append(name)
            elif mtype == "Group":
                self._group_names.append(name)
        # name(normalised) -> canonical name (first occurrence wins)
        self._ledger_index: dict[str, str] = {}
        for n in self._ledger_names:
            self._ledger_index.setdefault(_norm(n), n)

    # -- construction -------------------------------------------------------
    @classmethod
    def load(cls, path: str) -> "TallyMaster":
        """Load a master from a UTF-16 ``Master.json`` file on disk."""
        with open(path, "r", encoding=MASTER_ENCODING) as fh:
            return cls(json.load(fh))

    @classmethod
    def load_bytes(cls, raw: bytes) -> "TallyMaster":
        """Load a master from raw UTF-16 (optionally BOM-prefixed) bytes."""
        text = raw.decode(MASTER_ENCODING)
        return cls(json.loads(text))

    # -- access -------------------------------------------------------------
    @property
    def ledger_names(self) -> list[str]:
        """All ledger names in the master (first-seen order)."""
        return list(self._ledger_names)

    def has_ledger(self, name: str) -> bool:
        """True when a ledger with this (normalised) name exists exactly."""
        return _norm(name) in self._ledger_index

    # -- matching -----------------------------------------------------------
    def match(self, name: str, cutoff: float = DEFAULT_CUTOFF) -> MatchResult:
        """Fuzzy-match ``name`` against existing ledgers.

        An exact (normalised) hit scores 1.0. Otherwise the closest ledger by
        :class:`difflib.SequenceMatcher` ratio is considered, and accepted when
        its score is >= ``cutoff``. On acceptance the master's **exact** name is
        returned as ``canonical`` (source-of-truth spelling).
        """
        norm = _norm(name)
        if norm in self._ledger_index:
            return MatchResult(name, self._ledger_index[norm], 1.0, True)

        best_name: str | None = None
        best_score = 0.0
        for canonical in self._ledger_names:
            score = difflib.SequenceMatcher(None, norm, _norm(canonical)).ratio()
            if score > best_score:
                best_score, best_name = score, canonical

        matched = best_score >= cutoff
        return MatchResult(
            query=name,
            canonical=best_name if matched else None,
            score=best_score,
            matched=matched,
        )

    def resolve(self, name: str, cutoff: float = DEFAULT_CUTOFF) -> str:
        """Return the canonical master name for ``name`` (or ``name`` verbatim)."""
        return self.match(name, cutoff).name

    def match_rate_ledger(
        self, kind: str, rate: float, cutoff: float = DEFAULT_CUTOFF
    ) -> MatchResult:
        """Match a rate-scoped ledger (e.g. IGST/purchase ledgers).

        Tally master names for rate ledgers are inconsistently formatted
        ("IGST Purchase @ 5.00 %", "IGST Purchase @ 18%", "IGST PAYABLE @40%").
        Rather than reconstruct that spelling, this narrows to ledgers whose
        name contains every keyword in ``kind`` and whose embedded rate equals
        ``rate``, then returns the master's exact name. Falls back to a plain
        fuzzy :meth:`match` when no rate-bearing candidate is found.

        ``kind`` examples: ``"Factory Purchase Import"``, ``"IGST Purchase"``,
        ``"IGST Payable"``.
        """
        keywords = [k for k in _norm(kind).split() if k]
        # Ledger names embed the *percent* (5, 5.00, 18); ``rate`` is a fraction
        # (0.05). Compare on the percent value.
        target = _rate_key(rate * 100)
        candidates: list[str] = []
        for canonical in self._ledger_names:
            low = _norm(canonical)
            if all(k in low for k in keywords) and _rate_in(canonical) == target:
                candidates.append(canonical)
        if candidates:
            # Prefer the shortest (most specific) candidate name.
            best = min(candidates, key=len)
            return MatchResult(kind, best, 1.0, True)
        # No rate-bearing candidate; fall back to fuzzy match on a synthesised
        # canonical-looking name (percent form).
        synthetic = f"{kind} {_fmt_rate(rate * 100)}%"
        return self.match(synthetic, cutoff)

    def missing(self, names: list[str], cutoff: float = DEFAULT_CUTOFF) -> list[str]:
        """Return the subset of ``names`` with no acceptable master match."""
        out: list[str] = []
        seen: set[str] = set()
        for n in names:
            key = _norm(n)
            if key in seen:
                continue
            seen.add(key)
            if not self.match(n, cutoff).matched:
                out.append(n)
        return out

    # -- mutation -----------------------------------------------------------
    def add_ledger(self, name: str, parent: str = DEFAULT_SUPPLIER_PARENT) -> None:
        """Append a minimal new ``Ledger`` master under ``parent``.

        Enough fields are set for Tally to accept and file the ledger; the user
        can enrich it (GSTIN/address) inside Tally afterwards. A no-op when a
        ledger with this name already exists.
        """
        if self.has_ledger(name):
            return
        guid = f"{uuid.uuid4()}-{uuid.uuid4().hex[:8]}"
        ledger = {
            "metadata": {"type": "Ledger", "name": name, "reservedname": ""},
            "guid": guid,
            "parent": parent,
            "currencyname": "\u20b9",
            "objectupdateaction": "Create",
            "isbillwiseon": True,
            "iscostcentreson": True,
            "isdeleted": False,
            "affectsstock": False,
            "isgstapplicable": False,
            "openingbalance": "0.00",
        }
        self._doc.setdefault("tallymessage", []).append(ledger)
        self._ledger_names.append(name)
        self._ledger_index.setdefault(_norm(name), name)

    # -- serialisation ------------------------------------------------------
    def to_bytes(self) -> bytes:
        """Serialise the (possibly updated) master back to UTF-16 bytes."""
        buf = io.StringIO()
        json.dump(self._doc, buf, ensure_ascii=True, indent=1)
        return buf.getvalue().encode(MASTER_ENCODING)


# ---------------------------------------------------------------------------
# Rate helpers
# ---------------------------------------------------------------------------
def _fmt_rate(rate: float) -> str:
    """Format a GST rate for display: whole numbers without a decimal part."""
    if float(rate).is_integer():
        return str(int(rate))
    return ("%g" % rate)


def _rate_key(rate: float) -> int | None:
    """A comparable integer key for a rate (percent * 100), or None."""
    try:
        return round(float(rate) * 100)
    except (TypeError, ValueError):
        return None


def _rate_in(name: str) -> int | None:
    """Extract the first numeric rate embedded in a ledger name, as a key."""
    import re

    m = re.search(r"(\d+(?:\.\d+)?)", name)
    if not m:
        return None
    return round(float(m.group(1)) * 100)
