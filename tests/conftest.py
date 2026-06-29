"""Shared pytest fixtures and Hypothesis profile configuration.

The default Hypothesis profile runs ``max_examples=100`` per the design's
testing strategy. A faster ``dev`` profile and a thorough ``ci`` profile are
also registered; select one with ``--hypothesis-profile=<name>``.
"""

from __future__ import annotations

from hypothesis import HealthCheck, settings

# Default profile: 100 examples per property (design testing strategy).
settings.register_profile("default", max_examples=100)

# Faster profile for local iteration.
settings.register_profile("dev", max_examples=20)

# Thorough profile for CI runs.
settings.register_profile(
    "ci",
    max_examples=500,
    suppress_health_check=[HealthCheck.too_slow],
)

settings.load_profile("default")
