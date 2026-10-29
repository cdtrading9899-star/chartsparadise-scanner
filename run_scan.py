"""Headless entry point for GitHub Actions: no inline display, no Colab download."""
import os
import matplotlib

matplotlib.use("Agg")

import chartsparadise_scanner as scanner  # noqa: E402

overrides = {"show_inline": False, "download_zip": False}
if os.path.exists("universe.csv"):          # optional: commit a Symbol CSV if NSE blocks the runner's IP
    overrides.update(universe="csv", universe_csv="universe.csv")

scanner.main(overrides)
