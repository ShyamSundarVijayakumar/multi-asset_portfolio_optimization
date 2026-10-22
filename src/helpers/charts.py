from datetime import date
from pathlib import Path


def chart_dir(data_dir, stage: str, run_date: str = None) -> Path:
    """Returns DATA_DIR/charts/<stage>/<run_date or today>/, creating it if missing."""
    run_date = run_date or date.today().isoformat()
    path = Path(data_dir) / "charts" / stage / run_date
    path.mkdir(parents=True, exist_ok=True)
    return path