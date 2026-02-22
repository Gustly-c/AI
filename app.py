from pathlib import Path

from scheduler_desktop.repository import StateRepository
from scheduler_desktop.ui import SchedulerDesktopApp


def main() -> None:
    data_path = Path("data/state.xlsx")
    data_path.parent.mkdir(parents=True, exist_ok=True)
    app = SchedulerDesktopApp(StateRepository(data_path))
    app.mainloop()


if __name__ == "__main__":
    main()
