from tkinter import Tk

from ui import CheckadorApp


def main() -> None:
    root = Tk()
    app = CheckadorApp(root)
    root.mainloop()


if __name__ == "__main__":  # pragma: no cover - entry point
    main()

