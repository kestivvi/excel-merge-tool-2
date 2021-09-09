from model import Model
from view_tkinter import View
from controller import Controller


def main():
    model = Model()
    view = View(model, title="Excel Merge Tool 2")
    controller = Controller(model, view)
    controller.start()


if __name__ == "__main__":
    main()
