import tkinter, tkinter.filedialog, tkinter.messagebox
import threading

from typing import List
from model import Model
from view_abc import View, BtnStatus
from controller_abc import Controller


class View(View):
    # model: Model
    # window: tkinter.Tk
    # input_directory_text: tkinter.StringVar
    # save_location_text = tkinter.StringVar
    # progress_text: tkinter.StringVar
    # btn_merge_and_save: tkinter.Button


    def __init__(self, model: Model, title: str, icon_path: str = "./icon.ico") -> None:
        self.model = model

        self.window = tkinter.Tk()
        self.window.title(title)
        self.window.iconbitmap(icon_path)

        self.input_directory_text = tkinter.StringVar(self.window)
        self.save_location_text = tkinter.StringVar(self.window)
        self.progress_text = tkinter.StringVar(self.window)


    def setUpView(self, controller: Controller) -> None:

        # Two buttons
        btn_choose_input_dir = tkinter.Button(
            self.window, 
            text="Choose input directory", 
            command=controller.handle_choose_input_directory_click, 
            padx=15, 
            pady=5
        )
        btn_choose_input_dir.grid(row=0, column=0, columnspan=2, padx=10, pady=10)

        self.btn_merge_and_save = tkinter.Button(
            self.window, 
            text="Merge and save", 
            command=lambda : 
                threading.Thread(target=controller.handle_merge_click).start(), 
            padx=15, 
            pady=5, 
            state='disabled'
        )
        self.btn_merge_and_save.grid(row=0, column=2, columnspan=2, padx=10, pady=10)


        # Dates
        tkinter.Label(
            self.window, 
            text="Dates should be in format: YYYY-MM-DD"
        ).grid(row=2, column=2, columnspan=2, rowspan=2, padx=15, sticky='W')

        tkinter.Label(
            self.window,
            text="Start date:"
        ).grid(row=2, column=0, pady=5, sticky='E')
        
        self.start_date = tkinter.Entry()
        self.start_date.grid(row=2, column=1, pady=5, sticky='W')

        tkinter.Label(
            self.window,
            text="End date:"
        ).grid(row=3, column=0, pady=5, sticky='E')
        
        self.end_date = tkinter.Entry()
        self.end_date.grid(row=3, column=1, pady=5, sticky='W')


        # Info about input and output
        tkinter.Label(
            self.window,
            text="Input directory:"
        ).grid(row=4, column=0, sticky='E')

        tkinter.Label(
            self.window,
            textvariable=self.input_directory_text
        ).grid(row=4, column=1, columnspan=3, sticky="W", pady=5)

        tkinter.Label(
            self.window, 
            text="Save location:"
        ).grid(row=5, column=0, pady=5, sticky='E')

        tkinter.Label(
            self.window, 
            textvariable=self.save_location_text
        ).grid(row=5, column=1, columnspan=3, sticky="W", pady=5)

        
        # Progress Text
        tkinter.Label(
            self.window, 
            textvariable=self.progress_text, 
            padx=10, 
            pady=5
        ).grid(row=6, column=0, columnspan=3, sticky='W')


    def askForInputDirectory(self, title: str = "Choose input directory") -> str:
        return tkinter.filedialog.askdirectory(title=title)


    def getStartDate(self) -> str:
        return self.start_date.get()


    def getEndDate(self) -> str:
        return self.end_date.get()


    def setBtnMergeStatus(self, status: BtnStatus) -> None:
        self.btn_merge_and_save["state"] = status.value

    # TODO: This should take into the considerations the dates
    def checkBtnMergeStatus(self) -> None:
        if self.model.path_to_save is not None and self.model.inputDirectory is not None:
            self.setBtnMergeStatus(BtnStatus.NORMAL)


    def setProgressText(self, text: str) -> None:
        self.progress_text.set(text)


    def setInputDirectoryText(self, text: str) -> None:
        self.input_directory_text.set(text)


    def setSaveLocationText(self, text: str) -> None:
        self.save_location_text.set(text)


    def notifySound(self) -> None:
        self.window.bell()


    def alert_error(self, msg) -> None:
        tkinter.messagebox.showerror(title="Error", message=msg)


    def startMainLoop(self) -> None:
        self.window.mainloop()

