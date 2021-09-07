from typing import List
import tkinter, tkinter.filedialog, tkinter.messagebox, threading

from view_abc import View, BtnStatus
from model import Model
from controller_abc import Controller

class View(View):
    # model: Model
    # window: tkinter.Tk
    # btn_merge: tkinter.Button
    # progress_text: tkinter.StringVar
    # save_location_text = tkinter.StringVar


    def __init__(self, model: Model, title: str = "Excel Merge Tool 2", icon_path: str = "./icon.ico"):
        self.model = model
        self.window = tkinter.Tk()

        self.window.title(title)
        self.window.iconbitmap(icon_path)


    def setUpView(self, controller: Controller):
        btn_load = tkinter.Button(self.window, text="Choose input directory", command=controller.handle_choose_input_directory_click, padx=15, pady=5)
        btn_load.grid(row=0, column=0, columnspan=2, padx=10, pady=10)

        self.btn_merge = tkinter.Button(self.window, text="Merge and save", command=lambda : threading.Thread(target=controller.handle_merge_click).start(), padx=15, pady=5, state='disabled')
        self.btn_merge.grid(row=0, column=2, columnspan=2, padx=10, pady=10)


        label5 = tkinter.Label(self.window, text="Start date:")
        label5.grid(row=1, column=0, pady=5)
        self.start_date = tkinter.Entry()
        self.start_date.grid(row=1, column=1, pady=5)

        label6 = tkinter.Label(self.window, text="End date:")
        label6.grid(row=2, column=0, pady=5)
        self.end_date = tkinter.Entry()
        self.end_date.grid(row=2, column=1, pady=5)


        label1 = tkinter.Label(self.window, text="Input directory:")
        label1.grid(row=3, column=0)

        self.input_directory_text = tkinter.StringVar(self.window)
        label4 = tkinter.Label(self.window, textvariable=self.input_directory_text)
        label4.grid(row=3, column=1, columnspan=3, sticky="W", pady=5)
        self.input_directory_text.set("None")

        label2 = tkinter.Label(self.window, text="Save location : ")
        label2.grid(row=4, column=0, pady=5)

        self.save_location_text = tkinter.StringVar(self.window)
        label3 = tkinter.Label(self.window, textvariable=self.save_location_text)
        label3.grid(row=4, column=1, columnspan=3, sticky="W", pady=5)
        self.save_location_text.set("None")

        self.progress_text = tkinter.StringVar(self.window)
        progress_label = tkinter.Label(self.window, textvariable=self.progress_text, padx=10, pady=5)
        progress_label.grid(row=5, column=0, columnspan=3, sticky='W')


    def askForInputDirectory(self, title: str = "Choose input directory") -> List[str]:
        return tkinter.filedialog.askdirectory(title=title)


    def askForOutputFiles(self, title: str = "Save as") -> List[str]:
        return tkinter.filedialog.asksaveasfilename(title=title, filetypes=self.model.out_filetypes, defaultextension='.xlsx')
    

    def getStartDate(self):
        return self.start_date.get()


    def getEndDate(self):
        return self.end_date.get()


    def setBtnMergeStatus(self, status: BtnStatus):
        self.btn_merge["state"] = status.value


    def checkBtnMergeStatus(self):
        if self.model.path_to_save is not None and self.model.inputDirectory is not None:
            self.setBtnMergeStatus(BtnStatus.NORMAL)


    def updateFileList(self):
        self.input_directory_text.delete(0, tkinter.END)
        [self.input_directory_text.insert(tkinter.END, file) for file in self.model.inputDirectory]


    def setProgressText(self, text: str):
        self.progress_text.set(text)


    def setInputDirectoryText(self, text: str):
        self.input_directory_text.set(text)


    def setSaveLocationText(self, text: str):
        self.save_location_text.set(text)


    def notifySound(self):
        self.window.bell()


    def alert_error(self, msg):
        tkinter.messagebox.showerror(title="Error", message=msg)


    def startMainLoop(self):
        self.window.mainloop()
