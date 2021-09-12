from typing import List
from os import path
import datetime

from model import Model
from view_abc import View, BtnStatus
import excel

class Controller:

    def __init__(self, model: Model, view: View) -> None:
        self.model = model
        self.view = view


    def _are_files_exists(self, input_directory: str) -> List[str]:
        """Returns list of errors. If successful None."""
        errors = []
        for file in self.model.files_to_check:
            full_path = input_directory + '/' + file
            if not path.exists(full_path):
                errors.append(f"File {full_path} does not exists!")
        return errors


    def handle_choose_input_directory_click(self) -> None:
        input_directory = self.view.askForInputDirectory()

        errors = self._are_files_exists(input_directory)
        if errors:
            self.view.alert_error("\n".join(errors))
            return

        self.model.inputDirectory = input_directory
        self.model.path_to_save = self.model.inputDirectory + self.model.outputFilename

        self.view.setInputDirectoryText(self.model.inputDirectory)
        self.view.setSaveLocationText(self.model.path_to_save)

        self.view.checkBtnMergeStatus()
        

    def _checkDates(self, date_from: datetime, date_to: datetime) -> bool:
        if date_from > date_to:
            self.view.alert_error("Dates should be in format YYYY-MM-DD!\nERROR: The Start Date cannot be greater than the End Date!")
            return False
        else:
            return True


    def handle_merge_click(self) -> None:

        date_from = self.view.getStartDate()
        date_to = self.view.getEndDate()

        if not self._checkDates(date_from, date_to):
            return

        self.view.setBtnMergeStatus(BtnStatus.DISABLED)

        excel.copy_all(
            input_directory = self.model.inputDirectory, 
            filenames = self.model.files_to_check,
            path_to_save = self.model.path_to_save,
            date_from = date_from,
            date_to = date_to,
            set_progress_text_fn = self.view.setProgressText
        )

        self.view.notifySound()
        self.view.checkBtnMergeStatus()
    

    def start(self) -> None:
        self.view.setUpView(self)
        self.view.startMainLoop()

