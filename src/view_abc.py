from abc import ABC, abstractmethod
from datetime import datetime
from enum import Enum
from typing import List


class BtnStatus(Enum):
    DISABLED = "disabled"
    NORMAL = "normal"


class View(ABC):

    @abstractmethod
    def __init__(self, title: str, icon_path: str) -> None:
        """Initialize window, set up title and icon."""


    @abstractmethod
    def setUpView(self, controller) -> None:
        """Set up window, placement of labels and buttons."""


    @abstractmethod
    def startMainLoop(self) -> None:
        """Start the main loop of the program."""


    @abstractmethod
    def askForInputDirectory(self, title: str = "Choose input directory") -> str:
        """Ask user to select input files."""


    @abstractmethod
    def getStartDate(self) -> datetime:
        """Get the Start Date choosen by the user.
        
        Returns
        -------
        datetime
            If there were no erros.
        None
            If there were errors.
        """


    @abstractmethod
    def getEndDate(self) -> datetime:
        """Get the End Date choosen by the user.
        
        Returns
        -------
        datetime
            If there were no erros.
        None
            If there were errors.
        """


    @abstractmethod
    def checkBtnMergeStatus(self) -> None:
        """Check and set the status of the merge button."""
    

    @abstractmethod
    def setBtnMergeStatus(self, status: BtnStatus) -> None:
        """Set the status of the merge button."""


    @abstractmethod
    def setProgressText(self, text: str) -> None:
        """Set the text displayed as the progress."""


    @abstractmethod
    def setSaveLocationText(self, text: str) -> None:
        """Set the text with the choosen location of output file."""
    

    @abstractmethod
    def setInputDirectoryText(self, text: str) -> None:
        """Set the text with the choosen input directory."""
    

    @abstractmethod
    def notifySound(self) -> None:
        """Play the notification sound."""
    

    @abstractmethod
    def alert_error(self, msg: str) -> None:
        """Pop out error alert with a message."""

