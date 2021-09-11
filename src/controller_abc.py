from abc import ABC, abstractmethod

from model import Model
from view_abc import View

class Controller(ABC):

    @abstractmethod
    def __init__(self, model: Model, view: View) -> None:
        """Initialize controller."""

    
    @abstractmethod
    def handle_choose_input_directory_click(self) -> None:
        """Update model with selected files by the user."""


    @abstractmethod
    def handle_merge_click(self) -> None:
        """Merge files selected by the user and save new file at given location."""


    @abstractmethod
    def start(self) -> None:
        """Start program."""

