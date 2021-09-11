from typing import List

class Model:
    inputDirectory: str = None
    outputFilename: str = "/Directors Basic working file - output.xlsm"
    path_to_save: str = None
    files_to_check: List[str] = [
            "Directors Basic working file.xlsm",
            "I3.xlsx",
            "Recruits (FO).xlsx",
            "Signups.xlsx",
            "Skincare Sets.xlsx",
            "Starter Kits.xlsx",
            "Titles Report Mature Markets.xlsx",
            "Welcome Programme.xlsx",
            "YTD.xlsx"
    ]

