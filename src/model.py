
class Model:
    inputDirectory = None
    files_to_check = [
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
    output_workbook = None
    path_to_save = None
    in_filetypes = (    
                        ("Excel Workbook", "*.xlsx"),
                        ("Excel Macro-Enabled Workbook", "*.xlsm"),
                        ("Excel Macro-Enabled Workbook Template", "*.xltm"),
                        ("Excel Spreadsheet Template", "*.xltx"),
                        # ("OpenDocument Spreadsheet", "*.ods"),
                        # ("CSV (Comma delimited)", "*.csv"),
                        # ("Excel 97-2003 Workbook", "*.xls"),
                        # ("Excel Binary Workbook", "*.xlsb")
                    )

    out_filetypes = (   
                        ("Excel Workbook", "*.xlsx"),
                        # ("OpenDocument Spreadsheet", "*.ods"),
                        # ("CSV (Comma delimited)", "*.csv"),
                    )
