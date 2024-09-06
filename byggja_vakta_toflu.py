"""
This is a script used to read in an Excel spreadsheet from
Vinnustund and format the names based off time and date
"""

from pathlib import Path

from pandas import DataFrame, ExcelWriter, read_excel


class ProgExitError(Exception):
    """
    Custom Exception which is raised whenever the program needs to exit
    """


class CreateShiftsSheet:
    """
    Class for writing the script
    """

    def __init__(self) -> None:
        self.stdout = True
        try:
            _ = print("Program start") if self.stdout is True else None
            self.df_vs_file, check = self.check_workspace()
            if check:
                _ = print("Passed workspace checks") if self.stdout is True else None
                self.df_sheets: dict[str, DataFrame] = {}
                self.weekday_index: dict[int, dict[str, int]] = {}
                self.nicknames: dict[str, str] = self.create_name_nickname_dict()
                self.map_shifts()
                self.seperate_names()
                self.create_shift_excel()

        except ProgExitError:
            # If specified errors occur, the program will write them to the user, and then exit
            pass

    def check_workspace(self) -> tuple[DataFrame, bool]:
        """
        Check workspace to make sure there is only the template.xlsx file and 1 other excel file
        within the current files workspace
        """
        _ = print("Checking workspace") if self.stdout is True else None

        cwd = Path.cwd()
        # TODO: Remove. +3 accounts for .venv .git and .gitignore
        if len([*cwd.iterdir()]) == 4 + 3:
            vs_file = DataFrame()
            template = readme = get_times = extra_excel = False
            for file in cwd.iterdir():
                match file.name:
                    case "template.xlsx":
                        template = True
                    case "README.html":
                        readme = True
                    case "byggja_vakta_toflu.exe":
                        get_times = True
                    case _:
                        if file.suffix == ".xlsx":
                            vs_file = read_excel(file.name, header=None)
                            if vs_file.at[0, 0] == "Starfsmaður":
                                extra_excel = True

            if template and readme and extra_excel and get_times:
                return (vs_file, True)

        self.write_error(
            f"""
            There must only be 4 files in this folder:
            \ntemplate.xlsx, byggja_vakta_toflu.exe, README.html, & the Vinna Excel file\n
            Currently there are:\n
            {"\n".join(map(str, cwd.iterdir()))}
            """
        )
        return (DataFrame(), False)

    def create_name_nickname_dict(self) -> dict[str, str]:
        """
        Create a dictionary name where I truncate the employees full name to
        only use the first name unless there are duplicates, in which case
        I then use 2 or more parts of their name as needed.
        """
        _ = print("Creating name nickname dictionary") if self.stdout is True else None
        nicknames = set()
        ret: dict[str, str] = {}
        for emp_name in self.df_vs_file[1][2:]:
            done = False
            names = emp_name.split()
            index = -1
            while not done:
                index += 1
                try:
                    nickname = " ".join(names[: index + 1])
                    if nickname not in nicknames:
                        nicknames.add(nickname)
                        ret[emp_name] = nickname
                        done = True
                except IndexError as exc:
                    self.write_error(f"The '{emp_name}' name is already in use")
                    raise ProgExitError from exc

        return ret

    def map_name(
        self,
        emp_nickname: str,
        shift_time: str,
        date_day: list[str],
        week_sheet: DataFrame,
    ) -> DataFrame:
        """
        Take in strings employee name, time HH:MM-HH:MM, weekday & sheet denoting the time an
        employee is supposed to be working and write their shift time to the correct workbook sheet.
        """

        def get_time_col(weekday: str):
            for batch, days in self.weekday_index.items():
                if weekday in days:
                    return (week_sheet[batch], batch)

            self.write_error(
                f"""
                Weekday did not match between template sheet and vinna excel sheet\n
                {weekday} is not in template.xlsx
                """
            )
            raise ProgExitError

        def write_unknown_time(row_ind):
            try:
                while not isinstance(week_sheet.at[row_ind, weekday_index], float):
                    row_ind += 1

                week_sheet.at[row_ind, weekday_index] = f"{emp_nickname}: {shift_time}"
                return week_sheet

            except KeyError as exc:
                self.write_error(
                    f"""
                    Program tried to write more than 4 names outside of the normal shift times
                    on {date_day[1]} the {date_day[0]} at time {week_sheet.at[0, weekday_index]}.\n
                    Please write three extra "-" at the bottom of template.xlsx to allow
                    for more unorthodox shift times and see if those three were enough.
                    PS. the more "-" you add the slower the program runs, so only add as many as needed.
                    """
                )
                raise ProgExitError from exc

        (time_column, batch) = get_time_col(date_day[1])
        weekday_index = self.weekday_index[batch][date_day[1]]

        row_ind = -999
        time = ""
        for row_ind, time in enumerate(time_column):
            if isinstance(time, str):
                if time == shift_time:
                    if isinstance(week_sheet.at[row_ind, weekday_index], str):
                        week_sheet.at[row_ind, weekday_index] += f", {emp_nickname}"
                    else:
                        week_sheet.at[row_ind, weekday_index] = emp_nickname
                    return week_sheet

                if time == "Aðrir Tímar":
                    break

        if row_ind != -999 and time == "Aðrir Tímar":
            return write_unknown_time(row_ind)

        self.write_error(
            f"""
            Time {shift_time} on {date_day[1]} the {date_day[0]}.\n
            came from the Vinna Excel sheet but could not be found in template.xlsx\n
            Please add an "Aðrir Tímar" if you want to have unorthodox shift times. 
            """
        )
        raise ProgExitError

    def map_shifts(self):
        """
        Iterate through every shift in the Vinnustund workbook and map the names within it to
        the new employee workbook.
        """
        _ = print("Mapping shifts") if self.stdout is True else None

        def write_date(date_day: list[str], week_sheet):
            for column_index in range(len(week_sheet.columns)):
                if week_sheet.at[1, column_index] == date_day[1]:
                    week_sheet.at[0, column_index] = date_day[0]
                    return

            self.write_error(
                f"""
                Day could not be found to write the date '{date_day[0]}'.\n
                Make sure '{date_day[1]}' is in the second row of template.xlsx
                """
            )
            raise ProgExitError

        def create_time_weekday_index(week_sheet: DataFrame):
            batch = -1
            for column_index in range(len(week_sheet.columns)):
                if isinstance(week_sheet.at[2, column_index], str):
                    batch = column_index
                    self.weekday_index[batch] = {}
                else:
                    self.weekday_index[batch][
                        week_sheet.at[1, column_index]
                    ] = column_index

        week = 1
        self.df_sheets[f"V{week}"] = read_excel("template.xlsx", header=None)
        week_sheet = self.df_sheets[f"V{week}"]
        create_time_weekday_index(week_sheet)

        _ = print(f"V{week}") if self.stdout is True else None
        for column_index in range(2, len(self.df_vs_file.columns)):
            date_day = self.df_vs_file.at[0, column_index].split()
            if date_day[1] == "mán" and date_day[0].split(".")[0] != "11":
                week += 1
                self.df_sheets[f"V{week}"] = week_sheet = read_excel(
                    "template.xlsx", header=None
                )
                _ = print(f"V{week}") if self.stdout is True else None

            write_date(date_day, week_sheet)

            for row_ind, column_cell in enumerate(
                self.df_vs_file[column_index][2:], start=2
            ):
                if isinstance(column_cell, str):
                    self.map_name(
                        self.nicknames[self.df_vs_file.at[row_ind, 1]],
                        column_cell,  # shift_time
                        date_day,
                        week_sheet,
                    )

    def seperate_names(self):
        """
        Run through the time sheets and if a cell contains more than 1 name
        then check if there are enough empty rows beneath the cell to unpack
        its names to the cells below.
        """

        def seperating_names():
            names = column_cell.split(", ")
            num_names = len(names)
            count = 0
            for count in range(1, num_names + 1):
                if not isinstance(week_sheet.at[row_ind + count, column_index], float):
                    break
            if count >= num_names:
                for count, name in enumerate(names):
                    week_sheet.at[row_ind + count, column_index] = name

        _ = print("Seperating names") if self.stdout is True else None
        for sheet_name, week_sheet in self.df_sheets.items():
            _ = print(f"{sheet_name}") if self.stdout is True else None
            for column_index in range(1, len(week_sheet.columns)):
                for row_ind, column_cell in enumerate(
                    week_sheet[column_index][2:], start=2
                ):
                    if (
                        isinstance(column_cell, str)
                        and len(column_cell.split(", ")) > 1
                    ):
                        seperating_names()

    def create_shift_excel(self):
        """
        Create a new excel spreadsheet with the filled in names from Vinnustund
        based on a template.xlsx.
        """
        _ = print("Creating excel spreadsheet") if self.stdout is True else None

        with ExcelWriter("VaktaTafla.xlsx", engine="xlsxwriter") as writer:
            for sheet, df in self.df_sheets.items():
                df.to_excel(writer, sheet_name=sheet, header=False, index=False)
                workbook = writer.book
                worksheet = writer.sheets[sheet]

                bold_format = workbook.add_format({"bold": True})
                worksheet.set_column(0, len(df.columns) - 1, 16)

                worksheet.set_row(0, None, bold_format)
                worksheet.set_row(1, None, bold_format)
                for time_column in self.weekday_index:
                    worksheet.set_column(time_column, time_column, 16, bold_format)

    def write_error(self, msg: str):
        """
        Writes error to file for user to see
        """
        _ = print("Writing error") if self.stdout is True else None

        print(msg)
        _ = input("Press enter on the line to exit program _____")


CreateShiftsSheet()
