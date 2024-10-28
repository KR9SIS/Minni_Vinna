"""
This is a script used to read in an Excel spreadsheet from
Vinnustund and format the names based off time and date.
The code is owned and written by Kristófer Helgi Sigurðsson
Email: kristoferhelgi@protonmail.com
Github: https://github.com/KR9SIS/RVK_HMH
"""

from argparse import ArgumentParser
from pathlib import Path
from textwrap import dedent
from warnings import catch_warnings, simplefilter

from pandas import DataFrame, ExcelWriter, read_excel


class CreateShiftsSheet:
    """
    Class implementation of the application script
    """

    def __init__(
        self, template="template.xlsx", vinna_excel="", stdout=True, test_run=False
    ) -> None:
        self.stdout = stdout
        self.test_run = test_run
        # Program checks while running if there are gaps in data
        self.missing_dates = []
        try:
            _ = print("Program start") if self.stdout is True else None
            if vinna_excel:
                self.df_v_file = self.get_specific_vs_file(
                    str(Path(vinna_excel).resolve(strict=True))
                )
            else:
                self.df_v_file = self.check_workspace()

            _ = print("Passed workspace checks") if self.stdout is True else None
            self.df_sheets: dict[str, DataFrame] = {}
            self.weekday_index: dict[int, dict[str, int]] = {}
            self.nicknames: dict[str, str] = self.create_name_nickname_dict()
            self.map_shifts(str(Path(template).resolve(strict=True)))
            self.seperate_names()
            self.create_shift_excel()
            self.check_first_last_date()

        except ProgExitError:
            # If specified errors occur, the program will write them to the user, and then exit
            pass

    def check_workspace(self) -> DataFrame:
        """
        Check workspace to make sure it only contains the correct files
        The files allowed are:
        template.xlsx, README.html, byggja_vakta_toflu.exe or .py, VaktaTafla.xlsx and the Vinna excel sheet
        """
        template = readme = get_times = extra_excel = False
        cwd: set[str] = {file.name for file in Path.cwd().iterdir()}
        if "template.xlsx" in cwd:
            cwd.remove("template.xlsx")
            template = True

        if "README.html" in cwd:
            cwd.remove("README.html")
            readme = True

        if "byggja_vakta_toflu.py" in cwd:
            cwd.remove("byggja_vakta_toflu.py")
            get_times = True

        if "byggja_vakta_toflu.exe" in cwd:
            cwd.remove("byggja_vakta_toflu.exe")
            get_times = True

        if "VaktaTafla.xlsx" in cwd:
            cwd.remove("VaktaTafla.xlsx")

        if len(cwd) == 1:
            file = Path(cwd.pop())
            if file.suffix == ".xlsx":
                with catch_warnings():
                    # Ignores following warning
                    # openpyxl\styles\stylesheet.py:237: UserWarning:
                    # Workbook contains no default style, apply openpyxl's default
                    simplefilter("ignore", category=UserWarning)
                    vs_file = read_excel(file.name, header=None)
                if vs_file.at[0, 0] == "Starfsmaður":
                    extra_excel = True

                    if template and readme and get_times and extra_excel:
                        return vs_file

        if self.test_run:
            raise DirContentsError

        contents = "\n".join([path.name for path in Path.cwd().iterdir()])
        self.__write_error(
            dedent(
                """
                There must only be 4 or 5 files in this folder:
                template.xlsx, byggja_vakta_toflu.exe, README.html, VaktaTafla.xlsx
                & the Vinna Excel file where "Starfsmaður" is written in A1.
                Currently there are:
                """
            )
            + f"{contents}"
        )
        raise ProgExitError

    def get_specific_vs_file(self, file: str) -> DataFrame:
        """
        Get's a specific excel file to read from
        """
        return read_excel(file, header=None)

    def create_name_nickname_dict(self) -> dict[str, str]:
        """
        Create a dictionary name where I truncate the employees full name to
        only use the first name unless there are duplicates, in which case
        I then use 2 or more parts of their name as needed.
        """
        _ = print("Creating name nickname dictionary") if self.stdout is True else None
        nicknames = set()
        ret: dict[str, str] = {}
        for emp_name in self.df_v_file[1][2:]:
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
                    elif index > len(self.df_v_file[1][2:]):
                        raise IndexError
                except IndexError as exc:
                    if self.test_run:
                        raise TakenEmpNameError from exc
                    self.__write_error(f"The '{emp_name}' name is already in use")
                    raise ProgExitError from exc

        return ret

    def get_time_col(self, week_sheet: DataFrame, weekday: str):
        """
        Function which takes in a DataFrame and weekday
        and returns the column and index of the weekday in the sheet
        """
        for batch_index, days in self.weekday_index.items():
            if weekday in days:
                return (week_sheet[batch_index], batch_index)
        if self.test_run:
            raise WeekdayNotFoundError
        self.__write_error(
            f"""
            Weekday did not match between template sheet and vinna excel sheet
            {weekday} is not in template.xlsx
            """
        )
        raise ProgExitError

    def map_name(
        self,
        emp_nickname: str,
        shift_time: str,
        date_day: tuple[tuple[int, int, str], str],
        week_sheet: DataFrame,
    ) -> DataFrame:
        """
        Take in strings employee name, time HH:MM-HH:MM, weekday & sheet denoting the time an
        employee is supposed to be working and write their shift time to the correct workbook sheet.
        """

        def write_unknown_time(row_ind):
            try:
                while not isinstance(week_sheet.at[row_ind, weekday_index], float):
                    row_ind += 1

                week_sheet.at[row_ind, weekday_index] = f"{emp_nickname}: {shift_time}"
                return week_sheet

            except KeyError as exc:
                if self.test_run:
                    raise ShiftsOutOfBoundsError from exc
                self.__write_error(
                    f"""
                    Program tried to write more than 4 names outside of the normal shift times
                    on {date_day[1]} the {date_day[0][2]} at time {week_sheet.at[0, weekday_index]}.
                    Please write three extra '-' at the bottom of template.xlsx
                    to allow for more unorthodox shift times and see if those three were enough.
                    PS. the more '-' you add the slower the program runs,
                    so only add as many as needed.
                    """
                )
                raise ProgExitError from exc

        (time_column, batch) = self.get_time_col(week_sheet, date_day[1])
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

        if self.test_run:
            raise UnorthodoxShiftDeniedError
        self.__write_error(
            f"""
            Time {shift_time} on {date_day[1]} the {date_day[0][2]}.
            came from the Vinna Excel sheet but could not be found in template.xlsx
            Please add an 'Aðrir Tímar' if you want to have unorthodox shift times.
            """
        )
        raise ProgExitError

    def map_shifts(self, template):
        """
        Iterate through every shift in the Vinnustund workbook and map the names within it to
        the new employee workbook.
        """
        _ = print("Mapping shifts") if self.stdout is True else None

        def get_date_day(column_index) -> tuple[tuple[int, int, str], str]:
            date_weekday = self.df_v_file.at[0, column_index].split()
            date, month = date_weekday[0].split(".")
            return ((int(date), int(month), date_weekday[0]), date_weekday[1])

        def write_date(date_day: tuple[tuple[int, int, str], str], week_sheet):
            for column_index in range(len(week_sheet.columns)):
                if week_sheet.at[1, column_index] == date_day[1]:
                    week_sheet.at[0, column_index] = date_day[0][2]
                    return

            if self.test_run:
                raise WriteDateError
            self.__write_error(
                f"""
                Weekday could not be found to write the date '{date_day[0][2]}'.
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

        def iter_columns(column_index, date_day, week_sheet):
            for row_ind, column_cell in enumerate(
                self.df_v_file[column_index][2:], start=2
            ):
                if isinstance(column_cell, str):
                    self.map_name(
                        self.nicknames[self.df_v_file.at[row_ind, 1]],
                        column_cell,  # shift_time
                        date_day,
                        week_sheet,
                    )

        week = 1
        self.df_sheets[f"V{week}"] = read_excel(template, header=None)
        week_sheet = self.df_sheets[f"V{week}"]
        create_time_weekday_index(week_sheet)
        # Work through first date seperatively so
        # date checker works out and we can minimize edge case checks
        first_date_v_file = get_date_day(column_index=2)
        write_date(first_date_v_file, week_sheet)
        iter_columns(column_index=2, date_day=first_date_v_file, week_sheet=week_sheet)
        prev_date = first_date_v_file

        _ = print(f"V{week}") if self.stdout is True else None
        for column_index in range(3, len(self.df_v_file.columns)):
            date_day = get_date_day(column_index=column_index)
            if (
                date_day[0][0] != prev_date[0][0] + 1  # Check date against prev date
                and date_day[0][1] == prev_date[0][1]  # Check same month
            ):
                self.missing_dates.append(f"{prev_date[0][0]+1}.{prev_date[0][1]}")
            elif date_day[0][1] != prev_date[0][1] and date_day[0][1] == 1:
                self.missing_dates.append(f"1.{date_day[0][1]}")

            if date_day[1] == "mán" and date_day[0][0] != first_date_v_file[0][0]:
                week += 1
                self.df_sheets[f"V{week}"] = week_sheet = read_excel(
                    template, header=None
                )
                _ = print(f"V{week}") if self.stdout is True else None

            write_date(date_day, week_sheet)

            iter_columns(column_index, date_day, week_sheet)
            prev_date = date_day

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
                if not isinstance(
                    week_sheet.at[row_ind + count, column_index], float
                ) or not isinstance(week_sheet.at[row_ind + count, batch_index], float):
                    break
            if count >= num_names:
                for count, name in enumerate(names):
                    week_sheet.at[row_ind + count, column_index] = name

        _ = print("Seperating names") if self.stdout is True else None
        for sheet_name, week_sheet in self.df_sheets.items():
            _ = print(f"{sheet_name}") if self.stdout is True else None
            for column_index in range(1, len(week_sheet.columns)):
                if week_sheet.at[1, column_index] == "Tímar" or isinstance(
                    week_sheet.at[1, column_index], float
                ):
                    continue
                    # If the current column is a time column, then it contains no names to seperate
                (_, batch_index) = self.get_time_col(
                    week_sheet, week_sheet.at[1, column_index]
                )

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

    def check_first_last_date(self):
        """
        Checks that the last date of the month
        is the date before the first date of the month
        """
        date_day_row = self.df_v_file.iloc[0].to_numpy()
        start_date = int(date_day_row[2][0:2])
        end_date = int(date_day_row[-1][0:2])
        if end_date + 1 != start_date:
            self.missing_dates.append(end_date + 1)
        if self.missing_dates:
            if self.test_run:
                raise VinnaMissingDates(self.missing_dates)

            missing_dates = "\n".join(self.missing_dates)
            self.__write_error(
                dedent(
                    f"""
                        VaktaTafla has been created succesfully, but the program discovered that
                        there were {len(self.missing_dates)} dates missing.
                        The missing dates found were:
                        """
                )
                + f"{missing_dates}"
            )

    def __write_error(self, msg: str):
        """
        Writes error to file for user to see
        """
        _ = print("Writing error") if self.stdout is True else None

        print(dedent(msg), "\n")
        _ = input("Press enter on the line to exit program _____")

    @staticmethod
    def argparsing():
        """
        Argument parsing functionality for command line arguments
        """

        parser = ArgumentParser(description="Parser to check if debug mode is set")
        parser.add_argument(
            "-t",
            "--template",
            required=False,
            default="template.xlsx",
            type=str,
            help="Use specific template document",
        )
        parser.add_argument(
            "-ve",
            "--vinna_excel",
            required=False,
            default="",
            type=str,
            help="Use specific vinna Excel document",
        )
        parser.add_argument(
            "-s",
            "--stdout",
            required=False,
            default=True,
            action="store_false",
            help="Run program without printing messages to stdout",
        )
        parser.add_argument(
            "-test",
            "--test_run",
            required=False,
            default=False,
            action="store_true",
            help="Run program in testing mode",
        )

        return parser.parse_args()


class CustomException(Exception):
    """
    Custom exception for byggja_vakta_toflu.py
    """

    def __init__(self, message: str, error_code: int) -> None:
        self.message = message
        self.error_code = error_code

    def __str__(self) -> str:
        return f"\nMessage:\n{self.message}\nErrorCode: {self.error_code}"


class ProgExitError(CustomException):
    """
    Custom Exception which is raised whenever the program needs to exit.
    """

    def __init__(
        self,
        message: str = "Custom Exception which is raised whenever the program needs to exit.",
        error_code: int = -1,
    ) -> None:
        super().__init__(message, error_code)


class DirContentsError(CustomException):
    """
    Custom Exception which is raised whenever the cwd contents are incorrect during unittest.
    stdout:
        There must only be 4 files in this folder:
        template.xlsx, byggja_vakta_toflu.exe, README.html,
        & the Vinna Excel file where "Starfsmaður" is written in A1.
        Currently there are:
        {contents of the CWD}
    """

    def __init__(self, message: str = "DirContentsError", error_code: int = -2) -> None:
        super().__init__(message, error_code)


class TakenEmpNameError(CustomException):
    """
    Custom Exception which is raised whenever an employee name is already in use during unittest.
    stdout:
        The '{emp_name}' name is already in use
    """

    def __init__(
        self, message: str = "TakenEmpNameError", error_code: int = -3
    ) -> None:
        super().__init__(message, error_code)


class WeekdayNotFoundError(CustomException):
    """
    Custom Exception which is raised whenever a weekday from the Vinna excel sheet
    is not found in template.xlsx during unittest.
    stdout:
        Weekday did not match between template sheet and vinna excel sheet
        {weekday} is not in template.xlsx

    """

    def __init__(
        self, message: str = "WeekdayNotFoundError", error_code: int = -4
    ) -> None:
        super().__init__(message, error_code)


class ShiftsOutOfBoundsError(CustomException):
    """
    Custom Exception which is raised whenever there are
    too many unorthodox shift times during unittest.
    stdout:
        Program tried to write more than 4 names outside of the normal shift times
        on {date_day[1]} the {date_day[0]} at time {week_sheet.at[0, weekday_index]}.
        Please write three extra '-' at the bottom of template.xlsx
        to allow for more unorthodox shift times and see if those three were enough.
        PS. the more '-' you add the slower the program runs,
        so only add as many as needed.
    """

    def __init__(
        self, message: str = "ShiftsOutOfBoundsError", error_code: int = -5
    ) -> None:
        super().__init__(message, error_code)


class UnorthodoxShiftDeniedError(CustomException):
    """
    Custom Exception which is raised whenever "Aðrir Tímar" is not set
    but an unorthodox shift time is found in the Vinna Excel sheet during unittest.
    stdout:
        Time {shift_time} on {date_day[1]} the {date_day[0]}.
        came from the Vinna Excel sheet but could not be found in template.xlsx
        Please add an 'Aðrir Tímar' if you want to have unorthodox shift times.
    """

    def __init__(
        self, message: str = "UnorthodoxShiftDeniedError", error_code: int = -6
    ) -> None:
        super().__init__(message, error_code)


class WriteDateError(CustomException):
    """
    Custom Exception which is raised whenever the program encounters a date in the Vinna Excel sheet
    which is not within the template sheet during unittest.
    stdout:
        Weekday could not be found to write the date '{date_day[0]}'.
        Make sure '{date_day[1]}' is in the second row of template.xlsx
    """

    def __init__(self, message: str = "WriteDateError", error_code: int = -7) -> None:
        super().__init__(message, error_code)


class VinnaMissingDates(CustomException):
    """
    Custom Exception which is raised whenever the program notices
    that the Vinna Excel sheet is missing a date during unittest.
    stdout:
        VaktaTafla has been created succesfully, but the program discovered that
        there were {len(self.missing_dates)} dates missing.
        The missing dates found were:
        {".\n".join(self.missing_dates)}
    """

    def __init__(
        self,
        missing_dates: list[str],
        message: str = "VinnaMissingDates",
        error_code: int = -8,
    ) -> None:
        super().__init__(message, error_code)
        self.missing_dates = missing_dates

    def __str__(self) -> str:
        return f"\nMessage:\n{self.message}\nErrorCode: {self.error_code}\nMissingDates: {self.missing_dates}"


if __name__ == "__main__":
    args = CreateShiftsSheet.argparsing()
    CreateShiftsSheet(
        template=args.template,
        vinna_excel=args.vinna_excel,
        stdout=args.stdout,
        test_run=args.test_run,
    )
