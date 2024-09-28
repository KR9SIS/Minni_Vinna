"""
Custom Exceptions used within byggja_vakta_toflu.py
"""


class ProgExitError(Exception):
    """
    Custom Exception which is raised whenever the program needs to exit.
    """


class DirContentsError(Exception):
    """
    Custom Exception which is raised whenever the cwd contents are incorrect during unittest.
    stdout:
        There must only be 4 files in this folder:
        template.xlsx, byggja_vakta_toflu.exe, README.html,
        & the Vinna Excel file where "Starfsmaður" is written in A1.
        Currently there are:
        {contents of the CWD}
    """


class TakenEmpNameError(Exception):
    """
    Custom Exception which is raised whenever an employee name is already in use during unittest.
    stdout:
        The '{emp_name}' name is already in use
    """


class WeekdayNotFoundError(Exception):
    """
    Custom Exception which is raised whenever a weekday from the Vinna excel sheet
    is not found in template.xlsx during unittest.
    stdout:
        Weekday did not match between template sheet and vinna excel sheet
        {weekday} is not in template.xlsx

    """


class ShiftsOutOfBoundsError(Exception):
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


class UnorthodoxShiftDeniedError(Exception):
    """
    Custom Exception which is raised whenever "Aðrir Tímar" is not set
    but an unorthodox shift time is found in the Vinna Excel sheet during unittest.
    stdout:
        Time {shift_time} on {date_day[1]} the {date_day[0]}.
        came from the Vinna Excel sheet but could not be found in template.xlsx
        Please add an 'Aðrir Tímar' if you want to have unorthodox shift times.
    """


class WriteDateError(Exception):
    """
    Custom Exception which is raised whenever the program encounters a date in the Vinna Excel sheet
    which is not within the template sheet during unittest.
    stdout:
        Weekday could not be found to write the date '{date_day[0]}'.
        Make sure '{date_day[1]}' is in the second row of template.xlsx
    """
