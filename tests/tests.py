"""
Unittests for the byggja_vakta_toflu.exe application
"""

import unittest as ut
from os import chdir
from pathlib import Path
from sys import path

path.append(str(Path(__file__).parent.parent.resolve()))
from exceptions import (
    DirContentsError,
    ProgExitError,
    ShiftsOutOfBoundsError,
    TakenEmpNameError,
    UnorthodoxShiftDeniedError,
    WeekdayNotFoundError,
    WriteDateError,
)
from src import byggja_vakta_toflu as app


class TestByggjaVaktaTofluExe(ut.TestCase):
    """
    unittest test class
    """

    def test_happy_path(self):
        """
        Test that everything works when correct parameters are set
        """
        app.CreateShiftsSheet(stdout=False)
        shift_sheet = Path("VaktaTafla.xlsx")
        self.assertTrue(shift_sheet.exists(follow_symlinks=False))
        shift_sheet.unlink()

    def test_wrong_directory(self):
        """
        Test that the program fails when wrong amount of files are in the directory
        """
        extra_file = Path("EXTRA.txt")
        extra_file.touch()
        with self.assertRaises(DirContentsError):
            app.CreateShiftsSheet(test_run=True, stdout=False)

        extra_file.unlink()

    def test_emp_name_in_use(self):
        """
        Test that program fails when too many employees are named the same name
        """
        with self.assertRaises(TakenEmpNameError):
            app.CreateShiftsSheet(
                test_run=True, stdout=False, vinna_excel="../ve_duplicate_emp.xlsx"
            )

    def test_shifts_out_of_bounds(self):
        """
        Test to see if the program responds correctly to an empty template.xlsx file
        """
        with self.assertRaises(ShiftsOutOfBoundsError):
            app.CreateShiftsSheet(
                template="../t_empty.xlsx", test_run=True, stdout=False
            )

    def test_unorthodox_shifts_denied(self):
        with self.assertRaises(UnorthodoxShiftDeniedError):
            app.CreateShiftsSheet(
                test_run=True, stdout=False, template="../t_no_adrir_timar.xlsx"
            )

    def test_weekday_not_found(self):
        """
        Test whether weekdays match between template and vinna excel sheet
        """
        with self.assertRaises(WriteDateError):
            app.CreateShiftsSheet(
                test_run=True, stdout=False, template="../t_wrong_weekday.xlsx"
            )


if __name__ == "__main__":
    chdir(Path("test_env").resolve())
    ut.main(verbosity=2)
