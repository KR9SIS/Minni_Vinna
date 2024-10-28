"""
Unittests for the byggja_vakta_toflu.exe application
"""

import unittest as ut
from os import chdir
from pathlib import Path
from subprocess import CalledProcessError, run


class TestByggjaVaktaTofluExe(ut.TestCase):
    """
    unittest test class
    """

    def extract_error_code(self, err_msg: str) -> int:
        """
        Extract ErrorCode from error message and returns it as an integer
        """

        err_code_ind = err_msg.find("ErrorCode: ")
        return int(
            err_msg[
                err_code_ind
                + len("ErrorCode: ") : err_code_ind
                + len("ErrorCode: ")
                + 2
            ]
        )

    def setUp(self):
        self.program = str(Path("../../src/byggja_vakta_toflu.py").resolve())
        self.base_command = ["python3", self.program, "-s", "-test"]
        self.errors = {
            "DirContentsError": -2,
            "TakenEmpNameError": -3,
            "WeekdayNotFoundError": -4,
            "ShiftsOutOfBoundsError": -5,
            "UnorthodoxShiftDeniedError": -6,
            "WriteDateError": -7,
        }
        self.shift_sheet = Path("VaktaTafla.xlsx")

    def test_happy_path(self):
        """
        Test that everything works when correct parameters are set
        """
        try:
            run(["python3", self.program, "-s"], check=True)
            self.assertTrue(self.shift_sheet.exists(follow_symlinks=False))
            self.shift_sheet.unlink()

        except CalledProcessError as exc:
            self.fail(f"Happy Path Raised Error:\n{exc}")

    def test_wrong_directory(self):
        """
        Test that the program fails when wrong amount of files are in the directory
        """
        extra_file = Path("EXTRA.txt")
        extra_file.touch()

        try:
            _ = run(self.base_command, check=True, capture_output=True)
        except CalledProcessError as exc:
            err_code = self.extract_error_code(exc.stderr.decode())
            self.assertEqual(err_code, self.errors["DirContentsError"])

        extra_file.unlink()

    def test_emp_name_in_use(self):
        """
        Test that program fails when too many employees are named the same name
        """
        try:
            cmd = self.base_command
            cmd.extend(["-ve", "../ve_duplicate_emp.xlsx"])
            _ = run(
                cmd,
                check=True,
                capture_output=True,
            )
        except CalledProcessError as exc:
            err_code = self.extract_error_code(exc.stderr.decode())
            self.assertEqual(err_code, self.errors["TakenEmpNameError"])

    def test_shifts_out_of_bounds(self):
        """
        Test to see if the program responds correctly to an empty template.xlsx file
        """
        try:
            cmd = self.base_command
            cmd.extend(["-t", "../t_empty.xlsx"])
            _ = run(
                cmd,
                check=True,
                capture_output=True,
            )
        except CalledProcessError as exc:
            err_code = self.extract_error_code(exc.stderr.decode())
            self.assertEqual(err_code, self.errors["ShiftsOutOfBoundsError"])

    def test_unorthodox_shifts_denied(self):
        """
        Test if program raises right exception when Aðrir Tímar is not located in template.xlsx
        """
        try:
            cmd = self.base_command
            cmd.extend(["-t", "../t_no_adrir_timar.xlsx"])
            _ = run(
                cmd,
                check=True,
                capture_output=True,
            )
        except CalledProcessError as exc:
            err_code = self.extract_error_code(exc.stderr.decode())
            self.assertEqual(err_code, self.errors["UnorthodoxShiftDeniedError"])

    def test_write_date_error(self):
        """
        Test whether weekdays match between template and vinna excel sheet
        """
        try:
            cmd = self.base_command
            cmd.extend(["-t", "../t_wrong_weekday.xlsx"])
            _ = run(
                cmd,
                check=True,
                capture_output=True,
            )
        except CalledProcessError as exc:
            err_code = self.extract_error_code(exc.stderr.decode())
            self.assertEqual(err_code, self.errors["WriteDateError"])


if __name__ == "__main__":
    chdir(Path("test_env").resolve())
    ut.main(verbosity=2)
    _ = input("Press enter to exit")
