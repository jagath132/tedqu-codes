"""
==============================================================================
 STUDENT DATA CLEANER  -  student_data_cleaner.py
==============================================================================

 WHAT THIS PROGRAM DOES (in plain English)
 -----------------------------------------
 It reads a messy student Excel file, cleans it, and produces 3 tidy Excel
 files that are ready to upload into your school system:

   1) cleaned_students.xlsx   -> the final, clean, upload-ready student list
   2) duplicates_report.xlsx  -> brothers/sisters who shared a phone or email
   3) changes_report.xlsx      -> every other small correction we made

 You do NOT need to know programming to use this file.
 You only need to change the values inside the box marked:

        USER SETTINGS - UPDATE HERE ONLY

 Everything below that box runs automatically.

 IMPORTANT - COLUMN NAMES
 ------------------------
 The cleaned file now uses the SAME column names as your Excel file, in this
 exact order:

   phone_number | email_id | student_name | academic_year | grade |
   standard | parent_firstname | parent_lastname | date_of_birth | gender(M/F/O)

 Note: in your file the "standard" column holds the SECTION letters (A/B/C),
 so the program cleans it like a section (uppercase + fill blanks from the
 row above).

==============================================================================
 BEGINNER GUIDE  -  READ THIS ONCE
==============================================================================

 1) INSTALL (only the first time):
        Open Command Prompt / Terminal and type:
            pip install pandas openpyxl

 2) SET YOUR FILE PATHS below:
        - INPUT_FILE_PATH must point to the Excel FILE itself (ends in .xlsx)
        - OUTPUT_FOLDER_PATH must point to a FOLDER (where results are saved)

 3) RUN THE PROGRAM:
        Open Command Prompt / Terminal in the folder with this file and type:
            python student_data_cleaner.py

 4) FIND YOUR RESULTS in the folder you set in OUTPUT_FOLDER_PATH.

 WHERE TO CHANGE THINGS LATER (all inside USER SETTINGS):
   - Input Excel file ............ change  INPUT_FILE_PATH
   - Where files are saved ....... change  OUTPUT_FOLDER_PATH
   - Academic year ............... change  ACADEMIC_YEAR
   - Fake phone number series .... change  FAKE_PHONE_START
   - Mailinator email domain ..... change  MAILINATOR_DOMAIN
   - Add new grade rules ......... add a line in  GRADE_MAPPING
   - Add boy / girl name hints ... add names in  MALE_NAMES / FEMALE_NAMES
   - Recognise new Excel headings  edit  COLUMN_NAME_MAPPING
   - Change the output order ..... edit  FINAL_COLUMNS

 HOW TO RENAME A COLUMN (advanced)
 ---------------------------------
 If you ever rename a column, change it in THREE matching places so the
 program stays in sync:
   1) the KEY in COLUMN_NAME_MAPPING
   2) the entry in FINAL_COLUMNS
   3) the matching ROLE value in the "COLUMN ROLES" block
 (The ROLE block simply tells the program which column is the phone, which is
 the grade, and so on.)

==============================================================================
"""

import os
import re
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter


# ============================================================================
# ||                                                                        ||
# ||                    USER SETTINGS - UPDATE HERE ONLY                    ||
# ||                                                                        ||
# ============================================================================

# --- 1) INPUT FILE (must be the .xlsx FILE, not a folder) -------------------
# Windows example: r"C:\Users\Dell\Downloads\Cleaning-format\input\Module_Students_Data.xlsx"
INPUT_FILE_PATH = r"C:\Users\jagat\OneDrive\Desktop\Students Details.xlsx"
# --- 2) OUTPUT FOLDER (a FOLDER; it is created if it does not exist) ---------
OUTPUT_FOLDER_PATH = r"C:\Users\jagat\OneDrive\Desktop"
# --- 3) ACADEMIC YEAR -------------------------------------------------------
# This value is written into the academic_year column for every student.
ACADEMIC_YEAR = "2026"

# --- 4) FAKE PHONE NUMBER SERIES --------------------------------------------
# Used when a phone is missing/invalid OR is a duplicate: 9000000001, 9000000002, ...
FAKE_PHONE_START = 9000000001

# --- 5) MAILINATOR EMAIL --------------------------------------------------
MAILINATOR_DOMAIN = "mailinator.com"

# --- 6) GRADE MAPPING --------------------------------------------------------
# LEFT  = what might appear in your file (lower-case, spaces removed)
# RIGHT = the clean value we want.  Add a line to support a new grade.
#
# NOTE: You do NOT need to list every possible punctuation variant
# (e.g. "class-i", "class i", "class_i", "class.i") separately. The
# standardize_grade() function below automatically treats hyphens (-),
# underscores (_) and dots (.) the same as spaces before it looks the
# value up here, and also tries the value with spaces removed entirely.
# So a single entry like "class i" (or "Class I") already matches
# "CLASS-I", "Class_I", "class.i", "CLASSI", etc.
GRADE_MAPPING = {
    "Pre-KG": "Nursery", "pre-kg": "Nursery", "pre kg": "Nursery",
    "nursery": "Nursery", "nur": "Nursery",
    "lkg": "LKG", "l.k.g": "LKG",
    "ukg": "UKG", "u.k.g": "UKG",
    "1": "I", "1st": "I", "i": "I","Class I" : "I","CLASS I" : "I",
    "2": "II", "2nd": "II", "ii": "II","Class II" : "II","CLASS II" : "II",
    "3": "III", "3rd": "III", "iii": "III","Class III" : "III","CLASS III" : "III",
    "4": "IV", "4th": "IV", "iv": "IV","Class IV" : "IV","CLASS IV" : "IV",
    "5": "V", "5th": "V", "v": "V","Class V" : "V","CLASS V" : "V",
    "6": "VI", "6th": "VI", "vi": "VI","Class VI" : "VI","CLASS VI" : "VI",
    "7": "VII", "7th": "VII", "vii": "VII","Class VII" : "VII","CLASS VII" : "VII",
    "8": "VIII", "8th": "VIII", "viii": "VIII","Class VIII" : "VIII","CLASS VIII" : "VIII",
    "9": "IX", "9th": "IX", "ix": "IX","Class IX" : "IX","CLASS IX" : "IX",
    "10": "X", "10th": "X", "x": "X","Class X" : "X","CLASS X" : "X",
    "11": "XI - Group A1", "11th": "XI - Group A1", "xi": "XI - Group A1","XI - GROUP A1" : "XI - Group A1",
    "Class XI" : "XI - Group A1", "CLASS XI" : "XI - Group A1",
    "12": "XII - Group A1", "12th": "XII - Group A1", "xii": "XII - Group A1","XII - GROUP A1" : "XII - Group A1",
    "Class XII" : "XII - Group A1", "CLASS XII" : "XII - Group A1",

    # GCC / HIFZ / PCC / SCC grades (from the checklist screenshot).
    "gcc 1": "GCC I", "gcc1": "GCC I", "gcc i": "GCC I","GCC 1" : "GCC I",
    "gcc 2": "GCC II", "gcc2": "GCC II", "gcc ii": "GCC II",
    "gcc 3": "GCC III", "gcc3": "GCC III", "gcc iii": "GCC III",
    "gcc 4": "GCC IV", "gcc4": "GCC IV", "gcc iv": "GCC IV",

    "hifz 1st year": "HIFZ - 1ST YEAR", "hifz 1": "HIFZ - 1ST YEAR",
    "hifz year 1": "HIFZ - 1ST YEAR", "hifz i": "HIFZ - 1ST YEAR",
    "hifz first year": "HIFZ - 1ST YEAR",
    "hifz 2nd year": "HIFZ - 2ND YEAR", "hifz 2": "HIFZ - 2ND YEAR",
    "hifz year 2": "HIFZ - 2ND YEAR", "hifz ii": "HIFZ - 2ND YEAR",
    "hifz second year": "HIFZ - 2ND YEAR",
    "hifz 3rd year": "HIFZ - 3RD YEAR", "hifz 3": "HIFZ - 3RD YEAR",
    "hifz year 3": "HIFZ - 3RD YEAR", "hifz iii": "HIFZ - 3RD YEAR",
    "hifz third year": "HIFZ - 3RD YEAR",
    "hifz 4th year": "HIFZ - 4TH YEAR", "hifz 4": "HIFZ - 4TH YEAR",
    "hifz year 4": "HIFZ - 4TH YEAR", "hifz iv": "HIFZ - 4TH YEAR",
    "hifz fourth year": "HIFZ - 4TH YEAR",

    "pcc 1": "PCC I", "pcc1": "PCC I", "pcc i": "PCC I","PCC 1" : "PCC I",
    "pcc 2": "PCC II", "pcc2": "PCC II", "pcc ii": "PCC II","PCC 2" : "PCC II",
    "pcc 3": "PCC III", "pcc3": "PCC III", "pcc iii": "PCC III","PCC 3" : "PCC III",
    "pcc 4": "PCC IV", "pcc4": "PCC IV", "pcc iv": "PCC IV","PCC 4" : "PCC IV",

    "scc": "SCC",
}

# Order grades are sorted in for the final file (low to high).
#
# FIXED: This list previously had "GCC 1", "GCC 2", "GCC 3", "GCC 4" and
# "PCC 1", "PCC 2", "PCC 3", "PCC 4" (digits). But GRADE_MAPPING above always
# converts raw grade text into ROMAN NUMERALS ("GCC I", "GCC II", "PCC I",
# etc). Because "GCC 1" never equals "GCC I", none of the GCC/PCC rows
# matched anything in this list, so they all fell back to the same rank
# (999) and came out in a jumbled, unsorted order. Using the same Roman
# numeral strings that GRADE_MAPPING actually produces fixes the sorting.
GRADE_SORT_ORDER = [
    "Pre-KG","Nursery", "LKG", "UKG",
    "I", "II", "III", "IV", "V", "VI", "VII", "VIII", "IX", "X",
    "XI - Group A1", "XII - Group A1",
    "GCC I", "GCC II", "GCC III", "GCC IV",
    "HIFZ - 1ST YEAR", "HIFZ - 2ND YEAR", "HIFZ - 3RD YEAR", "HIFZ - 4TH YEAR",
    "PCC I", "PCC II", "PCC III", "PCC IV",
    "SCC",
]

# --- 6b) DOB PREDICTION BY GRADE --------------------------------------------
# A date of birth MUST be earlier than today. When a DOB is missing, unreadable,
# or equal to / later than today, we predict a sensible one from the grade:
#       birth_year = int(ACADEMIC_YEAR) - typical_age_for_that_grade
# and set the DOB to "birth_year-01-01". Adjust the typical ages if needed.
GRADE_TYPICAL_AGE = {
    "Nursery": 3, "LKG": 4, "UKG": 5,
    "I": 6, "II": 7, "III": 8, "IV": 9, "V": 10,
    "VI": 11, "VII": 12, "VIII": 13, "IX": 14, "X": 15,
    "XI - Group A1": 16, "XII - Group A1": 17,
    # NOTE: GCC / HIFZ / PCC / SCC are not given a typical age here since
    # there's no standard age for these. If a student in one of these
    # grades has a missing/invalid DOB, it is simply left as-is (no
    # prediction). Add a line like "GCC I": 18 below if you want the
    # program to predict a DOB for these grades too.
}

# --- 6c) PARENT FIRST NAME MINIMUM LENGTH -----------------------------------
# parent_firstname must have at least this many letters. If it is shorter
# (e.g. "MD"), the first word of parent_lastname is pulled in to make it valid.
PARENT_FIRSTNAME_MIN_LEN = 3

# --- 7) GENDER INFERENCE NAMES ----------------------------------------------
# If gender is blank, we guess from the student's FIRST name. Add names (lower-case).
# Gender is always written as just M or F (never O). "Other" is treated as
# unknown and we try to guess from the name.
# If we still cannot decide, we use DEFAULT_GENDER (leave "" for blank, or set
# "M" / "F" to force a value so the column is never empty).
DEFAULT_GENDER = ""
MALE_NAMES = [
    "rahul", "ramesh", "arjun", "arun", "vikram", "rohan", "karthik", "suresh",
    "amit", "ravi", "kiran", "sanjay", "deepak", "anil", "manoj", "raj", "dinesh",
    "raghul", "pranav", "gokul", "sathish", "vignesh", "aravind", "surya",
]
FEMALE_NAMES = [
    "priya", "anita", "kavya", "divya", "sneha", "pooja", "meena",
    "lakshmi", "deepa", "anjali", "sita", "geeta", "geetha", "radha", "nisha",
    "asha", "sowmya", "nithya", "revathi", "anitha", "saranya",
]

# --- 8) COLUMN NAME MAPPING --------------------------------------------------
# LEFT (the KEY) = the column name we use, and also the name in the output file.
# RIGHT (the list) = possible headings in your messy file (case/space/underscore
# do not matter, and anything in (brackets) is ignored when matching).
COLUMN_NAME_MAPPING = {
    "student_name":     ["student name", "name", "student", "studentname"],
    "phone_number":     ["phone number", "phone", "mobile", "contact", "phone no"],
    "email_id":         ["email id", "email", "e-mail", "mail", "emailid"],
    "grade":            ["grade", "class"],
    # In your file "standard" holds the SECTION letters (A/B/C).
    "standard":         ["standard", "section standard", "section", "sec", "division"],
    "parent_firstname": ["parent first name", "parent firstname", "parent name",
                         "father name", "guardian name", "parent"],
    "parent_lastname":  ["parent last name", "parent lastname", "last name",
                         "surname", "lastname"],
    "date_of_birth":    ["date of birth", "dob", "d.o.b", "birth date"],
    "gender(M/F/O)":    ["gender", "sex"],
}

# --- 9) COLUMN ROLES ---------------------------------------------------------
# These tell the program which column does what. Each value must match a KEY in
# COLUMN_NAME_MAPPING above. (academic_year is filled in by the program.)
ROLE_STUDENT_NAME  = "student_name"
ROLE_PHONE         = "phone_number"
ROLE_EMAIL         = "email_id"
ROLE_ACADEMIC_YEAR = "academic_year"
ROLE_GRADE         = "grade"
ROLE_SECTION       = "standard"          # the A/B/C column
ROLE_PARENT_FIRST  = "parent_firstname"
ROLE_PARENT_LAST   = "parent_lastname"
ROLE_DOB           = "date_of_birth"
ROLE_GENDER        = "gender(M/F/O)"

# --- 10) FINAL OUTPUT ORDER --------------------------------------------------
# The exact columns (and order) of cleaned_students.xlsx.
FINAL_COLUMNS = [
    ROLE_PHONE, ROLE_EMAIL, ROLE_STUDENT_NAME, ROLE_ACADEMIC_YEAR,
    ROLE_GRADE, ROLE_SECTION, ROLE_PARENT_FIRST, ROLE_PARENT_LAST,
    ROLE_DOB, ROLE_GENDER,
]

# ----------------------------------------------------------------------------
# ADD FUTURE CONDITIONS HERE
# ----------------------------------------------------------------------------
# New cleaning rule later? Add to a mapping above, or find the matching
# function below (each has a simple comment) and extend it.
# ============================================================================
# ||                       END OF USER SETTINGS                             ||
# ============================================================================


# Columns used to identify a student inside the two report files.
REPORT_ID_COLUMNS = [
    ROLE_STUDENT_NAME, ROLE_GRADE, ROLE_SECTION,
    ROLE_PARENT_FIRST, ROLE_PARENT_LAST, ROLE_DOB, ROLE_GENDER,
]


# ----------------------------------------------------------------------------
# Small helper: turn any cell into a clean, space-trimmed text value.
# Empty / NaN cells become an empty string "".
# ----------------------------------------------------------------------------
def as_text(value):
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except (TypeError, ValueError):
        pass
    return str(value).strip()


# ----------------------------------------------------------------------------
# Hands out fake phone numbers one by one (9000000001, 9000000002, ...) and
# never repeats a number that is already in use by a real student.
# ----------------------------------------------------------------------------
class FakePhoneMaker:
    def __init__(self, start_number, already_used):
        self.next_number = int(start_number)
        self.used = set(already_used)

    def get_next(self):
        while str(self.next_number) in self.used:
            self.next_number += 1
        new_phone = str(self.next_number)
        self.used.add(new_phone)
        self.next_number += 1
        return new_phone


# ----------------------------------------------------------------------------
# Hands out UNIQUE mailinator emails: tries firstname@, then firstnamelastname@,
# then firstname01@, firstname02@, ...
# ----------------------------------------------------------------------------
class EmailMaker:
    def __init__(self, already_used):
        self.used = set(e.lower() for e in already_used)

    def _clean(self, text):
        return re.sub(r"[^a-z0-9]", "", as_text(text).lower())

    def make_unique(self, first_name, last_name=""):
        first = self._clean(first_name) or "student"
        last = self._clean(last_name)

        candidate = f"{first}@{MAILINATOR_DOMAIN}"
        if candidate.lower() not in self.used:
            self.used.add(candidate.lower())
            return candidate

        if last:
            candidate = f"{first}{last}@{MAILINATOR_DOMAIN}"
            if candidate.lower() not in self.used:
                self.used.add(candidate.lower())
                return candidate

        counter = 1
        while True:
            candidate = f"{first}{counter:02d}@{MAILINATOR_DOMAIN}"
            if candidate.lower() not in self.used:
                self.used.add(candidate.lower())
                return candidate
            counter += 1


# ----------------------------------------------------------------------------
# STEP 1: Read the Excel file and tidy up:
#   - friendly errors for wrong path / folder / file-open-in-Excel
#   - remove fully blank rows, strip spaces, drop empty columns
#   - match your headings to our column names (using COLUMN_NAME_MAPPING)
# Returns a list of student dictionaries.
# ----------------------------------------------------------------------------
def read_and_prepare(input_path):
    if not os.path.exists(input_path):
        raise FileNotFoundError(
            f"\nCould not find the input file:\n    {input_path}\n"
            f"Please open the script and fix INPUT_FILE_PATH in USER SETTINGS."
        )

    if os.path.isdir(input_path):
        raise FileNotFoundError(
            f"\nINPUT_FILE_PATH points to a FOLDER, not an Excel file:\n"
            f"    {input_path}\n"
            f"Set it to the full path of the .xlsx file itself, for example:\n"
            f"    {os.path.join(input_path, 'Module_Students_Data.xlsx')}\n"
            f"(Tip: INPUT_FILE_PATH must end in .xlsx . OUTPUT_FOLDER_PATH is "
            f"the one that should be a folder.)"
        )

    try:
        raw = pd.read_excel(input_path, dtype=str)
    except PermissionError:
        raise PermissionError(
            f"\nCould not open the file (Permission denied):\n    {input_path}\n"
            f"This usually means the file is still OPEN in Excel. Please close "
            f"it and run the program again."
        )

    # Turn a heading into a simple comparison key: lower-case, drop "(...)",
    # remove spaces/underscores/punctuation. So "phone_number", "Phone Number"
    # and "gender(M/F/O)" all become easy to match.
    def heading_key(text):
        text = re.sub(r"\(.*?\)", "", as_text(text).lower())
        return re.sub(r"[^a-z0-9]", "", text)

    # Drop columns with no name (e.g. 'Unnamed: 5') or fully empty.
    keep_cols = []
    for col in raw.columns:
        col_name = as_text(col)
        if col_name == "" or col_name.lower().startswith("unnamed"):
            if raw[col].isna().all():
                continue
        keep_cols.append(col)
    raw = raw[keep_cols]

    # comparison-key-of-your-heading -> our column name
    heading_lookup = {}
    for column_name, possible in COLUMN_NAME_MAPPING.items():
        for option in possible:
            heading_lookup[heading_key(option)] = column_name

    found = {}
    for col in raw.columns:
        key = heading_key(col)
        if key in heading_lookup:
            found[heading_lookup[key]] = col

    # Check the truly required columns are present.
    required = [ROLE_PHONE, ROLE_EMAIL, ROLE_STUDENT_NAME, ROLE_GRADE, ROLE_SECTION]
    missing = [c for c in required if c not in found]
    if missing:
        raise ValueError(
            "\nThese required columns are missing from your Excel file:\n"
            f"    {', '.join(missing)}\n"
            "Please add them, or add your heading names to COLUMN_NAME_MAPPING."
        )

    students = []
    for _, row in raw.iterrows():
        record = {}
        for column_name in COLUMN_NAME_MAPPING:
            source_col = found.get(column_name)
            record[column_name] = as_text(row[source_col]) if source_col else ""
        if all(v == "" for v in record.values()):
            continue   # skip fully blank rows
        students.append(record)

    return students


# ----------------------------------------------------------------------------
# Records one correction for the changes report.
# ----------------------------------------------------------------------------
def log_change(changes, student_index, field, original, new, comment):
    # Route the comment into the right column based on which field changed.
    # Email and phone changes get their own dedicated columns; everything else
    # goes into the general Comment column.
    email_remark    = comment if field == "email_id"     else ""
    phone_remark    = comment if field == "phone_number" else ""
    general_comment = comment if field not in ("email_id", "phone_number") else ""

    changes.append({
        "student_index":       student_index,
        "Field Changed":       field,
        "Original Value":      original,
        "New Value":           new,
        "Comment":             general_comment,
        "Email ID Remark":     email_remark,
        "Phone Number Remark": phone_remark,
    })


# ----------------------------------------------------------------------------
# STEP 2: Clean the student name.
#   - remove digits, make UPPERCASE
#   - move leading initials to the end (S RAHUL -> RAHUL S)
# ----------------------------------------------------------------------------
def clean_student_name(name):
    # Only letters are allowed. Remove digits and every other character
    # (keep spaces and dots; dots are used below to detect initials).
    text = re.sub(r"[^A-Za-z\s.]", "", as_text(name))
    text = re.sub(r"\s+", " ", text).strip().upper()
    if text == "":
        return ""
    parts = text.replace(".", " ").split()
    initials = []
    while len(parts) > 1 and len(parts[0]) == 1:
        initials.append(parts.pop(0))
    return " ".join(parts + initials)


# ----------------------------------------------------------------------------
# STEP 3: Clean the parent name.
#   - only letters allowed (digits/symbols removed), make UPPERCASE
#   - move leading initials into parent_lastname (S RAMESH -> RAMESH , last = S)
#   - parent_firstname must be at least PARENT_FIRSTNAME_MIN_LEN letters long.
#     If it is too short (e.g. "MD"), borrow the first word of the last name
#     to build a proper first name ("MD" + "FAIZAL" -> first "MD FAIZAL").
# Returns (parent_firstname, parent_lastname).
# ----------------------------------------------------------------------------
def clean_parent_name(parent_first, last_name):
    # Only letters allowed (keep spaces and dots for the initials logic below).
    text = re.sub(r"[^A-Za-z\s.]", "", as_text(parent_first))
    text = re.sub(r"\s+", " ", text).strip().upper()
    last = re.sub(r"[^A-Za-z\s.]", "", as_text(last_name))
    last = re.sub(r"\s+", " ", last).replace(".", " ").strip().upper()
    if text == "":
        return "", last
    parts = text.replace(".", " ").split()
    moved = []
    while len(parts) > 1 and len(parts[0]) == 1:
        moved.append(parts.pop(0))
    new_first = " ".join(parts)
    if moved and last == "":
        last = " ".join(moved)

    # If the first name is still too short (e.g. "MD"), pull the first word of
    # the last name into the first name so it meets the minimum length.
    if len(new_first.replace(" ", "")) < PARENT_FIRSTNAME_MIN_LEN and last:
        last_parts = last.split()
        borrowed = last_parts.pop(0)
        new_first = f"{new_first} {borrowed}".strip()
        last = " ".join(last_parts)

    return new_first, last


# ----------------------------------------------------------------------------
# STEP 4: Validate a phone number.
#   - split a cell with several numbers, keep the FIRST valid one
#   - valid = exactly 10 digits (also trims a leading 0 or 91 country code)
# Returns (clean_phone_or_blank, had_multiple_numbers).
# ----------------------------------------------------------------------------
def validate_phone(raw_phone):
    text = as_text(raw_phone)
    if text == "":
        return "", False
    parts = re.split(r"[\/,;&\n]| and ", text)
    parts = [p for p in parts if p.strip() != ""]
    had_multiple = len(parts) > 1
    for part in parts:
        digits = re.sub(r"\D", "", part)
        if len(digits) == 12 and digits.startswith("91"):
            digits = digits[2:]
        elif len(digits) == 11 and digits.startswith("0"):
            digits = digits[1:]
        if len(digits) == 10:
            return digits, had_multiple
    squashed = re.sub(r"\D", "", text)
    if len(squashed) == 10:
        return squashed, had_multiple
    return "", had_multiple


# ----------------------------------------------------------------------------
# STEP 5: Does an email look like a real email address?
# ----------------------------------------------------------------------------
def is_valid_email(email):
    text = as_text(email)
    pattern = r"^[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}$"
    return re.match(pattern, text) is not None


# ----------------------------------------------------------------------------
# STEP 6: Standardize the grade using GRADE_MAPPING.
#
#   FIXED: grades that use punctuation as a separator - e.g. "CLASS-I",
#   "Class_I", "class.I", "CLASS - I" - were falling through unmatched and
#   being returned as-is (e.g. "CLASS-I") instead of the Roman numeral
#   format ("I"). We now normalize hyphens/underscores/dots to spaces (and
#   collapse repeated spaces) before comparing against GRADE_MAPPING, in
#   addition to the existing "spaces removed" fallback check.
# ----------------------------------------------------------------------------
_GRADE_MAPPING_LOWER = {k.strip().lower(): v for k, v in GRADE_MAPPING.items()}

def standardize_grade(grade):
    text = as_text(grade).strip().lower()
    # Treat -, _ and . as word separators, same as a space.
    text = re.sub(r"[\-_.]+", " ", text)
    key = re.sub(r"\s+", " ", text).strip()
    if key == "":
        return ""

    # Build a list of candidate keys to look up in GRADE_MAPPING:
    #   1) the key as typed             e.g. "class ukg"
    #   2) the key with a leading       e.g. "ukg"
    #      class/standard/std/grade
    #      word removed
    #   3) each of the above again      e.g. "classukg", "ukg"
    #      with all spaces removed
    # This means a single mapping entry like "ukg" -> "UKG" automatically
    # also matches "CLASS-UKG", "Class_UKG", "STD UKG", "standard.ukg", etc.
    # without needing a separate line added for every prefix variant.
    candidates = [key]
    stripped = re.sub(r"^(class|standard|std|grade)\s+", "", key)
    if stripped and stripped != key:
        candidates.append(stripped)
    for candidate in list(candidates):
        no_space = candidate.replace(" ", "")
        if no_space not in candidates:
            candidates.append(no_space)

    for candidate in candidates:
        if candidate in _GRADE_MAPPING_LOWER:
            return _GRADE_MAPPING_LOWER[candidate]

    return as_text(grade).upper()


# ----------------------------------------------------------------------------
# STEP 7: Standardize a date of birth into YYYY-MM-DD text.
# ----------------------------------------------------------------------------
def standardize_dob(dob):
    text = as_text(dob)
    if text == "":
        return ""
    if re.fullmatch(r"\d{5}", text):   # Excel serial number, e.g. 39448
        try:
            converted = pd.to_datetime(int(text), unit="D", origin="1899-12-30")
            return converted.strftime("%Y-%m-%d")
        except (ValueError, OverflowError):
            pass
    # If it starts with a 4-digit year, read year-first; else day-first.
    year_first = re.match(r"^\d{4}[-/]", text) is not None
    parsed = pd.to_datetime(text, errors="coerce", dayfirst=not year_first)
    if pd.isna(parsed):
        parsed = pd.to_datetime(text, errors="coerce")
    if pd.isna(parsed):
        return text
    return parsed.strftime("%Y-%m-%d")


# ----------------------------------------------------------------------------
# STEP 7b: Is a (already standardized) DOB a real date that is earlier than
# today? Blank / unreadable / today / future all count as NOT valid.
# ----------------------------------------------------------------------------
def dob_is_before_today(dob_text):
    parsed = pd.to_datetime(as_text(dob_text), errors="coerce")
    if pd.isna(parsed):
        return False
    return parsed.normalize() < pd.Timestamp.today().normalize()


# ----------------------------------------------------------------------------
# STEP 7c: Predict a DOB from the (standardized) grade using GRADE_TYPICAL_AGE.
# Returns "YYYY-01-01" text, or "" if the grade is unknown.
# ----------------------------------------------------------------------------
def predict_dob_from_grade(grade):
    age = GRADE_TYPICAL_AGE.get(as_text(grade))
    if age is None:
        return ""
    try:
        birth_year = int(ACADEMIC_YEAR) - age
    except (ValueError, TypeError):
        return ""
    return f"{birth_year}-01-01"


# ----------------------------------------------------------------------------
# STEP 8: Standardize gender to ONLY M or F (never O).
#   - Male->M, Female->F
#   - "Other", blank, or anything else -> guess from the student's first name
#   - if still unknown -> DEFAULT_GENDER (blank by default)
# ----------------------------------------------------------------------------
def standardize_gender(gender, student_name):
    text = as_text(gender).lower()
    if text in ("m", "male"):
        return "M"
    if text in ("f", "female"):
        return "F"
    # "o" / "other" / blank / unknown -> try to guess from the first name.
    first = as_text(student_name).split()[0].lower() if as_text(student_name) else ""
    if first in MALE_NAMES:
        return "M"
    if first in FEMALE_NAMES:
        return "F"
    return DEFAULT_GENDER   # never returns O


# ----------------------------------------------------------------------------
# STEP 9: Clean the section (your "standard" column):
#   - make it UPPERCASE
#   - fill blanks from the row above (merged-cell exports leave blanks)
# ----------------------------------------------------------------------------
def standardize_all_sections(students, changes):
    last_seen = ""
    for i, student in enumerate(students):
        original = student[ROLE_SECTION]
        new = original.upper()
        if new == "":
            new = last_seen
        else:
            last_seen = new
        if new != original:
            student[ROLE_SECTION] = new
            log_change(changes, i, ROLE_SECTION, original, new,
                       "Standard standardized.")


# ----------------------------------------------------------------------------
# Run all single-row cleaning steps and record each correction.
# ----------------------------------------------------------------------------
def clean_each_student(students, changes):
    for i, s in enumerate(students):
        # Student name
        original = s[ROLE_STUDENT_NAME]
        new = clean_student_name(original)
        if new != original:
            log_change(changes, i, ROLE_STUDENT_NAME, original, new,
                       "Student name standardized.")
        s[ROLE_STUDENT_NAME] = new

        # Parent name + last name
        orig_parent = s[ROLE_PARENT_FIRST]
        orig_last = s[ROLE_PARENT_LAST]
        new_parent, new_last = clean_parent_name(orig_parent, orig_last)
        if new_parent != orig_parent or new_last != orig_last:
            log_change(changes, i, ROLE_PARENT_FIRST, orig_parent, new_parent,
                       "Parent name standardized.")
        s[ROLE_PARENT_FIRST] = new_parent
        s[ROLE_PARENT_LAST] = new_last

        # Grade
        orig_grade = s[ROLE_GRADE]
        new_grade = standardize_grade(orig_grade)
        if new_grade != orig_grade:
            log_change(changes, i, ROLE_GRADE, orig_grade, new_grade,
                       "Grade standardized.")
        s[ROLE_GRADE] = new_grade

        # Date of birth: format it, then make sure it is earlier than today.
        # If it is missing / unreadable / not earlier than today, predict a
        # DOB from the grade (this is recorded in the changes report).
        orig_dob = s[ROLE_DOB]
        new_dob = standardize_dob(orig_dob)
        comment = "DOB format standardized."
        if not dob_is_before_today(new_dob):
            predicted = predict_dob_from_grade(s[ROLE_GRADE])
            if predicted:
                new_dob = predicted
                comment = ("DOB missing or not earlier than today; "
                           "predicted from grade.")
        if new_dob != orig_dob:
            log_change(changes, i, ROLE_DOB, orig_dob, new_dob, comment)
        s[ROLE_DOB] = new_dob

        # Gender
        orig_gender = s[ROLE_GENDER]
        new_gender = standardize_gender(orig_gender, s[ROLE_STUDENT_NAME])
        if new_gender != orig_gender:
            log_change(changes, i, ROLE_GENDER, orig_gender, new_gender,
                       "Gender standardized or inferred.")
        s[ROLE_GENDER] = new_gender

    # Section needs the whole list (to look at the row above).
    standardize_all_sections(students, changes)


# ----------------------------------------------------------------------------
# Clean phone numbers for every student (duplicates handled later).
# ----------------------------------------------------------------------------
def clean_phone_numbers(students, changes, phone_maker):
    for i, s in enumerate(students):
        original = s[ROLE_PHONE]
        clean, had_multiple = validate_phone(original)
        if clean != "":
            if had_multiple:
                log_change(changes, i, ROLE_PHONE, original, clean,
                           "Multiple phone numbers found. First valid phone retained.")
            s[ROLE_PHONE] = clean
        else:
            fake = phone_maker.get_next()
            log_change(changes, i, ROLE_PHONE, original, fake,
                       "Invalid phone number replaced with generated phone number.")
            s[ROLE_PHONE] = fake


# ----------------------------------------------------------------------------
# Clean emails for every student (duplicates handled later).
# ----------------------------------------------------------------------------
def clean_emails(students, changes, email_maker):
    for i, s in enumerate(students):
        original = s[ROLE_EMAIL]
        if is_valid_email(original):
            lowered = original.lower()       # emails are always lower-case
            if lowered != original:
                log_change(changes, i, ROLE_EMAIL, original, lowered,
                           "Email converted to lowercase.")
            s[ROLE_EMAIL] = lowered
            email_maker.used.add(lowered)
            continue
        name = as_text(s[ROLE_STUDENT_NAME])
        first = name.split()[0] if name else "student"
        new_email = email_maker.make_unique(first, s[ROLE_PARENT_LAST])
        log_change(changes, i, ROLE_EMAIL, original, new_email,
                   "Invalid email format corrected with unique mailinator email.")
        s[ROLE_EMAIL] = new_email


# ----------------------------------------------------------------------------
# DUPLICATE HANDLING (siblings).
#   - same phone OR same email = siblings; first kept (Original), rest Changed
#   - these corrections go ONLY into the duplicates report, never changes report
# Returns a list of duplicate-report rows.
# ----------------------------------------------------------------------------
def handle_duplicates(students, phone_maker, email_maker):
    phone_groups = {}
    for i, s in enumerate(students):
        phone_groups.setdefault(s[ROLE_PHONE], []).append(i)
    phone_groups = {p: idx for p, idx in phone_groups.items() if len(idx) > 1}

    email_groups = {}
    for i, s in enumerate(students):
        email_groups.setdefault(s[ROLE_EMAIL].lower(), []).append(i)
    email_groups = {e: idx for e, idx in email_groups.items() if len(idx) > 1}

    email_member_sets = {frozenset(idx) for idx in email_groups.values()}

    report_rows = []
    group_counter = [0]
    handled_sets = set()

    def first_name_of(student):
        name = as_text(student[ROLE_STUDENT_NAME])
        return name.split()[0] if name else "student"

    def emit_group(members, dtype, change_phone, change_email):
        group_counter[0] += 1
        group_id = f"G{group_counter[0]:03d}"
        for position, idx in enumerate(members):
            s = students[idx]
            original_phone = s[ROLE_PHONE]
            original_email = s[ROLE_EMAIL]
            new_phone = ""
            new_email = ""
            if position == 0:
                status = "Original"
                comment = "Original sibling record kept unchanged"
            else:
                status = "Changed"
                if change_phone:
                    new_phone = phone_maker.get_next()
                    s[ROLE_PHONE] = new_phone
                if change_email:
                    new_email = email_maker.make_unique(
                        first_name_of(s), s[ROLE_PARENT_LAST])
                    s[ROLE_EMAIL] = new_email
                if change_phone and change_email:
                    comment = "Duplicate phone and email changed"
                elif change_phone:
                    comment = "Duplicate phone changed to fake phone"
                else:
                    comment = "Duplicate email changed to mailinator email"

            row = {
                "Duplicate Group ID": group_id,
                "Duplicate Type": dtype,
            }
            for col in REPORT_ID_COLUMNS:
                row[col] = s[col]
            row.update({
                "Original Phone Number": original_phone,
                "New Phone Number": new_phone,
                "Original Email ID": original_email,
                "New Email ID": new_email,
                "Status": status,
                "Comments": comment,
            })
            report_rows.append(row)

    # 1) Combined groups: same members share BOTH a phone and an email.
    for members in phone_groups.values():
        member_set = frozenset(members)
        if member_set in email_member_sets:
            emit_group(members, "Duplicate Phone + Email", True, True)
            handled_sets.add(member_set)

    # 2) Phone-only groups.
    for members in phone_groups.values():
        member_set = frozenset(members)
        if member_set in handled_sets:
            continue
        emit_group(members, "Duplicate Phone", True, False)
        handled_sets.add(member_set)

    # 3) Email-only groups.
    for members in email_groups.values():
        member_set = frozenset(members)
        if member_set in handled_sets:
            continue
        emit_group(members, "Duplicate Email", False, True)
        handled_sets.add(member_set)

    return report_rows


# ----------------------------------------------------------------------------
# Build the changes-report table (using each student's final details).
# ----------------------------------------------------------------------------
def build_changes_report(students, changes):
    rows = []
    for change in changes:
        s = students[change["student_index"]]
        row = {col: s[col] for col in REPORT_ID_COLUMNS}
        row["Field Changed"] = change["Field Changed"]
        row["Original Value"] = change["Original Value"]
        row["New Value"] = change["New Value"]
        row["Comment"]             = change["Comment"]
        row["Email ID Remark"]     = change.get("Email ID Remark", "")
        row["Phone Number Remark"] = change.get("Phone Number Remark", "")
        rows.append(row)
    columns = REPORT_ID_COLUMNS + ["Field Changed", "Original Value",
                                   "New Value", "Comment",
                                   "Email ID Remark", "Phone Number Remark"]
    return pd.DataFrame(rows, columns=columns)


# ----------------------------------------------------------------------------
# Build the final cleaned table, in the required column order and sort order.
# ----------------------------------------------------------------------------
def build_cleaned_table(students):
    for s in students:
        s[ROLE_ACADEMIC_YEAR] = ACADEMIC_YEAR

    df = pd.DataFrame(students)
    df = df[FINAL_COLUMNS]

    grade_rank = {g: n for n, g in enumerate(GRADE_SORT_ORDER)}
    df["_grade_rank"] = df[ROLE_GRADE].map(lambda g: grade_rank.get(as_text(g), 999))
    df = df.sort_values(by=["_grade_rank", ROLE_SECTION, ROLE_STUDENT_NAME],
                        kind="stable")
    df = df.drop(columns=["_grade_rank"]).reset_index(drop=True)
    return df


# ----------------------------------------------------------------------------
# Build the duplicates-report table, leaving one blank row after each group.
# ----------------------------------------------------------------------------
def build_duplicates_report(duplicate_rows):
    columns = (["Duplicate Group ID", "Duplicate Type"] + REPORT_ID_COLUMNS +
               ["Original Phone Number", "New Phone Number",
                "Original Email ID", "New Email ID", "Status", "Comments"])
    if not duplicate_rows:
        return pd.DataFrame(columns=columns)

    spaced = []
    previous_group = None
    blank = {c: "" for c in columns}
    for row in duplicate_rows:
        if previous_group is not None and row["Duplicate Group ID"] != previous_group:
            spaced.append(dict(blank))
        spaced.append(row)
        previous_group = row["Duplicate Group ID"]
    return pd.DataFrame(spaced, columns=columns)


# ----------------------------------------------------------------------------
# Make any Excel file look professional: bold header, filters, auto-fit,
# centered cells. If highlight_changed, colour 'Changed' rows yellow.
# ----------------------------------------------------------------------------
def format_excel(file_path, highlight_changed=False):
    wb = load_workbook(file_path)
    ws = wb.active
    if ws is None:
        # A workbook always has an active sheet once saved by pandas, but the
        # openpyxl type stubs mark .active as Optional, so we guard for it
        # explicitly (this also satisfies static type checkers like Pylance).
        raise ValueError(f"'{file_path}' has no active worksheet to format.")

    bold = Font(bold=True)
    center = Alignment(horizontal="center", vertical="center")
    yellow = PatternFill(start_color="FFFF00", end_color="FFFF00",
                         fill_type="solid")

    for cell in ws[1]:
        cell.font = bold
        cell.alignment = center

    if ws.max_row >= 1 and ws.max_column >= 1:
        ws.auto_filter.ref = f"A1:{get_column_letter(ws.max_column)}{ws.max_row}"

    for col_index, col_cells in enumerate(ws.columns, start=1):
        longest = 0
        letter = get_column_letter(col_index)
        for cell in col_cells:
            value = "" if cell.value is None else str(cell.value)
            longest = max(longest, len(value))
        ws.column_dimensions[letter].width = min(longest + 3, 45)

    status_col = None
    if highlight_changed:
        for cell in ws[1]:
            if as_text(cell.value).lower() == "status":
                status_col = cell.column
                break

    for row in ws.iter_rows(min_row=2):
        is_changed = False
        if highlight_changed and status_col:
            is_changed = (as_text(row[status_col - 1].value).lower() == "changed")
        for cell in row:
            cell.alignment = center
            if is_changed:
                cell.fill = yellow

    wb.save(file_path)


# ----------------------------------------------------------------------------
# Save a DataFrame to Excel, then apply the professional formatting.
# ----------------------------------------------------------------------------
def save_excel(df, file_path, highlight_changed=False):
    df.to_excel(file_path, index=False)
    format_excel(file_path, highlight_changed=highlight_changed)


# ----------------------------------------------------------------------------
# THE MAIN PROGRAM - runs when you type: python student_data_cleaner.py
# ----------------------------------------------------------------------------
def main():
    print("Student Data Cleaner - starting...")

    if not os.path.exists(OUTPUT_FOLDER_PATH):
        os.makedirs(OUTPUT_FOLDER_PATH)
        print(f"Created output folder: {OUTPUT_FOLDER_PATH}")

    students = read_and_prepare(INPUT_FILE_PATH)
    print(f"Read {len(students)} student rows.")

    changes = []
    clean_each_student(students, changes)

    existing_phones = [s[ROLE_PHONE] for s in students]
    phone_maker = FakePhoneMaker(FAKE_PHONE_START, existing_phones)
    email_maker = EmailMaker([])

    clean_phone_numbers(students, changes, phone_maker)
    clean_emails(students, changes, email_maker)

    duplicate_rows = handle_duplicates(students, phone_maker, email_maker)

    cleaned_df = build_cleaned_table(students)
    changes_df = build_changes_report(students, changes)
    duplicates_df = build_duplicates_report(duplicate_rows)

    cleaned_path = os.path.join(OUTPUT_FOLDER_PATH, "cleaned_students.xlsx")
    changes_path = os.path.join(OUTPUT_FOLDER_PATH, "changes_report.xlsx")
    duplicates_path = os.path.join(OUTPUT_FOLDER_PATH, "duplicates_report.xlsx")

    save_excel(cleaned_df, cleaned_path)
    save_excel(changes_df, changes_path)
    save_excel(duplicates_df, duplicates_path, highlight_changed=True)

    print("\nDone! Files saved in:", os.path.abspath(OUTPUT_FOLDER_PATH))
    print(f"  1) cleaned_students.xlsx  ({len(cleaned_df)} students)")
    print(f"  2) changes_report.xlsx     ({len(changes_df)} corrections)")
    print(f"  3) duplicates_report.xlsx  ({len(duplicate_rows)} sibling rows)")


if __name__ == "__main__":
    try:
        main()
    except FileNotFoundError as error:
        print("INPUT FILE PROBLEM:", error)
    except ValueError as error:
        print("COLUMN PROBLEM:", error)
    except PermissionError as error:
        print("PERMISSION PROBLEM:", error)
        print("If an OUTPUT file is open in Excel (cleaned_students.xlsx etc.), "
              "close it and run again.")
    except Exception as error:
        print("Something went wrong:", error)

