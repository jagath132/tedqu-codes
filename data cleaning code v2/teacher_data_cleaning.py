"""
==============================================================================
 TEACHER DATA CLEANER  -  teacher_data_cleaner.py
==============================================================================

 WHAT THIS PROGRAM DOES (in plain English)
 -----------------------------------------
 It reads a messy teacher Excel file, cleans it, and produces 3 tidy Excel
 files that are ready to upload into your school system:

   1) cleaned_teachers.xlsx    -> the final, clean, upload-ready teacher list
   2) duplicates_report.xlsx   -> teachers who shared a phone or email
   3) changes_report.xlsx      -> every other small correction we made

 You do NOT need to know programming to use this file.
 You only need to change the values inside the box marked:

        USER SETTINGS - UPDATE HERE ONLY

 Everything below that box runs automatically.

 IMPORTANT - COLUMN NAMES
 ------------------------
 The cleaned file uses the SAME column names as your teacher sample, in this
 exact order:

   firstname | lastname | email_id | phone_number | date_of_birth |
   gender(M/F/O) | staffcode

 NAME RULE (as requested):
   The teacher's NAME becomes the FIRST NAME, and any INITIAL becomes the
   LAST NAME. Examples:
     "S RAHUL"          -> firstname = RAHUL,   lastname = S
     "RAHUL S"          -> firstname = RAHUL,   lastname = S
     "STEPHEN STRANGE"  -> firstname = STEPHEN, lastname = STRANGE
   If your file already has separate firstname / lastname columns, they are
   respected and only tidied (case + trimming).

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
            python teacher_data_cleaner.py

 4) FIND YOUR RESULTS in the folder you set in OUTPUT_FOLDER_PATH.

 WHERE TO CHANGE THINGS LATER (all inside USER SETTINGS):
   - Input Excel file ............ change  INPUT_FILE_PATH
   - Where files are saved ....... change  OUTPUT_FOLDER_PATH
   - Fake phone number series .... change  FAKE_PHONE_START
   - Mailinator email domain ..... change  MAILINATOR_DOMAIN
   - Name capitalisation ......... change  NAME_CASE
   - Duplicate-email fallback .... change  FALLBACK_EMAIL_STYLE
   - Add boy / girl name hints ... add names in  MALE_NAMES / FEMALE_NAMES
   - Recognise new Excel headings  edit  COLUMN_NAME_MAPPING
   - Change the output order ..... edit  FINAL_COLUMNS

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
# Windows example: r"C:\Users\Dell\Downloads\Cleaning-format\input\Fake_Teachers.xlsx"
INPUT_FILE_PATH = r"D:\daps  school\APS SCHOOL\Teacher Details.xlsx"

# --- 2) OUTPUT FOLDER (a FOLDER; it is created if it does not exist) ---------
OUTPUT_FOLDER_PATH = r"D:\daps  school\APS SCHOOL\teacher details"
# --- 3) FAKE PHONE NUMBER SERIES --------------------------------------------
# Used when a phone is missing/invalid OR is a duplicate: 9000000001, 9000000002, ...
FAKE_PHONE_START = 9000000001

# --- 4) MAILINATOR EMAIL -----------------------------------------------------
MAILINATOR_DOMAIN = "mailinator.com"

# --- 5) NAME CAPITALISATION --------------------------------------------------
# How teacher first/last names appear in the cleaned file:
#   "title" -> Stephen Strange   (matches your Fake_Teachers.xlsx sample)
#   "upper" -> STEPHEN STRANGE
#   "keep"  -> leave as typed (only trims extra spaces)
NAME_CASE = "title"

# --- 6) DUPLICATE-EMAIL FALLBACK STYLE ---------------------------------------
# When "firstname@mailinator.com" is already taken and a last name exists,
# the next email is built as:
#   "concat" -> firstnamelastname@mailinator.com   (e.g. stephenstrange@...)
#   "dot"    -> firstname.lastname@mailinator.com   (e.g. stephen.strange@...)
FALLBACK_EMAIL_STYLE = "concat"

# --- 7) GENDER INFERENCE NAMES ----------------------------------------------
# If gender is blank, we guess from the teacher's FIRST name. Add names (lower-case).
# Male -> M, Female -> F, Other -> O. If gender is blank AND the name is unknown,
# we use DEFAULT_GENDER (leave "" for blank, or set "M"/"F"/"O" to force a value).
DEFAULT_GENDER = ""
MALE_NAMES = [
    "rahul", "ramesh", "arjun", "arun", "vikram", "rohan", "karthik", "suresh",
    "amit", "ravi", "kiran", "sanjay", "deepak", "anil", "manoj", "raj", "dinesh",
    "raghul", "pranav", "gokul", "sathish", "vignesh", "aravind", "surya",
    "stephen", "bruce", "loki", "sam", "challa", "peter", "clint",
]
FEMALE_NAMES = [
    "priya", "anita", "kavya", "divya", "sneha", "pooja", "meena",
    "lakshmi", "deepa", "anjali", "sita", "geeta", "geetha", "radha", "nisha",
    "asha", "sowmya", "nithya", "revathi", "anitha", "saranya",
    "carol", "shuri", "natasha",
]

# --- 8) COLUMN NAME MAPPING --------------------------------------------------
# LEFT (the KEY) = the column name we use, and also the name in the output file.
# RIGHT (the list) = possible headings in your messy file (case/space/underscore
# do not matter, and anything in (brackets) is ignored when matching).
#
# NOTE: "teacher name" / "name" is treated as a COMBINED name and is split into
# firstname + lastname (initial). If your file has separate firstname / lastname
# columns, those are used directly instead.
COLUMN_NAME_MAPPING = {
    "firstname":     ["firstname", "first name", "teacher first name",
                      "teacher name", "name", "teacher", "staff name",
                      "teachername", "staffname"],
    "lastname":      ["lastname", "last name", "surname", "initial",
                      "teacher last name"],
    "email_id":      ["email id", "email", "e-mail", "mail", "emailid"],
    "phone_number":  ["phone number", "phone", "mobile", "contact", "phone no"],
    "date_of_birth": ["date of birth", "dob", "d.o.b", "birth date"],
    "gender(M/F/O)": ["gender", "sex"],
    "staffcode":     ["staffcode", "staff code", "staff id", "employee code",
                      "emp code", "teacher code", "code"],
}

# --- 9) COLUMN ROLES ---------------------------------------------------------
# These tell the program which column does what. Each value must match a KEY in
# COLUMN_NAME_MAPPING above.
ROLE_FIRST     = "firstname"
ROLE_LAST      = "lastname"
ROLE_EMAIL     = "email_id"
ROLE_PHONE     = "phone_number"
ROLE_DOB       = "date_of_birth"
ROLE_GENDER    = "gender(M/F/O)"
ROLE_STAFFCODE = "staffcode"

# --- 10) FINAL OUTPUT ORDER --------------------------------------------------
# The exact columns (and order) of cleaned_teachers.xlsx.
FINAL_COLUMNS = [
    ROLE_FIRST, ROLE_LAST, ROLE_EMAIL, ROLE_PHONE,
    ROLE_DOB, ROLE_GENDER, ROLE_STAFFCODE,
]

# ----------------------------------------------------------------------------
# ADD FUTURE CONDITIONS HERE
# ----------------------------------------------------------------------------
# New cleaning rule later? Add to a mapping above, or find the matching
# function below (each has a simple comment) and extend it.
# ============================================================================
# ||                       END OF USER SETTINGS                             ||
# ============================================================================


# Columns used to identify a teacher inside the two report files.
REPORT_ID_COLUMNS = [
    ROLE_FIRST, ROLE_LAST, ROLE_STAFFCODE, ROLE_DOB, ROLE_GENDER,
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
# Apply the chosen NAME_CASE ("title" / "upper" / "keep") to a name.
# ----------------------------------------------------------------------------
def apply_name_case(text):
    text = re.sub(r"\s+", " ", as_text(text)).strip()
    if text == "":
        return ""
    if NAME_CASE == "upper":
        return text.upper()
    if NAME_CASE == "title":
        return text.title()
    return text   # "keep"


# ----------------------------------------------------------------------------
# Hands out fake phone numbers one by one (9000000001, 9000000002, ...) and
# never repeats a number that is already in use by a real teacher.
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
# Hands out UNIQUE mailinator emails:
#   tries firstname@, then firstnamelastname@ (or firstname.lastname@),
#   then firstname01@, firstname02@, ...
# ----------------------------------------------------------------------------
class EmailMaker:
    def __init__(self, already_used):
        self.used = set(e.lower() for e in already_used)

    def _clean(self, text):
        return re.sub(r"[^a-z0-9]", "", as_text(text).lower())

    def make_unique(self, first_name, last_name=""):
        first = self._clean(first_name) or "teacher"
        last = self._clean(last_name)

        candidate = f"{first}@{MAILINATOR_DOMAIN}"
        if candidate.lower() not in self.used:
            self.used.add(candidate.lower())
            return candidate

        if last:
            joiner = "." if FALLBACK_EMAIL_STYLE == "dot" else ""
            candidate = f"{first}{joiner}{last}@{MAILINATOR_DOMAIN}"
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
# Returns a list of teacher dictionaries.
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
            f"    {os.path.join(input_path, 'Fake_Teachers.xlsx')}\n"
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

    # comparison-key-of-your-heading -> our column name.
    # (First match wins, so more specific headings are listed first above.)
    heading_lookup = {}
    for column_name, possible in COLUMN_NAME_MAPPING.items():
        for option in possible:
            heading_lookup.setdefault(heading_key(option), column_name)

    found = {}
    for col in raw.columns:
        key = heading_key(col)
        if key in heading_lookup:
            target = heading_lookup[key]
            found.setdefault(target, col)   # keep the first column found per role

    # We need at least a name for every teacher.
    if ROLE_FIRST not in found:
        raise ValueError(
            "\nCould not find a teacher name column in your Excel file.\n"
            "Add a 'firstname' (or 'teacher name'/'name') column, or add your "
            "heading to COLUMN_NAME_MAPPING."
        )

    teachers = []
    for _, row in raw.iterrows():
        record = {}
        for column_name in COLUMN_NAME_MAPPING:
            source_col = found.get(column_name)
            record[column_name] = as_text(row[source_col]) if source_col else ""
        if all(v == "" for v in record.values()):
            continue   # skip fully blank rows
        teachers.append(record)

    return teachers


# ----------------------------------------------------------------------------
# Records one correction for the changes report.
# ----------------------------------------------------------------------------
def log_change(changes, teacher_index, field, original, new, comment):
    # Email and phone changes get their own dedicated columns; everything else
    # goes into the general Comment column.
    email_remark    = comment if field == ROLE_EMAIL else ""
    phone_remark    = comment if field == ROLE_PHONE else ""
    general_comment = comment if field not in (ROLE_EMAIL, ROLE_PHONE) else ""

    changes.append({
        "teacher_index":       teacher_index,
        "Field Changed":       field,
        "Original Value":      original,
        "New Value":           new,
        "Comment":             general_comment,
        "Email ID Remark":     email_remark,
        "Phone Number Remark": phone_remark,
    })


# ----------------------------------------------------------------------------
# STEP 2: Split / clean the teacher name.
#   RULE: the name becomes the FIRST NAME and any INITIAL becomes the LAST NAME.
#   - "S RAHUL"         -> first = RAHUL,   last = S
#   - "RAHUL S"         -> first = RAHUL,   last = S
#   - "STEPHEN STRANGE" -> first = STEPHEN, last = STRANGE
#   If a last name is already provided separately, we keep it and only tidy it.
# Returns (firstname, lastname).
# ----------------------------------------------------------------------------
def clean_teacher_name(first_raw, last_raw):
    text = re.sub(r"\d", "", as_text(first_raw))       # remove any digits
    text = re.sub(r"\s+", " ", text.replace(".", " ")).strip()
    last = as_text(last_raw)

    # If a last name already exists, just tidy both and return.
    if last != "":
        return apply_name_case(text), apply_name_case(last)

    if text == "":
        return "", ""

    parts = text.split()
    if len(parts) == 1:
        return apply_name_case(parts[0]), ""

    # Leading single-letter initial:  "S RAHUL" -> RAHUL / S
    if len(parts[0]) == 1:
        initial = parts[0]
        first = " ".join(parts[1:])
        return apply_name_case(first), apply_name_case(initial)

    # Trailing single-letter initial: "RAHUL S" -> RAHUL / S
    if len(parts[-1]) == 1:
        initial = parts[-1]
        first = " ".join(parts[:-1])
        return apply_name_case(first), apply_name_case(initial)

    # Otherwise treat the first word as first name, the rest as last name.
    return apply_name_case(parts[0]), apply_name_case(" ".join(parts[1:]))


# ----------------------------------------------------------------------------
# STEP 3: Validate a phone number.
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
# STEP 4: Does an email look like a real email address?
# ----------------------------------------------------------------------------
def is_valid_email(email):
    text = as_text(email)
    pattern = r"^[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}$"
    return re.match(pattern, text) is not None


# ----------------------------------------------------------------------------
# STEP 5: Standardize a date of birth into YYYY-MM-DD text.
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
# STEP 6: Standardize gender to M / F / O.
#   - Male->M, Female->F, Other->O
#   - blank / unknown -> guess from the teacher's first name
#   - if still unknown -> DEFAULT_GENDER (blank by default)
# ----------------------------------------------------------------------------
def standardize_gender(gender, first_name):
    text = as_text(gender).lower()
    if text in ("m", "male"):
        return "M"
    if text in ("f", "female"):
        return "F"
    if text in ("o", "other"):
        return "O"
    # blank / unknown -> try to guess from the first name.
    first = as_text(first_name).split()[0].lower() if as_text(first_name) else ""
    if first in MALE_NAMES:
        return "M"
    if first in FEMALE_NAMES:
        return "F"
    return DEFAULT_GENDER


# ----------------------------------------------------------------------------
# STEP 7: Tidy the staff code (trim + uppercase).
# ----------------------------------------------------------------------------
def standardize_staffcode(code):
    return re.sub(r"\s+", "", as_text(code)).upper()


# ----------------------------------------------------------------------------
# Run all single-row cleaning steps and record each correction.
# ----------------------------------------------------------------------------
def clean_each_teacher(teachers, changes):
    for i, t in enumerate(teachers):
        # Name -> firstname + lastname (initial)
        orig_first = t[ROLE_FIRST]
        orig_last = t[ROLE_LAST]
        new_first, new_last = clean_teacher_name(orig_first, orig_last)
        if new_first != orig_first or new_last != orig_last:
            log_change(changes, i, ROLE_FIRST, orig_first, new_first,
                       "Teacher name standardized.")
        t[ROLE_FIRST] = new_first
        t[ROLE_LAST] = new_last

        # Date of birth
        orig_dob = t[ROLE_DOB]
        new_dob = standardize_dob(orig_dob)
        if new_dob != orig_dob:
            log_change(changes, i, ROLE_DOB, orig_dob, new_dob,
                       "DOB format standardized.")
        t[ROLE_DOB] = new_dob

        # Gender
        orig_gender = t[ROLE_GENDER]
        new_gender = standardize_gender(orig_gender, t[ROLE_FIRST])
        if new_gender != orig_gender:
            log_change(changes, i, ROLE_GENDER, orig_gender, new_gender,
                       "Gender standardized or inferred.")
        t[ROLE_GENDER] = new_gender

        # Staff code
        orig_code = t[ROLE_STAFFCODE]
        new_code = standardize_staffcode(orig_code)
        if new_code != orig_code:
            log_change(changes, i, ROLE_STAFFCODE, orig_code, new_code,
                       "Staff code standardized.")
        t[ROLE_STAFFCODE] = new_code


# ----------------------------------------------------------------------------
# Clean phone numbers for every teacher (duplicates handled later).
# ----------------------------------------------------------------------------
def clean_phone_numbers(teachers, changes, phone_maker):
    for i, t in enumerate(teachers):
        original = t[ROLE_PHONE]
        clean, had_multiple = validate_phone(original)
        if clean != "":
            if had_multiple:
                log_change(changes, i, ROLE_PHONE, original, clean,
                           "Multiple phone numbers found. First valid phone retained.")
            t[ROLE_PHONE] = clean
        else:
            fake = phone_maker.get_next()
            log_change(changes, i, ROLE_PHONE, original, fake,
                       "Invalid phone number replaced with generated phone number.")
            t[ROLE_PHONE] = fake


# ----------------------------------------------------------------------------
# Clean emails for every teacher (duplicates handled later).
# ----------------------------------------------------------------------------
def clean_emails(teachers, changes, email_maker):
    for i, t in enumerate(teachers):
        original = t[ROLE_EMAIL]
        if is_valid_email(original):
            lowered = original.lower()       # emails are always lower-case
            if lowered != original:
                log_change(changes, i, ROLE_EMAIL, original, lowered,
                           "Email converted to lowercase.")
            t[ROLE_EMAIL] = lowered
            email_maker.used.add(lowered)
            continue
        new_email = email_maker.make_unique(t[ROLE_FIRST], t[ROLE_LAST])
        log_change(changes, i, ROLE_EMAIL, original, new_email,
                   "Invalid email format corrected with unique mailinator email.")
        t[ROLE_EMAIL] = new_email


# ----------------------------------------------------------------------------
# DUPLICATE HANDLING.
#   - same phone OR same email = duplicate; first kept (Original), rest Changed
#   - these corrections go ONLY into the duplicates report, never changes report
# Returns a list of duplicate-report rows.
# ----------------------------------------------------------------------------
def handle_duplicates(teachers, phone_maker, email_maker):
    phone_groups = {}
    for i, t in enumerate(teachers):
        phone_groups.setdefault(t[ROLE_PHONE], []).append(i)
    phone_groups = {p: idx for p, idx in phone_groups.items() if len(idx) > 1}

    email_groups = {}
    for i, t in enumerate(teachers):
        email_groups.setdefault(t[ROLE_EMAIL].lower(), []).append(i)
    email_groups = {e: idx for e, idx in email_groups.items() if len(idx) > 1}

    email_member_sets = {frozenset(idx) for idx in email_groups.values()}

    report_rows = []
    group_counter = [0]
    handled_sets = set()

    def emit_group(members, dtype, change_phone, change_email):
        group_counter[0] += 1
        group_id = f"G{group_counter[0]:03d}"
        for position, idx in enumerate(members):
            t = teachers[idx]
            original_phone = t[ROLE_PHONE]
            original_email = t[ROLE_EMAIL]
            new_phone = ""
            new_email = ""
            if position == 0:
                status = "Original"
                comment = "Original teacher record kept unchanged"
            else:
                status = "Changed"
                if change_phone:
                    new_phone = phone_maker.get_next()
                    t[ROLE_PHONE] = new_phone
                if change_email:
                    new_email = email_maker.make_unique(t[ROLE_FIRST], t[ROLE_LAST])
                    t[ROLE_EMAIL] = new_email
                if change_phone and change_email:
                    comment = "Duplicate phone and email changed"
                elif change_phone:
                    comment = "Duplicate phone changed to generated phone"
                else:
                    comment = "Duplicate email changed to mailinator email"

            row = {
                "Duplicate Group ID": group_id,
                "Duplicate Type": dtype,
            }
            for col in REPORT_ID_COLUMNS:
                row[col] = t[col]
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
# Build the changes-report table (using each teacher's final details).
# ----------------------------------------------------------------------------
def build_changes_report(teachers, changes):
    rows = []
    for change in changes:
        t = teachers[change["teacher_index"]]
        row = {col: t[col] for col in REPORT_ID_COLUMNS}
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
def build_cleaned_table(teachers):
    df = pd.DataFrame(teachers)
    df = df[FINAL_COLUMNS]
    df = df.sort_values(by=[ROLE_FIRST, ROLE_LAST], kind="stable")
    df = df.reset_index(drop=True)
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
    # openpyxl types wb.active as Worksheet | None. A workbook we just loaded
    # (or just saved with pandas) always has an active sheet, so this assert
    # tells Pylance (and confirms at runtime) that ws is never None from here
    # on -- this resolves every "not a known attribute of None" warning below.
    assert ws is not None, f"Workbook at {file_path} has no active worksheet"

    bold = Font(bold=True)
    center = Alignment(horizontal="center", vertical="center")
    yellow = PatternFill(start_color="FFFF00", end_color="FFFF00",
                         fill_type="solid")

    for cell in ws[1]:
        cell.font = bold
        cell.alignment = center

    if ws.max_row >= 1 and ws.max_column >= 1:
        ws.auto_filter.ref = f"A1:{get_column_letter(ws.max_column)}{ws.max_row}"

    for col_cells in ws.columns:
        longest = 0
        col_index = col_cells[0].column
        assert col_index is not None
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
# THE MAIN PROGRAM - runs when you type: python teacher_data_cleaner.py
# ----------------------------------------------------------------------------
def main():
    print("Teacher Data Cleaner - starting...")

    if not os.path.exists(OUTPUT_FOLDER_PATH):
        os.makedirs(OUTPUT_FOLDER_PATH)
        print(f"Created output folder: {OUTPUT_FOLDER_PATH}")

    teachers = read_and_prepare(INPUT_FILE_PATH)
    print(f"Read {len(teachers)} teacher rows.")

    changes = []
    clean_each_teacher(teachers, changes)

    existing_phones = [t[ROLE_PHONE] for t in teachers]
    phone_maker = FakePhoneMaker(FAKE_PHONE_START, existing_phones)
    email_maker = EmailMaker([])

    clean_phone_numbers(teachers, changes, phone_maker)
    clean_emails(teachers, changes, email_maker)

    duplicate_rows = handle_duplicates(teachers, phone_maker, email_maker)

    cleaned_df = build_cleaned_table(teachers)
    changes_df = build_changes_report(teachers, changes)
    duplicates_df = build_duplicates_report(duplicate_rows)

    cleaned_path = os.path.join(OUTPUT_FOLDER_PATH, "cleaned_teachers.xlsx")
    changes_path = os.path.join(OUTPUT_FOLDER_PATH, "changes_report.xlsx")
    duplicates_path = os.path.join(OUTPUT_FOLDER_PATH, "duplicates_report.xlsx")

    save_excel(cleaned_df, cleaned_path)
    save_excel(changes_df, changes_path)
    save_excel(duplicates_df, duplicates_path, highlight_changed=True)

    print("\nDone! Files saved in:", os.path.abspath(OUTPUT_FOLDER_PATH))
    print(f"  1) cleaned_teachers.xlsx   ({len(cleaned_df)} teachers)")
    print(f"  2) changes_report.xlsx      ({len(changes_df)} corrections)")
    print(f"  3) duplicates_report.xlsx   ({len(duplicate_rows)} duplicate rows)")


if __name__ == "__main__":
    try:
        main()
    except FileNotFoundError as error:
        print("INPUT FILE PROBLEM:", error)
    except ValueError as error:
        print("COLUMN PROBLEM:", error)
    except PermissionError as error:
        print("PERMISSION PROBLEM:", error)
        print("If an OUTPUT file is open in Excel (cleaned_teachers.xlsx etc.), "
              "close it and run again.")
    except Exception as error:
        print("Something went wrong:", error)