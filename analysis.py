"""
analysis.py - Comprehensive Result Analysis for Bangladesh Board Examinations
Calculates:
- Regular vs Irregular Candidate Performance Analysis & Comparison
- Subject Wise Grade Analysis (A+, A, A-, B, C, D, F, Appeared, Passed, Pass Rate)
- 1-Subject Fail Breakdown and Student Roster
- 2-Subject Fail Breakdown and Student Roster
- 3+ Subject Fail List
- GPA & Group Distributions
- Excel Analysis Workbook Generation
"""

import io
import re
from collections import defaultdict
from typing import Any, Dict, List, Optional
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from eiin_lookup import (
    CURATED_INSTITUTIONS,
    resolve_institution_address_dynamic,
    resolve_institution_details,
    resolve_institution_eiin_dynamic,
)


VALID_GRADES = ["A+", "A", "A-", "B", "C", "D", "F"]

GRADE_POINTS = {
    "A+": 5.0,
    "A": 4.0,
    "A-": 3.5,
    "B": 3.0,
    "C": 2.0,
    "D": 1.0,
    "F": 0.0
}

KNOWN_INSTITUTIONS_REGISTRY = CURATED_INSTITUTIONS


def resolve_institution_eiin(raw_val: str = "", institute: str = "", board: str = "") -> str:
    """
    Dynamically resolves a valid EIIN from explicit records, embedded digits,
    or the comprehensive Bangladesh institutions directory (~29,500+ institutions).
    Returns "" if the institution is not found rather than substituting an unrelated college's EIIN.
    """
    return resolve_institution_eiin_dynamic(raw_val=raw_val, institute=institute, board=board)


def resolve_institution_address(institute: str = "", eiin: str = "", board: str = "") -> str:
    """Resolves the physical address dynamically based on EIIN, institute name, and board."""
    return resolve_institution_address_dynamic(institute=institute, eiin=eiin, board=board)


def identify_fourth_subject(subject_grades: Dict[str, str], group: str) -> Optional[str]:
    """
    Identifies the 4th (optional) subject from a student's enrolled subject list.
    In Bangladesh curriculum, regular students have 3 compulsory subjects (Bangla, English, ICT),
    3 group-core electives, and 1 fourth/optional subject (7 subjects total in HSC/SSC).
    """
    if not subject_grades:
        return None

    # Check if any subject explicitly contains 4th/fourth/optional in name
    for s in subject_grades.keys():
        s_low = s.lower()
        if any(term in s_low for term in ["4th", "fourth", "optional", "additional"]):
            return s

    # Typical 4th subjects in Bangladesh education boards
    typical_fourth_keywords = [
        "agriculture", "higher math", "biology", "statistics", "psychology",
        "home science", "social work", "finance", "logic", "geography",
        "production management", "islamic studies", "music", "arts"
    ]

    subjs = list(subject_grades.keys())

    # If student has 7 or more subjects, the 4th subject is placed last in board tables
    if len(subjs) >= 7:
        last_subj = subjs[-1]
        last_subj_low = last_subj.lower()
        # If last subject is not Bangla, English, or ICT:
        if not any(comp in last_subj_low for comp in ["bangla", "english", "information & communication"]):
            return last_subj

    # If typical keyword found in any subject (starting check from reverse):
    for s in reversed(subjs):
        s_low = s.lower()
        if any(kw in s_low for kw in typical_fourth_keywords):
            return s

    # If 7 or more subjects and no other match, return the last subject
    if len(subjs) >= 7:
        return subjs[-1]

    return None


def normalize_record(raw: Dict[str, Any]) -> Dict[str, Any]:
    """
    Normalizes a student record from scraper dictionary or DataFrame row.
    Extracts Roll, Reg, Name, Group, Type (Regular/Irregular), Result, GPA (With & Without 4th),
    Subject grades, 4th subject, and Absent status.
    """
    roll = str(raw.get("Roll", "")).strip().split(".")[0]
    reg = str(raw.get("Registration", "") or raw.get("Reg", "")).strip().split(".")[0]
    if reg == "nan":
        reg = ""
    name = str(raw.get("Name", "N/A")).strip()
    if name == "nan":
        name = "N/A"

    father_name = str(raw.get("Father Name", "") or raw.get("Father's Name", "")).strip()
    if father_name == "nan":
        father_name = ""

    institute = str(raw.get("Institute", "") or raw.get("Institution", "")).strip()
    if institute == "nan":
        institute = ""

    board = str(raw.get("Board", "") or raw.get("board", "")).strip()
    if board == "nan":
        board = ""

    eiin_raw = str(raw.get("EIIN", "") or raw.get("eiin", "") or raw.get("Eiin", "")).strip()
    if eiin_raw == "nan":
        eiin_raw = ""
    eiin = resolve_institution_eiin(eiin_raw, institute, board)

    group = str(raw.get("Group", "") or "GENERAL").strip().upper()
    if not group or group in ["NAN", "NONE", "-"]:
        group = "GENERAL"

    result_status = str(raw.get("Result Status", "") or raw.get("Result", "") or "").strip().upper()
    status_field = str(raw.get("Status", "")).strip().upper()

    gpa_raw = str(raw.get("GPA", "0.00")).strip()
    try:
        gpa = float(gpa_raw)
    except Exception:
        gpa = 0.00

    subject_grades: Dict[str, str] = {}

    # 1. Check if Subject_Grades dict is directly present
    if isinstance(raw.get("Subject_Grades"), dict) and raw["Subject_Grades"]:
        for k, v in raw["Subject_Grades"].items():
            g = str(v).strip().upper()
            if g in VALID_GRADES or g in ["ABS", "ABSENT"]:
                subject_grades[str(k).strip()] = g

    # 2. Check if Subject Grades Summary string is present (e.g. "101 BANGLA: F, 107 ENGLISH: B")
    summary_str = str(raw.get("Subject Grades Summary", "") or raw.get("Subject_Grades_Summary", "")).strip()
    if not subject_grades and summary_str and summary_str != "nan":
        parts = summary_str.split(",")
        for part in parts:
            if ":" in part:
                s_name, s_grade = part.split(":", 1)
                s_grade = s_grade.strip().upper()
                if s_grade in VALID_GRADES or s_grade in ["ABS", "ABSENT"]:
                    subject_grades[s_name.strip()] = s_grade

    # 3. Check individual column keys
    if not subject_grades:
        exclude_keys = {
            "roll", "registration", "reg", "name", "father name", "mother name",
            "institute", "eiin", "board", "group", "exam", "year", "result status",
            "result", "gpa", "status", "subject grades summary", "subject_grades_summary",
            "type", "session", "created_at"
        }
        for k, v in raw.items():
            if str(k).lower().strip() not in exclude_keys:
                grade_val = str(v).strip().upper()
                if grade_val in VALID_GRADES or grade_val in ["ABS", "ABSENT"]:
                    subject_grades[str(k).strip()] = grade_val

    # Detect candidate type (Regular / Irregular / Improvement / Private)
    cand_type_raw = str(
        raw.get("Type") or raw.get("Candidate Type") or raw.get("Student Type") or ""
    ).strip().upper()

    if not cand_type_raw or cand_type_raw in ["NAN", "NONE", "-"]:
        if 0 < len(subject_grades) <= 4:
            cand_type = "IRREGULAR"
        else:
            cand_type = "REGULAR"
    elif "IRREG" in cand_type_raw:
        cand_type = "IRREGULAR"
    elif "REG" in cand_type_raw:
        cand_type = "REGULAR"
    elif "IMP" in cand_type_raw:
        cand_type = "IMPROVEMENT"
    elif "PRIV" in cand_type_raw:
        cand_type = "PRIVATE"
    else:
        cand_type = cand_type_raw

    # Identify 4th Subject
    fourth_subj = identify_fourth_subject(subject_grades, group)
    fourth_grade = subject_grades.get(fourth_subj) if fourth_subj else None
    fourth_gp = GRADE_POINTS.get(fourth_grade, 0.0) if fourth_grade else 0.0

    # Separate compulsory/core subjects from 4th subject
    main_subjects = {k: v for k, v in subject_grades.items() if k != fourth_subj}

    # Detect absent status
    is_absent = (
        ("ABSENT" in result_status) or
        (result_status in ["ABS", "A"]) or
        ("ABSENT" in status_field) or
        ("ABSENT" in summary_str.upper()) or
        (len(subject_grades) > 0 and all(g in ["ABS", "ABSENT", "-"] for g in subject_grades.values()))
    )

    # Detect failed subjects (only grades == 'F')
    failed_subjects = [s for s, g in subject_grades.items() if g == "F"]
    passed_subjects = [f"{s} ({g})" for s, g in subject_grades.items() if g != "F" and g not in ["ABS", "ABSENT"]]

    # In Bangladesh board rules: Failing 4th subject does NOT fail the student if all main subjects are passed!
    main_failed_subjects = [s for s in failed_subjects if s != fourth_subj]
    is_fail = False
    is_passed = False

    if is_absent:
        final_status = "ABSENT"
    else:
        if ("FAIL" in result_status) or (len(main_failed_subjects) > 0):
            is_fail = True
            final_status = "FAIL"
        elif gpa > 0.0 or result_status == "PASSED":
            is_passed = True
            final_status = "PASSED"
        else:
            final_status = result_status if result_status else "UNKNOWN"

    # Calculate GPA With 4th Subject and GPA Without 4th Subject
    if is_passed:
        # GPA Without 4th Subject: strictly on the compulsory & core subjects (Golden A+ = >= 5.00 without 4th)
        if len(main_subjects) > 0:
            sum_main_gp = sum(GRADE_POINTS.get(g, 0.0) for g in main_subjects.values())
            gpa_without_4th = min(5.00, round(sum_main_gp / len(main_subjects), 2))
        else:
            gpa_without_4th = gpa

        # GPA With 4th Subject: official board GPA or computed with 4th subject bonus
        bonus = max(0.0, fourth_gp - 2.0) if fourth_subj else 0.0
        if len(main_subjects) > 0:
            sum_main_gp = sum(GRADE_POINTS.get(g, 0.0) for g in main_subjects.values())
            computed_w4 = min(5.00, round((sum_main_gp + bonus) / len(main_subjects), 2))
            gpa_with_4th = max(gpa, computed_w4) if gpa > 0.0 else computed_w4
        else:
            gpa_with_4th = gpa
    else:
        gpa_with_4th = 0.00
        gpa_without_4th = 0.00

    is_gpa5_with_4th = is_passed and gpa_with_4th >= 5.00
    is_gpa5_without_4th = is_passed and gpa_without_4th >= 5.00

    return {
        "Roll": roll,
        "Registration": reg,
        "Name": name,
        "Father Name": father_name,
        "Institute": institute,
        "EIIN": eiin,
        "Board": board.title() if board else "N/A",
        "Group": group,
        "Type": cand_type,
        "Result": final_status,
        "GPA": f"{gpa_with_4th:.2f}",
        "GPANumeric": gpa_with_4th,
        "GPA_With_4th": f"{gpa_with_4th:.2f}",
        "GPA_Without_4th": f"{gpa_without_4th:.2f}",
        "GPA_With_4th_Numeric": gpa_with_4th,
        "GPA_Without_4th_Numeric": gpa_without_4th,
        "Fourth_Subject": fourth_subj,
        "Fourth_Subject_Grade": fourth_grade,
        "Is_GPA5_With_4th": is_gpa5_with_4th,
        "Is_GPA5_Without_4th": is_gpa5_without_4th,
        "Is_Absent": is_absent,
        "Subject_Grades": subject_grades,
        "Failed_Subjects": failed_subjects,
        "Failed_Count": len(failed_subjects),
        "Main_Failed_Subjects": main_failed_subjects,
        "Main_Failed_Count": len(main_failed_subjects),
        "Passed_Grades_Summary": ", ".join(passed_subjects) if passed_subjects else "None",
        "Raw": raw
    }


def build_fail_subject_summary_matrix(normalized_records: List[Dict[str, Any]]) -> Dict[str, Any]:
    """
    Constructs the official Fail Subject Summary Matrix (Without 4th Subject)
    categorizing students by exact number of failed main subjects:
    - One Subject
    - Two Subject
    - Three Subject
    - Four Subject
    - Five or more Subject
    across Academic Groups (Science, Business Studies, Humanities, etc.)
    and Candidate Types (Regular, Irregular) with SubTotals and Grand Totals.
    """
    def map_matrix_group(raw_grp: str) -> str:
        g = str(raw_grp or "").strip().upper()
        if "SCI" in g:
            return "Science"
        elif "BUS" in g or "COMMERCE" in g or "B.STUDIES" in g or "MGT" in g:
            return "Business Studies"
        elif "HUMAN" in g or "ARTS" in g:
            return "Humanities"
        elif not g or g in ["NAN", "NONE", "-", "GENERAL"]:
            return "General"
        else:
            return g.title()

    core_groups = ["Science", "Business Studies", "Humanities"]

    failing_students = []
    detected_extra = []

    for st in normalized_records:
        res = st.get("Result", "")
        if res == "ABSENT":
            continue
        fails = st.get("Main_Failed_Subjects") if "Main_Failed_Subjects" in st else [
            s for s in st.get("Failed_Subjects", []) if s != st.get("Fourth_Subject")
        ]
        fail_cnt = len(fails)
        if fail_cnt >= 1:
            m_grp = map_matrix_group(st.get("Group", ""))
            if m_grp not in core_groups and m_grp not in detected_extra:
                detected_extra.append(m_grp)
            failing_students.append({
                "fail_cnt": fail_cnt,
                "group": m_grp,
                "is_reg": (st.get("Type", "") == "REGULAR")
            })

    all_groups = core_groups + sorted(detected_extra)

    row_specs = [
        ("one_subject", "One Subject", lambda c: c == 1),
        ("two_subject", "Two Subject", lambda c: c == 2),
        ("three_subject", "Three Subject", lambda c: c == 3),
        ("four_subject", "Four Subject", lambda c: c == 4),
        ("five_plus_subject", "Five or more Subject", lambda c: c >= 5),
    ]

    matrix_rows = []
    for row_key, row_label, pred in row_specs:
        counts = {}
        row_sub_reg = 0
        row_sub_irreg = 0
        for grp in all_groups:
            reg_c = sum(1 for s in failing_students if pred(s["fail_cnt"]) and s["group"] == grp and s["is_reg"])
            irreg_c = sum(1 for s in failing_students if pred(s["fail_cnt"]) and s["group"] == grp and not s["is_reg"])
            counts[grp] = {
                "regular": reg_c,
                "irregular": irreg_c,
                "total": reg_c + irreg_c
            }
            row_sub_reg += reg_c
            row_sub_irreg += irreg_c
        row_total = row_sub_reg + row_sub_irreg
        matrix_rows.append({
            "key": row_key,
            "label": row_label,
            "counts": counts,
            "subtotal": {
                "regular": row_sub_reg,
                "irregular": row_sub_irreg,
                "total": row_total
            },
            "total": row_total
        })

    # SubTotal row (Row 6)
    subtotal_counts = {}
    grand_sub_reg = 0
    grand_sub_irreg = 0
    for grp in all_groups:
        grp_reg_tot = sum(r["counts"][grp]["regular"] for r in matrix_rows)
        grp_irreg_tot = sum(r["counts"][grp]["irregular"] for r in matrix_rows)
        subtotal_counts[grp] = {
            "regular": grp_reg_tot,
            "irregular": grp_irreg_tot,
            "total": grp_reg_tot + grp_irreg_tot
        }
        grand_sub_reg += grp_reg_tot
        grand_sub_irreg += grp_irreg_tot

    grand_total = grand_sub_reg + grand_sub_irreg
    subtotal_row = {
        "label": "SubTotal",
        "counts": subtotal_counts,
        "subtotal": {
            "regular": grand_sub_reg,
            "irregular": grand_sub_irreg,
            "total": grand_total
        },
        "total": grand_total
    }

    # Total row (Row 7)
    total_counts = {}
    for grp in all_groups:
        total_counts[grp] = subtotal_counts[grp]["total"]

    total_row = {
        "label": "Total",
        "counts": total_counts,
        "subtotal": grand_total,
        "total": grand_total
    }

    return {
        "groups": all_groups,
        "rows": matrix_rows,
        "subtotal_row": subtotal_row,
        "total_row": total_row
    }


def compute_result_analysis(records: List[Dict[str, Any]]) -> Dict[str, Any]:
    """
    Computes all analytical aggregations across the given list of student records,
    including Regular vs Irregular candidate analyses.
    """
    normalized = [normalize_record(r) for r in records if r.get("Roll") or r.get("roll")]
    total_students = len(normalized)

    if total_students == 0:
        return {
            "total_students": 0,
            "summary": {},
            "groups_summary": {},
            "types_summary": {},
            "regular_summary": {},
            "irregular_summary": {},
            "group_type_comparison": [],
            "subject_analysis": [],
            "merit_list": {"total": 0, "students": []},
            "failed_list_subject_wise": {
                "total_failed_students": 0,
                "total_subject_failures": 0,
                "subjects": [],
                "entries": []
            },
            "one_subject_fail": {"total": 0, "subjects_summary": {}, "students": []},
            "two_subject_fail": {"total": 0, "subjects_summary": {}, "pair_summary": {}, "students": []},
            "multi_subject_fail": {"total": 0, "students": []},
            "fail_subject_summary_matrix": {
                "groups": ["Science", "Business Studies", "Humanities"],
                "rows": [],
                "subtotal_row": {"label": "SubTotal", "counts": {}, "subtotal": {"regular": 0, "irregular": 0, "total": 0}, "total": 0},
                "total_row": {"label": "Total", "counts": {}, "subtotal": 0, "total": 0}
            },
            "gpa_distribution": {},
            "available_groups": [],
            "available_types": []
        }

    # Tracking containers
    passed_students = []
    failed_students = []
    absent_students = []
    other_students = []

    one_subject_fail_students = []
    two_subject_fail_students = []
    multi_subject_fail_students = []

    # Frequency maps - All
    one_sub_freq = defaultdict(int)
    two_sub_freq = defaultdict(int)
    two_sub_pairs = defaultdict(int)
    three_plus_sub_freq = defaultdict(int)

    # Frequency maps with Regular & Irregular breakdown - All
    one_sub_type_freq = defaultdict(lambda: {"regular": 0, "irregular": 0, "total": 0})
    two_sub_type_freq = defaultdict(lambda: {"regular": 0, "irregular": 0, "total": 0})
    three_plus_type_freq = defaultdict(lambda: {"regular": 0, "irregular": 0, "total": 0})

    # Frequency maps with Regular & Irregular breakdown - Group-wise
    one_sub_group_freq = defaultdict(lambda: defaultdict(int))
    two_sub_group_freq = defaultdict(lambda: defaultdict(int))
    three_plus_group_freq = defaultdict(lambda: defaultdict(int))
    one_sub_group_type_freq = defaultdict(lambda: defaultdict(lambda: {"regular": 0, "irregular": 0, "total": 0}))
    two_sub_group_type_freq = defaultdict(lambda: defaultdict(lambda: {"regular": 0, "irregular": 0, "total": 0}))
    three_plus_group_type_freq = defaultdict(lambda: defaultdict(lambda: {"regular": 0, "irregular": 0, "total": 0}))

    # GPA Distributions (Both With 4th and Without 4th subject)
    gpa_dist_with_4th = {
        "5.00 (A+)": 0,
        "4.00 - 4.99 (A)": 0,
        "3.50 - 3.99 (A-)": 0,
        "3.00 - 3.49 (B)": 0,
        "2.00 - 2.99 (C)": 0,
        "1.00 - 1.99 (D)": 0,
        "0.00 (F / Fail)": 0
    }
    gpa_dist_without_4th = {
        "5.00 (A+)": 0,
        "4.00 - 4.99 (A)": 0,
        "3.50 - 3.99 (A-)": 0,
        "3.00 - 3.49 (B)": 0,
        "2.00 - 2.99 (C)": 0,
        "1.00 - 1.99 (D)": 0,
        "0.00 (F / Fail)": 0
    }

    # Group tracking: group_name -> list of students
    group_students = defaultdict(list)

    # Type tracking: type_name -> list of students
    type_students = defaultdict(list)

    # Subject tracking per group & candidate type:
    # (group, type, subject) -> {"A+": 0, "A": 0, ...}
    subject_stats = defaultdict(lambda: {g: 0 for g in VALID_GRADES})

    for st in normalized:
        grp = st["Group"]
        c_type = st["Type"]
        group_students[grp].append(st)
        type_students[c_type].append(st)

        gpa_w4 = st["GPA_With_4th_Numeric"]
        gpa_no4 = st["GPA_Without_4th_Numeric"]
        res = st["Result"]
        # Failed subjects strictly Without 4th Subject
        fails = st.get("Main_Failed_Subjects") if "Main_Failed_Subjects" in st else [s for s in st.get("Failed_Subjects", []) if s != st.get("Fourth_Subject")]
        fail_cnt = len(fails)
        is_p = (res == "PASSED")

        # Result grouping
        if res == "PASSED":
            passed_students.append(st)
        elif res == "FAIL":
            failed_students.append(st)
        elif res == "ABSENT":
            absent_students.append({
                "Roll": st["Roll"],
                "Registration": st["Registration"],
                "Name": st["Name"],
                "Father Name": st.get("Father Name", ""),
                "Institute": st.get("Institute", ""),
                "Board": st.get("Board", "N/A"),
                "Group": grp,
                "Type": c_type,
                "Result": "ABSENT",
                "Status": "Absent in Examination",
                "Remarks": "Did not appear"
            })
        else:
            other_students.append(st)

        # Subject-fail categorization (Without 4th Subject - only for students who appeared and failed)
        is_reg = (c_type == "REGULAR")
        if res != "ABSENT":
            if fail_cnt == 1:
                s = fails[0]
                one_subject_fail_students.append({
                    "Roll": st["Roll"],
                    "Registration": st["Registration"],
                    "Name": st["Name"],
                    "Board": st.get("Board", "N/A"),
                    "Group": grp,
                    "Type": c_type,
                    "Failed_Subject": s,
                    "Passed_Grades_Summary": st["Passed_Grades_Summary"],
                    "GPA": st["GPA_With_4th"],
                    "GPA_With_4th": st["GPA_With_4th"],
                    "GPA_Without_4th": st["GPA_Without_4th"]
                })
                one_sub_freq[s] += 1
                one_sub_group_freq[grp][s] += 1
                one_sub_type_freq[s]["total"] += 1
                one_sub_group_type_freq[grp][s]["total"] += 1
                if is_reg:
                    one_sub_type_freq[s]["regular"] += 1
                    one_sub_group_type_freq[grp][s]["regular"] += 1
                else:
                    one_sub_type_freq[s]["irregular"] += 1
                    one_sub_group_type_freq[grp][s]["irregular"] += 1
            elif fail_cnt == 2:
                two_subject_fail_students.append({
                    "Roll": st["Roll"],
                    "Registration": st["Registration"],
                    "Name": st["Name"],
                    "Board": st.get("Board", "N/A"),
                    "Group": grp,
                    "Type": c_type,
                    "Failed_Subjects": fails,
                    "Passed_Grades_Summary": st["Passed_Grades_Summary"],
                    "GPA": st["GPA_With_4th"],
                    "GPA_With_4th": st["GPA_With_4th"],
                    "GPA_Without_4th": st["GPA_Without_4th"]
                })
                for s in fails:
                    two_sub_freq[s] += 1
                    two_sub_group_freq[grp][s] += 1
                    two_sub_type_freq[s]["total"] += 1
                    two_sub_group_type_freq[grp][s]["total"] += 1
                    if is_reg:
                        two_sub_type_freq[s]["regular"] += 1
                        two_sub_group_type_freq[grp][s]["regular"] += 1
                    else:
                        two_sub_type_freq[s]["irregular"] += 1
                        two_sub_group_type_freq[grp][s]["irregular"] += 1
                pair_key = " & ".join(sorted(fails))
                two_sub_pairs[pair_key] += 1
            elif fail_cnt >= 3:
                multi_subject_fail_students.append({
                    "Roll": st["Roll"],
                    "Registration": st["Registration"],
                    "Name": st["Name"],
                    "Board": st.get("Board", "N/A"),
                    "Group": grp,
                    "Type": c_type,
                    "Failed_Count": fail_cnt,
                    "Failed_Subjects": fails,
                    "Passed_Grades_Summary": st["Passed_Grades_Summary"],
                    "GPA": st["GPA_With_4th"],
                    "GPA_With_4th": st["GPA_With_4th"],
                    "GPA_Without_4th": st["GPA_Without_4th"]
                })
                for s in fails:
                    three_plus_sub_freq[s] += 1
                    three_plus_group_freq[grp][s] += 1
                    three_plus_type_freq[s]["total"] += 1
                    three_plus_group_type_freq[grp][s]["total"] += 1
                    if is_reg:
                        three_plus_type_freq[s]["regular"] += 1
                        three_plus_group_type_freq[grp][s]["regular"] += 1
                    else:
                        three_plus_type_freq[s]["irregular"] += 1
                        three_plus_group_type_freq[grp][s]["irregular"] += 1

        # GPA distribution: With 4th Subject
        if res == "FAIL" or fail_cnt > 0 or not is_p:
            gpa_dist_with_4th["0.00 (F / Fail)"] += 1
        elif gpa_w4 >= 5.00:
            gpa_dist_with_4th["5.00 (A+)"] += 1
        elif gpa_w4 >= 4.00:
            gpa_dist_with_4th["4.00 - 4.99 (A)"] += 1
        elif gpa_w4 >= 3.50:
            gpa_dist_with_4th["3.50 - 3.99 (A-)"] += 1
        elif gpa_w4 >= 3.00:
            gpa_dist_with_4th["3.00 - 3.49 (B)"] += 1
        elif gpa_w4 >= 2.00:
            gpa_dist_with_4th["2.00 - 2.99 (C)"] += 1
        elif gpa_w4 >= 1.00:
            gpa_dist_with_4th["1.00 - 1.99 (D)"] += 1
        else:
            gpa_dist_with_4th["0.00 (F / Fail)"] += 1

        # GPA distribution: Without 4th Subject (Pure / Golden GPA)
        if res == "FAIL" or fail_cnt > 0 or not is_p:
            gpa_dist_without_4th["0.00 (F / Fail)"] += 1
        elif gpa_no4 >= 5.00:
            gpa_dist_without_4th["5.00 (A+)"] += 1
        elif gpa_no4 >= 4.00:
            gpa_dist_without_4th["4.00 - 4.99 (A)"] += 1
        elif gpa_no4 >= 3.50:
            gpa_dist_without_4th["3.50 - 3.99 (A-)"] += 1
        elif gpa_no4 >= 3.00:
            gpa_dist_without_4th["3.00 - 3.49 (B)"] += 1
        elif gpa_no4 >= 2.00:
            gpa_dist_without_4th["2.00 - 2.99 (C)"] += 1
        elif gpa_no4 >= 1.00:
            gpa_dist_without_4th["1.00 - 1.99 (D)"] += 1
        else:
            gpa_dist_without_4th["0.00 (F / Fail)"] += 1

        # Subject-wise tracking across group and candidate type
        if res != "ABSENT":
            for subj, gr in st["Subject_Grades"].items():
                if gr in VALID_GRADES:
                    subject_stats[(grp, c_type, subj)][gr] += 1
                    subject_stats[("ALL", c_type, subj)][gr] += 1
                    subject_stats[(grp, "ALL", subj)][gr] += 1
                    subject_stats[("ALL", "ALL", subj)][gr] += 1

    # Overall Summary calculation
    total_passed = len(passed_students)
    total_failed = len(failed_students)
    total_absent = len(absent_students)
    appeared = total_students - total_absent
    pass_rate = round((total_passed / appeared * 100), 2) if appeared > 0 else 0.0
    fail_rate = round((total_failed / appeared * 100), 2) if appeared > 0 else 0.0
    absent_rate = round((total_absent / total_students * 100), 2) if total_students > 0 else 0.0

    # GPA 5.0 With 4th Subject vs Without 4th Subject
    gpa_5_count_with_4th = sum(1 for s in passed_students if s["Is_GPA5_With_4th"])
    gpa_5_rate_with_4th = round((gpa_5_count_with_4th / appeared * 100), 2) if appeared > 0 else 0.0

    gpa_5_count_without_4th = sum(1 for s in passed_students if s["Is_GPA5_Without_4th"])
    gpa_5_rate_without_4th = round((gpa_5_count_without_4th / appeared * 100), 2) if appeared > 0 else 0.0

    # Group Summary calculation
    groups_summary = {}
    for grp, s_list in group_students.items():
        grp_tot = len(s_list)
        grp_pass = sum(1 for s in s_list if s["Result"] == "PASSED")
        grp_fail = sum(1 for s in s_list if s["Result"] == "FAIL")
        grp_absent = sum(1 for s in s_list if s["Result"] == "ABSENT")
        grp_app = grp_tot - grp_absent
        grp_gpa5_w4 = sum(1 for s in s_list if s["Is_GPA5_With_4th"])
        grp_gpa5_no4 = sum(1 for s in s_list if s["Is_GPA5_Without_4th"])
        grp_1fail = sum(1 for s in s_list if s["Failed_Count"] == 1 and s["Result"] != "ABSENT")
        grp_2fail = sum(1 for s in s_list if s["Failed_Count"] == 2 and s["Result"] != "ABSENT")
        grp_3fail = sum(1 for s in s_list if s["Failed_Count"] >= 3 and s["Result"] != "ABSENT")
        grp_rate = round((grp_pass / grp_app * 100), 2) if grp_app > 0 else 0.0
        grp_absent_rate = round((grp_absent / grp_tot * 100), 2) if grp_tot > 0 else 0.0

        groups_summary[grp] = {
            "group": grp,
            "total": grp_tot,
            "appeared": grp_app,
            "passed": grp_pass,
            "failed": grp_fail,
            "absent": grp_absent,
            "absent_rate": grp_absent_rate,
            "pass_rate": grp_rate,
            "gpa5_count": grp_gpa5_w4,
            "gpa5_count_with_4th": grp_gpa5_w4,
            "gpa5_count_without_4th": grp_gpa5_no4,
            "one_fail_count": grp_1fail,
            "two_fail_count": grp_2fail,
            "multi_fail_count": grp_3fail
        }

    # Candidate Type Summary (Regular vs Irregular)
    types_summary = {}
    for c_type, s_list in type_students.items():
        t_tot = len(s_list)
        t_pass = sum(1 for s in s_list if s["Result"] == "PASSED")
        t_fail = sum(1 for s in s_list if s["Result"] == "FAIL")
        t_abs = sum(1 for s in s_list if s["Result"] == "ABSENT")
        t_app = t_tot - t_abs
        t_pass_rate = round((t_pass / t_app * 100), 2) if t_app > 0 else 0.0
        t_fail_rate = round((t_fail / t_app * 100), 2) if t_app > 0 else 0.0
        t_abs_rate = round((t_abs / t_tot * 100), 2) if t_tot > 0 else 0.0
        t_gpa5_w4 = sum(1 for s in s_list if s["Is_GPA5_With_4th"])
        t_gpa5_no4 = sum(1 for s in s_list if s["Is_GPA5_Without_4th"])
        t_gpa5_rate_w4 = round((t_gpa5_w4 / t_app * 100), 2) if t_app > 0 else 0.0
        t_gpa5_rate_no4 = round((t_gpa5_no4 / t_app * 100), 2) if t_app > 0 else 0.0
        t_1fail = sum(1 for s in s_list if s["Failed_Count"] == 1 and s["Result"] != "ABSENT")
        t_2fail = sum(1 for s in s_list if s["Failed_Count"] == 2 and s["Result"] != "ABSENT")
        t_mfail = sum(1 for s in s_list if s["Failed_Count"] >= 3 and s["Result"] != "ABSENT")

        # Group-wise breakdown inside this type
        grp_breakdown = {}
        for s in s_list:
            g = s["Group"]
            if g not in grp_breakdown:
                grp_breakdown[g] = {"total": 0, "passed": 0, "failed": 0, "absent": 0, "gpa5_with_4th": 0, "gpa5_without_4th": 0}
            grp_breakdown[g]["total"] += 1
            if s["Result"] == "PASSED":
                grp_breakdown[g]["passed"] += 1
                if s["Is_GPA5_With_4th"]:
                    grp_breakdown[g]["gpa5_with_4th"] += 1
                if s["Is_GPA5_Without_4th"]:
                    grp_breakdown[g]["gpa5_without_4th"] += 1
            elif s["Result"] == "FAIL":
                grp_breakdown[g]["failed"] += 1
            elif s["Result"] == "ABSENT":
                grp_breakdown[g]["absent"] += 1

        for g_data in grp_breakdown.values():
            g_data["pass_rate"] = round((g_data["passed"] / g_data["total"] * 100), 2) if g_data["total"] > 0 else 0.0

        types_summary[c_type] = {
            "type": c_type,
            "total": t_tot,
            "appeared": t_app,
            "passed": t_pass,
            "failed": t_fail,
            "absent": t_abs,
            "absent_rate": t_abs_rate,
            "pass_rate": t_pass_rate,
            "fail_rate": t_fail_rate,
            "gpa5_count": t_gpa5_w4,
            "gpa5_count_with_4th": t_gpa5_w4,
            "gpa5_count_without_4th": t_gpa5_no4,
            "gpa5_rate": t_gpa5_rate_w4,
            "gpa5_rate_with_4th": t_gpa5_rate_w4,
            "gpa5_rate_without_4th": t_gpa5_rate_no4,
            "one_fail_count": t_1fail,
            "two_fail_count": t_2fail,
            "multi_fail_count": t_mfail,
            "group_breakdown": grp_breakdown
        }

    # Cross-tab comparison: Group x Type (e.g. Science Regular vs Science Irregular)
    group_type_comparison = []
    for grp in sorted(list(group_students.keys())):
        for c_type in ["REGULAR", "IRREGULAR", "IMPROVEMENT", "PRIVATE"]:
            s_sub = [s for s in group_students[grp] if s["Type"] == c_type]
            if s_sub:
                sub_tot = len(s_sub)
                sub_pass = sum(1 for s in s_sub if s["Result"] == "PASSED")
                sub_fail = sum(1 for s in s_sub if s["Result"] == "FAIL")
                sub_abs = sum(1 for s in s_sub if s["Result"] == "ABSENT")
                sub_app = sub_tot - sub_abs
                sub_rate = round((sub_pass / sub_app * 100), 2) if sub_app > 0 else 0.0
                sub_gpa5_w4 = sum(1 for s in s_sub if s["Is_GPA5_With_4th"])
                sub_gpa5_no4 = sum(1 for s in s_sub if s["Is_GPA5_Without_4th"])
                group_type_comparison.append({
                    "group": grp,
                    "type": c_type,
                    "total": sub_tot,
                    "appeared": sub_app,
                    "absent": sub_abs,
                    "passed": sub_pass,
                    "failed": sub_fail,
                    "pass_rate": sub_rate,
                    "gpa5_count": sub_gpa5_w4,
                    "gpa5_count_with_4th": sub_gpa5_w4,
                    "gpa5_count_without_4th": sub_gpa5_no4,
                    "one_fail_count": sum(1 for s in s_sub if s["Failed_Count"] == 1 and s["Result"] != "ABSENT"),
                    "two_fail_count": sum(1 for s in s_sub if s["Failed_Count"] == 2 and s["Result"] != "ABSENT"),
                    "multi_fail_count": sum(1 for s in s_sub if s["Failed_Count"] >= 3 and s["Result"] != "ABSENT")
                })

    regular_summary = types_summary.get("REGULAR", {
        "type": "REGULAR", "total": 0, "appeared": 0, "absent": 0, "absent_rate": 0.0,
        "passed": 0, "failed": 0, "pass_rate": 0.0, "fail_rate": 0.0,
        "gpa5_count": 0, "gpa5_count_with_4th": 0, "gpa5_count_without_4th": 0,
        "gpa5_rate": 0.0, "gpa5_rate_with_4th": 0.0, "gpa5_rate_without_4th": 0.0,
        "one_fail_count": 0, "two_fail_count": 0, "multi_fail_count": 0, "group_breakdown": {}
    })
    irregular_summary = types_summary.get("IRREGULAR", {
        "type": "IRREGULAR", "total": 0, "appeared": 0, "absent": 0, "absent_rate": 0.0,
        "passed": 0, "failed": 0, "pass_rate": 0.0, "fail_rate": 0.0,
        "gpa5_count": 0, "gpa5_count_with_4th": 0, "gpa5_count_without_4th": 0,
        "gpa5_rate": 0.0, "gpa5_rate_with_4th": 0.0, "gpa5_rate_without_4th": 0.0,
        "one_fail_count": 0, "two_fail_count": 0, "multi_fail_count": 0, "group_breakdown": {}
    })

    # Absent Analysis aggregation
    absent_analysis = {
        "total_absent": total_absent,
        "absent_rate": absent_rate,
        "regular_absent": sum(1 for s in absent_students if s["Type"] == "REGULAR"),
        "irregular_absent": sum(1 for s in absent_students if s["Type"] == "IRREGULAR"),
        "groups_absent": {
            grp: {
                "total": len(s_list),
                "absent": sum(1 for s in s_list if s["Result"] == "ABSENT"),
                "rate": round(sum(1 for s in s_list if s["Result"] == "ABSENT") / len(s_list) * 100, 2) if len(s_list) > 0 else 0.0
            }
            for grp, s_list in group_students.items()
        },
        "types_absent": {
            t: {
                "total": len(s_list),
                "absent": sum(1 for s in s_list if s["Result"] == "ABSENT"),
                "rate": round(sum(1 for s in s_list if s["Result"] == "ABSENT") / len(s_list) * 100, 2) if len(s_list) > 0 else 0.0
            }
            for t, s_list in type_students.items()
        },
        "students": absent_students
    }

    # Group & Subject-Wise Grade Analysis table records
    subject_analysis = []
    for (grp, c_type, subj), counts in subject_stats.items():
        tot_subj_appeared = sum(counts.values())
        if tot_subj_appeared == 0:
            continue
        subj_failed = counts.get("F", 0)
        subj_passed = tot_subj_appeared - subj_failed
        subj_pass_rate = round((subj_passed / tot_subj_appeared * 100), 2)

        subject_analysis.append({
            "group": grp,
            "type": c_type,
            "subject": subj,
            "appeared": tot_subj_appeared,
            "passed": subj_passed,
            "failed": subj_failed,
            "pass_rate": subj_pass_rate,
            "grades": {g: counts.get(g, 0) for g in VALID_GRADES}
        })

    # Sort subjects: first by group, type, then subject
    subject_analysis.sort(key=lambda x: (x["group"], x["type"], x["subject"]))

    # Sort 1-fail, 2-fail, and 3+-fail subjects by frequency descending
    sorted_one_sub_freq = dict(sorted(one_sub_freq.items(), key=lambda item: item[1], reverse=True))
    sorted_two_sub_freq = dict(sorted(two_sub_freq.items(), key=lambda item: item[1], reverse=True))
    sorted_two_sub_pairs = dict(sorted(two_sub_pairs.items(), key=lambda item: item[1], reverse=True))
    sorted_three_plus_sub_freq = dict(sorted(three_plus_sub_freq.items(), key=lambda item: item[1], reverse=True))

    available_groups = sorted(list(group_students.keys()))
    available_types = sorted(list(type_students.keys()))
    available_boards = sorted(list({r["Board"] for r in normalized if r.get("Board") and r["Board"] != "N/A"}))

    # Build standardized Subject Fail Tables (Without 4th Subject) for All and Group-Wise with Regular & Irregular Breakdown
    sorted_one_sub_type_all = dict(sorted(one_sub_type_freq.items(), key=lambda item: (-item[1]["total"], item[0])))
    sorted_two_sub_type_all = dict(sorted(two_sub_type_freq.items(), key=lambda item: (-item[1]["total"], item[0])))
    sorted_three_plus_type_all = dict(sorted(three_plus_type_freq.items(), key=lambda item: (-item[1]["total"], item[0])))

    subject_fail_tables = {
        "one_subject_fail": {
            "title": "One Subject Fail (Without 4th Subject)",
            "all": sorted_one_sub_type_all,
            "by_group": {
                grp: dict(sorted(one_sub_group_type_freq[grp].items(), key=lambda item: (-item[1]["total"], item[0])))
                for grp in available_groups
            }
        },
        "two_subject_fail": {
            "title": "Two Subject Fail (Without 4th Subject)",
            "all": sorted_two_sub_type_all,
            "by_group": {
                grp: dict(sorted(two_sub_group_type_freq[grp].items(), key=lambda item: (-item[1]["total"], item[0])))
                for grp in available_groups
            }
        },
        "three_plus_subject_fail": {
            "title": "Three or More Subject Fail (Without 4th Subject)",
            "all": sorted_three_plus_type_all,
            "by_group": {
                grp: dict(sorted(three_plus_group_type_freq[grp].items(), key=lambda item: (-item[1]["total"], item[0])))
                for grp in available_groups
            }
        }
    }

    # =========================================================================
    # 1. STUDENT MERIT LIST
    # =========================================================================
    def merit_sort_key(s):
        gpa_w4 = s.get("GPA_With_4th_Numeric", 0.0)
        gpa_no4 = s.get("GPA_Without_4th_Numeric", 0.0)
        subj_grades = s.get("Subject_Grades", {})
        a_plus_cnt = sum(1 for g in subj_grades.values() if g == "A+")
        a_cnt = sum(1 for g in subj_grades.values() if g == "A")
        total_gp = sum(GRADE_POINTS.get(g, 0.0) for g in subj_grades.values())
        try:
            roll_val = int(str(s.get("Roll", 0)).strip())
        except (ValueError, TypeError):
            roll_val = 99999999
        return (-gpa_w4, -gpa_no4, -a_plus_cnt, -a_cnt, -total_gp, roll_val)

    sorted_passed = sorted(passed_students, key=merit_sort_key)

    group_rank_counters = defaultdict(int)
    merit_students = []
    for rank, st in enumerate(sorted_passed, start=1):
        grp = st["Group"]
        group_rank_counters[grp] += 1
        grp_rank = group_rank_counters[grp]

        subj_grades = st.get("Subject_Grades", {})
        a_plus_cnt = sum(1 for g in subj_grades.values() if g == "A+")
        a_cnt = sum(1 for g in subj_grades.values() if g == "A")
        a_minus_cnt = sum(1 for g in subj_grades.values() if g == "A-")
        b_cnt = sum(1 for g in subj_grades.values() if g == "B")
        c_cnt = sum(1 for g in subj_grades.values() if g == "C")
        d_cnt = sum(1 for g in subj_grades.values() if g == "D")

        merit_students.append({
            "Merit_Position": rank,
            "Group_Merit_Position": grp_rank,
            "Roll": st["Roll"],
            "Registration": st.get("Registration", ""),
            "Name": st["Name"],
            "Father Name": st.get("Father Name", ""),
            "Institute": st.get("Institute", ""),
            "Board": st.get("Board", "N/A"),
            "Group": grp,
            "Type": st["Type"],
            "GPA": st["GPA_With_4th"],
            "GPA_With_4th": st["GPA_With_4th"],
            "GPA_Without_4th": st["GPA_Without_4th"],
            "GPA_With_4th_Numeric": st["GPA_With_4th_Numeric"],
            "GPA_Without_4th_Numeric": st["GPA_Without_4th_Numeric"],
            "Is_GPA5_With_4th": st["Is_GPA5_With_4th"],
            "Is_GPA5_Without_4th": st["Is_GPA5_Without_4th"],
            "A_Plus_Count": a_plus_cnt,
            "A_Count": a_cnt,
            "A_Minus_Count": a_minus_cnt,
            "B_Count": b_cnt,
            "C_Count": c_cnt,
            "D_Count": d_cnt,
            "Subject_Grades": subj_grades,
            "Passed_Grades_Summary": st.get("Passed_Grades_Summary", "")
        })

    # =========================================================================
    # 2. FAILED LIST SUBJECT WISE (ASCENDING ORDER)
    # =========================================================================
    subj_fail_map = defaultdict(list)
    for st in failed_students:
        for subj in st.get("Failed_Subjects", []):
            subj_fail_map[subj].append(st)

    # Sort subject names in Ascending order (A -> Z / Subject Code)
    sorted_fail_subjs_asc = sorted(subj_fail_map.keys())

    failed_subjects_summary = []
    flat_failed_entries = []

    def roll_sort_key(s):
        try:
            return int(str(s.get("Roll", 0)).strip())
        except (ValueError, TypeError):
            return str(s.get("Roll", ""))

    for subj in sorted_fail_subjs_asc:
        s_list = subj_fail_map[subj]
        s_list_sorted = sorted(s_list, key=roll_sort_key)

        failed_subjects_summary.append({
            "subject": subj,
            "count": len(s_list_sorted),
            "regular_count": sum(1 for s in s_list_sorted if s.get("Type") == "REGULAR"),
            "irregular_count": sum(1 for s in s_list_sorted if s.get("Type") == "IRREGULAR"),
            "students": [
                {
                    "Roll": s["Roll"],
                    "Registration": s.get("Registration", ""),
                    "Name": s["Name"],
                    "Group": s["Group"],
                    "Type": s["Type"],
                    "Failed_Count": s["Failed_Count"],
                    "All_Failed_Subjects": s.get("Failed_Subjects", []),
                    "Passed_Grades_Summary": s.get("Passed_Grades_Summary", "")
                }
                for s in s_list_sorted
            ]
        })

        for s in s_list_sorted:
            flat_failed_entries.append({
                "Failed_Subject": subj,
                "Roll": s["Roll"],
                "Registration": s.get("Registration", ""),
                "Name": s["Name"],
                "Father Name": s.get("Father Name", ""),
                "Institute": s.get("Institute", ""),
                "Board": s.get("Board", "N/A"),
                "Group": s["Group"],
                "Type": s["Type"],
                "Failed_Count": s["Failed_Count"],
                "All_Failed_Subjects": s.get("Failed_Subjects", []),
                "Passed_Grades_Summary": s.get("Passed_Grades_Summary", "")
            })

    all_institutes = [
        str(r.get("Institute", "")).strip()
        for r in normalized
        if r.get("Institute") and str(r.get("Institute")).strip().lower() not in ["", "nan", "none", "-", "n/a"]
    ]
    from collections import Counter
    top_institute = Counter(all_institutes).most_common(1)[0][0] if all_institutes else ""

    all_boards = [
        str(r.get("Board", "")).strip()
        for r in normalized
        if r.get("Board") and str(r.get("Board")).strip().lower() not in ["", "nan", "none", "-", "n/a"]
    ]
    top_board = Counter(all_boards).most_common(1)[0][0] if all_boards else ""

    all_eiins = [
        str(r.get("EIIN", "")).strip()
        for r in normalized
        if r.get("EIIN") and str(r.get("EIIN")).strip().lower() not in ["", "nan", "none", "-", "n/a"]
    ]
    top_eiin = (
        Counter(all_eiins).most_common(1)[0][0]
        if all_eiins
        else resolve_institution_eiin("", top_institute, top_board)
    )
    inst_details = resolve_institution_details(raw_eiin=top_eiin, institute=top_institute, board=top_board)
    top_eiin = inst_details.get("eiin", "") or top_eiin
    top_address = inst_details.get("address", "")

    fail_subject_summary_matrix = build_fail_subject_summary_matrix(normalized)

    return {
        "total_students": total_students,
        "institute": top_institute,
        "eiin": top_eiin,
        "address": top_address,
        "board": top_board,
        "summary": {
            "institute": top_institute,
            "eiin": top_eiin,
            "address": top_address,
            "board": top_board,
            "total_students": total_students,
            "appeared": appeared,
            "passed": total_passed,
            "failed": total_failed,
            "absent": total_absent,
            "absent_rate": absent_rate,
            "pass_rate": pass_rate,
            "fail_rate": fail_rate,
            "gpa_5_count": gpa_5_count_with_4th,
            "gpa_5_rate": gpa_5_rate_with_4th,
            "gpa_5_count_with_4th": gpa_5_count_with_4th,
            "gpa_5_rate_with_4th": gpa_5_rate_with_4th,
            "gpa_5_count_without_4th": gpa_5_count_without_4th,
            "gpa_5_rate_without_4th": gpa_5_rate_without_4th,
            "one_subject_fail_count": len(one_subject_fail_students),
            "two_subject_fail_count": len(two_subject_fail_students),
            "multi_subject_fail_count": len(multi_subject_fail_students)
        },
        "groups_summary": groups_summary,
        "types_summary": types_summary,
        "regular_summary": regular_summary,
        "irregular_summary": irregular_summary,
        "group_type_comparison": group_type_comparison,
        "subject_analysis": subject_analysis,
        "merit_list": {
            "total": len(merit_students),
            "students": merit_students
        },
        "failed_list_subject_wise": {
            "total_failed_students": len(failed_students),
            "total_subject_failures": len(flat_failed_entries),
            "subjects": failed_subjects_summary,
            "entries": flat_failed_entries
        },
        "subject_fail_tables": subject_fail_tables,
        "fail_subject_summary_matrix": fail_subject_summary_matrix,
        "one_subject_fail": {
            "total": len(one_subject_fail_students),
            "subjects_summary": sorted_one_sub_freq,
            "by_group": subject_fail_tables["one_subject_fail"]["by_group"],
            "students": sorted(one_subject_fail_students, key=lambda s: (s["Failed_Subject"], s["Roll"]))
        },
        "two_subject_fail": {
            "total": len(two_subject_fail_students),
            "subjects_summary": sorted_two_sub_freq,
            "pair_summary": sorted_two_sub_pairs,
            "by_group": subject_fail_tables["two_subject_fail"]["by_group"],
            "students": sorted(two_subject_fail_students, key=lambda s: (s["Failed_Subjects"][0] if s["Failed_Subjects"] else "", s["Roll"]))
        },
        "multi_subject_fail": {
            "total": len(multi_subject_fail_students),
            "subjects_summary": sorted_three_plus_sub_freq,
            "by_group": subject_fail_tables["three_plus_subject_fail"]["by_group"],
            "students": sorted(multi_subject_fail_students, key=lambda s: (-s["Failed_Count"], s["Roll"]))
        },
        "gpa_distribution": gpa_dist_with_4th,
        "gpa_distribution_with_4th": gpa_dist_with_4th,
        "gpa_distribution_without_4th": gpa_dist_without_4th,
        "absent_analysis": absent_analysis,
        "available_groups": available_groups,
        "available_types": available_types,
        "available_boards": available_boards
    }


def generate_analysis_excel_bytes(analysis: Dict[str, Any], meta: Optional[Dict[str, str]] = None) -> io.BytesIO:
    """
    Generates an Excel workbook (.xlsx) containing formatted analysis worksheets:
    1. Executive Summary & Overview
    2. Regular vs Irregular Analysis
    3. Group & Subject Grade Analysis
    4. 1-Subject Fail Breakdown
    5. 2-Subject Fail Breakdown
    6. 3+ Subject Fail List
    """
    wb = openpyxl.Workbook()
    wb.remove(wb.active)  # Remove default sheet

    meta = meta or {}
    exam = meta.get("exam", "HSC / SSC").upper()
    year = meta.get("year", "")
    board = meta.get("board", "").title()
    eiin = meta.get("eiin") or analysis.get("eiin") or (analysis.get("summary", {}).get("eiin", ""))
    eiin_str = f" | EIIN: {eiin}" if eiin else ""
    title_sub = f"Examination: {exam} {year} | Board: {board}{eiin_str}"

    # Palette Styles
    header_fill = PatternFill(start_color="0F172A", end_color="0F172A", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")

    card_header_fill = PatternFill(start_color="E2E8F0", end_color="E2E8F0", fill_type="solid")
    bold_font = Font(name="Calibri", size=11, bold=True)
    normal_font = Font(name="Calibri", size=11)

    fail_badge_fill = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid")
    fail_font = Font(name="Calibri", size=11, bold=True, color="991B1B")

    pass_badge_fill = PatternFill(start_color="DCFCE7", end_color="DCFCE7", fill_type="solid")
    pass_font = Font(name="Calibri", size=11, bold=True, color="166534")

    thin_border = Border(
        left=Side(style="thin", color="CBD5E1"),
        right=Side(style="thin", color="CBD5E1"),
        top=Side(style="thin", color="CBD5E1"),
        bottom=Side(style="thin", color="CBD5E1")
    )

    align_center = Alignment(horizontal="center", vertical="center")
    align_left = Alignment(horizontal="left", vertical="center")
    align_right = Alignment(horizontal="right", vertical="center")

    summary = analysis.get("summary", {})

    # =========================================================================
    # SHEET 1: EXECUTIVE SUMMARY
    # =========================================================================
    ws_sum = wb.create_sheet(title="Executive Summary")
    ws_sum.views.sheetView[0].showGridLines = True

    # Title Banner
    ws_sum.merge_cells("A1:H1")
    ws_sum["A1"] = "BOARD RESULT COMPREHENSIVE ANALYSIS REPORT"
    ws_sum["A1"].font = Font(name="Calibri", size=16, bold=True, color="FFFFFF")
    ws_sum["A1"].fill = PatternFill(start_color="0F172A", end_color="0F172A", fill_type="solid")
    ws_sum["A1"].alignment = align_center
    ws_sum.row_dimensions[1].height = 36

    ws_sum.merge_cells("A2:H2")
    ws_sum["A2"] = title_sub
    ws_sum["A2"].font = Font(name="Calibri", size=11, italic=True, color="FFFFFF")
    ws_sum["A2"].fill = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
    ws_sum["A2"].alignment = align_center
    ws_sum.row_dimensions[2].height = 22

    # KPI Summary Table
    ws_sum.cell(row=4, column=1, value="OVERALL RESULT METRICS").font = Font(name="Calibri", size=12, bold=True, color="0F172A")
    kpis = [
        ("Total Students Processed", summary.get("total_students", 0)),
        ("Students Appeared", summary.get("appeared", 0)),
        ("Students Absent", f"{summary.get('absent', 0)} ({summary.get('absent_rate', 0)}%)"),
        ("Total Passed", f"{summary.get('passed', 0)} ({summary.get('pass_rate', 0)}%)"),
        ("Total Failed", f"{summary.get('failed', 0)} ({summary.get('fail_rate', 0)}%)"),
        ("GPA 5.00 With 4th Subject", f"{summary.get('gpa_5_count_with_4th', 0)} ({summary.get('gpa_5_rate_with_4th', 0)}%)"),
        ("GPA 5.00 Without 4th Sub (Golden A+)", f"{summary.get('gpa_5_count_without_4th', 0)} ({summary.get('gpa_5_rate_without_4th', 0)}%)"),
        ("One-Subject Fail (1 Sub)", summary.get("one_subject_fail_count", 0)),
        ("Two-Subject Fail (2 Subs)", summary.get("two_subject_fail_count", 0)),
        ("Three or More Subjects Fail (3+)", summary.get("multi_subject_fail_count", 0)),
    ]

    r_idx = 5
    for label, val in kpis:
        c1 = ws_sum.cell(row=r_idx, column=1, value=label)
        c2 = ws_sum.cell(row=r_idx, column=2, value=val)
        c1.font = bold_font
        c2.font = normal_font
        c1.border = thin_border
        c2.border = thin_border
        c2.alignment = align_center
        r_idx += 1

    # GPA Distribution (With 4th Subject)
    ws_sum.cell(row=4, column=4, value="GPA (WITH 4TH SUB)").font = Font(name="Calibri", size=12, bold=True, color="0F172A")
    r_idx_gpa1 = 5
    for tier, count in analysis.get("gpa_distribution_with_4th", analysis.get("gpa_distribution", {})).items():
        c1 = ws_sum.cell(row=r_idx_gpa1, column=4, value=tier)
        c2 = ws_sum.cell(row=r_idx_gpa1, column=5, value=count)
        c1.font = bold_font
        c2.font = normal_font
        c1.border = thin_border
        c2.border = thin_border
        c2.alignment = align_center
        r_idx_gpa1 += 1

    # GPA Distribution (Without 4th Subject)
    ws_sum.cell(row=4, column=7, value="GPA (WITHOUT 4TH SUB)").font = Font(name="Calibri", size=12, bold=True, color="0F172A")
    r_idx_gpa2 = 5
    for tier, count in analysis.get("gpa_distribution_without_4th", {}).items():
        c1 = ws_sum.cell(row=r_idx_gpa2, column=7, value=tier)
        c2 = ws_sum.cell(row=r_idx_gpa2, column=8, value=count)
        c1.font = bold_font
        c2.font = normal_font
        c1.border = thin_border
        c2.border = thin_border
        c2.alignment = align_center
        r_idx_gpa2 += 1

    # Regular vs Irregular Overview in Executive Summary
    type_start_row = max(r_idx, r_idx_gpa1, r_idx_gpa2) + 2
    ws_sum.cell(row=type_start_row, column=1, value="REGULAR VS IRREGULAR CANDIDATE PERFORMANCE").font = Font(name="Calibri", size=12, bold=True, color="0F172A")

    type_headers = ["Candidate Type", "Total", "Appeared", "Absent", "Passed", "Failed", "Pass %", "GPA 5 (With 4th)", "GPA 5 (Without 4th)", "1 Sub Fail", "2 Sub Fail", "3+ Sub Fail"]
    h_row_type = type_start_row + 1
    for col_idx, h_text in enumerate(type_headers, start=1):
        c = ws_sum.cell(row=h_row_type, column=col_idx, value=h_text)
        c.fill = header_fill
        c.font = header_font
        c.alignment = align_center
        c.border = thin_border

    curr_r = h_row_type + 1
    for t_data in analysis.get("types_summary", {}).values():
        row_vals = [
            t_data["type"],
            t_data["total"],
            t_data["appeared"],
            t_data.get("absent", 0),
            t_data["passed"],
            t_data["failed"],
            f"{t_data['pass_rate']}%",
            t_data.get("gpa5_count_with_4th", t_data.get("gpa5_count", 0)),
            t_data.get("gpa5_count_without_4th", 0),
            t_data["one_fail_count"],
            t_data["two_fail_count"],
            t_data["multi_fail_count"],
        ]
        for col_idx, val in enumerate(row_vals, start=1):
            c = ws_sum.cell(row=curr_r, column=col_idx, value=val)
            c.font = bold_font if col_idx == 1 else normal_font
            c.border = thin_border
            c.alignment = align_left if col_idx == 1 else align_center
            if col_idx == 7:
                if t_data["pass_rate"] >= 80:
                    c.fill = pass_badge_fill
                    c.font = pass_font
                elif t_data["pass_rate"] < 50:
                    c.fill = fail_badge_fill
                    c.font = fail_font
        curr_r += 1

    # Group Performance Comparison Table
    group_start_row = curr_r + 2
    ws_sum.cell(row=group_start_row, column=1, value="GROUP-WISE PERFORMANCE COMPARISON").font = Font(name="Calibri", size=12, bold=True, color="0F172A")

    grp_headers = ["Group", "Total", "Appeared", "Absent", "Passed", "Failed", "Pass %", "GPA 5 (With 4th)", "GPA 5 (Without 4th)", "1 Sub Fail", "2 Sub Fail", "3+ Sub Fail"]
    h_row = group_start_row + 1
    for col_idx, h_text in enumerate(grp_headers, start=1):
        c = ws_sum.cell(row=h_row, column=col_idx, value=h_text)
        c.fill = header_fill
        c.font = header_font
        c.alignment = align_center
        c.border = thin_border

    curr_r = h_row + 1
    for grp_data in analysis.get("groups_summary", {}).values():
        row_vals = [
            grp_data["group"],
            grp_data["total"],
            grp_data["appeared"],
            grp_data.get("absent", 0),
            grp_data["passed"],
            grp_data["failed"],
            f"{grp_data['pass_rate']}%",
            grp_data.get("gpa5_count_with_4th", grp_data.get("gpa5_count", 0)),
            grp_data.get("gpa5_count_without_4th", 0),
            grp_data["one_fail_count"],
            grp_data["two_fail_count"],
            grp_data["multi_fail_count"],
        ]
        for col_idx, val in enumerate(row_vals, start=1):
            c = ws_sum.cell(row=curr_r, column=col_idx, value=val)
            c.font = bold_font if col_idx == 1 else normal_font
            c.border = thin_border
            c.alignment = align_left if col_idx == 1 else align_center
            if col_idx == 7:
                if grp_data["pass_rate"] >= 80:
                    c.fill = pass_badge_fill
                    c.font = pass_font
                elif grp_data["pass_rate"] < 50:
                    c.fill = fail_badge_fill
                    c.font = fail_font
        curr_r += 1
        curr_r += 1

    for col in ws_sum.columns:
        max_len = max(len(str(cell.value or "")) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws_sum.column_dimensions[col_letter].width = max(max_len + 4, 14)

    # =========================================================================
    # SHEET 2: REGULAR VS IRREGULAR ANALYSIS
    # =========================================================================
    ws_type_sheet = wb.create_sheet(title="Regular vs Irregular")
    ws_type_sheet.views.sheetView[0].showGridLines = True

    # Title
    ws_type_sheet.merge_cells("A1:L1")
    ws_type_sheet["A1"] = "REGULAR VS IRREGULAR CANDIDATE RESULT ANALYSIS"
    ws_type_sheet["A1"].font = Font(name="Calibri", size=14, bold=True, color="FFFFFF")
    ws_type_sheet["A1"].fill = header_fill
    ws_type_sheet["A1"].alignment = align_center
    ws_type_sheet.row_dimensions[1].height = 30

    # Section 1: Candidate Type Overall Comparison
    ws_type_sheet.cell(row=3, column=1, value="CANDIDATE TYPE OVERALL PERFORMANCE").font = Font(name="Calibri", size=12, bold=True, color="0F172A")
    type_cols = ["Candidate Type", "Total Students", "Appeared", "Absent", "Passed", "Failed", "Pass Rate (%)", "GPA 5 (With 4th)", "GPA 5 (Without 4th)", "1-Sub Fail", "2-Sub Fail", "3+ Sub Fail"]
    for col_idx, h_text in enumerate(type_cols, start=1):
        c = ws_type_sheet.cell(row=4, column=col_idx, value=h_text)
        c.fill = header_fill
        c.font = header_font
        c.alignment = align_center
        c.border = thin_border

    curr_r = 5
    for t_data in analysis.get("types_summary", {}).values():
        row_vals = [
            t_data["type"],
            t_data["total"],
            t_data["appeared"],
            t_data.get("absent", 0),
            t_data["passed"],
            t_data["failed"],
            f"{t_data['pass_rate']}%",
            t_data.get("gpa5_count_with_4th", t_data.get("gpa5_count", 0)),
            t_data.get("gpa5_count_without_4th", 0),
            t_data["one_fail_count"],
            t_data["two_fail_count"],
            t_data["multi_fail_count"],
        ]
        for col_idx, val in enumerate(row_vals, start=1):
            c = ws_type_sheet.cell(row=curr_r, column=col_idx, value=val)
            c.font = bold_font if col_idx == 1 else normal_font
            c.border = thin_border
            c.alignment = align_left if col_idx == 1 else align_center
            if col_idx == 7:
                if t_data["pass_rate"] >= 80:
                    c.fill = pass_badge_fill
                    c.font = pass_font
                elif t_data["pass_rate"] < 50:
                    c.fill = fail_badge_fill
                    c.font = fail_font
        curr_r += 1

    # Section 2: Group x Candidate Type Matrix
    curr_r += 2
    ws_type_sheet.cell(row=curr_r, column=1, value="GROUP-WISE REGULAR VS IRREGULAR BREAKDOWN").font = Font(name="Calibri", size=12, bold=True, color="0F172A")
    curr_r += 1

    group_type_cols = ["Group Name", "Candidate Type", "Total Students", "Appeared", "Absent", "Passed", "Failed", "Pass Rate (%)", "GPA 5 (With 4th)", "GPA 5 (Without 4th)", "1-Sub Fail", "2-Sub Fail", "3+ Sub Fail"]
    for col_idx, h_text in enumerate(group_type_cols, start=1):
        c = ws_type_sheet.cell(row=curr_r, column=col_idx, value=h_text)
        c.fill = header_fill
        c.font = header_font
        c.alignment = align_center
        c.border = thin_border

    curr_r += 1
    for row_item in analysis.get("group_type_comparison", []):
        row_vals = [
            row_item["group"],
            row_item["type"],
            row_item["total"],
            row_item["appeared"],
            row_item.get("absent", 0),
            row_item["passed"],
            row_item["failed"],
            f"{row_item['pass_rate']}%",
            row_item.get("gpa5_count_with_4th", row_item.get("gpa5_count", 0)),
            row_item.get("gpa5_count_without_4th", 0),
            row_item["one_fail_count"],
            row_item["two_fail_count"],
            row_item["multi_fail_count"],
        ]
        for col_idx, val in enumerate(row_vals, start=1):
            c = ws_type_sheet.cell(row=curr_r, column=col_idx, value=val)
            c.font = bold_font if col_idx in [1, 2] else normal_font
            c.border = thin_border
            c.alignment = align_left if col_idx in [1, 2] else align_center
            if col_idx == 8:
                if row_item["pass_rate"] >= 80:
                    c.fill = pass_badge_fill
                    c.font = pass_font
                elif row_item["pass_rate"] < 50:
                    c.fill = fail_badge_fill
                    c.font = fail_font
        curr_r += 1

    for col in ws_type_sheet.columns:
        max_len = max(len(str(cell.value or "")) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws_type_sheet.column_dimensions[col_letter].width = max(max_len + 3, 13)

    # =========================================================================
    # SHEET 3: GROUP & SUBJECT-WISE GRADE ANALYSIS
    # =========================================================================
    ws_subj = wb.create_sheet(title="Subject Wise Grade Analysis")
    ws_subj.views.sheetView[0].showGridLines = True

    subj_headers = [
        "Group", "Candidate Type", "Subject Code & Name", "Appeared", "Passed", "Failed", "Pass Rate (%)",
        "A+", "A", "A-", "B", "C", "D", "F"
    ]
    for col_idx, h_text in enumerate(subj_headers, start=1):
        c = ws_subj.cell(row=1, column=col_idx, value=h_text)
        c.fill = header_fill
        c.font = header_font
        c.alignment = align_center
        c.border = thin_border
    ws_subj.row_dimensions[1].height = 28

    curr_r = 2
    for item in analysis.get("subject_analysis", []):
        grades = item["grades"]
        row_vals = [
            item["group"],
            item.get("type", "ALL"),
            item["subject"],
            item["appeared"],
            item["passed"],
            item["failed"],
            f"{item['pass_rate']}%",
            grades.get("A+", 0),
            grades.get("A", 0),
            grades.get("A-", 0),
            grades.get("B", 0),
            grades.get("C", 0),
            grades.get("D", 0),
            grades.get("F", 0)
        ]
        for col_idx, val in enumerate(row_vals, start=1):
            c = ws_subj.cell(row=curr_r, column=col_idx, value=val)
            c.font = normal_font
            c.border = thin_border
            c.alignment = align_left if col_idx in [1, 2, 3] else align_center
            if col_idx == 7:  # Pass rate
                if item["pass_rate"] >= 80:
                    c.fill = pass_badge_fill
                    c.font = pass_font
                elif item["pass_rate"] < 50:
                    c.fill = fail_badge_fill
                    c.font = fail_font
            elif col_idx == 14 and grades.get("F", 0) > 0:  # F count
                c.fill = fail_badge_fill
                c.font = fail_font

        curr_r += 1

    for col in ws_subj.columns:
        max_len = max(len(str(cell.value or "")) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws_subj.column_dimensions[col_letter].width = max(max_len + 3, 11)

    # =========================================================================
    # SHEET 4: ONE SUBJECT FAIL BREAKDOWN
    # =========================================================================
    ws_1fail = wb.create_sheet(title="1 Subject Fail")
    ws_1fail.views.sheetView[0].showGridLines = True

    one_fail_data = analysis.get("one_subject_fail", {})

    # Top summary of 1-fail by subject
    ws_1fail.cell(row=1, column=1, value="ONE-SUBJECT FAIL SUMMARY BY SUBJECT").font = Font(name="Calibri", size=12, bold=True, color="0F172A")
    ws_1fail.cell(row=2, column=1, value="Failed Subject").fill = header_fill
    ws_1fail.cell(row=2, column=1).font = header_font
    ws_1fail.cell(row=2, column=2, value="Student Count").fill = header_fill
    ws_1fail.cell(row=2, column=2).font = header_font

    r_sum = 3
    for subj_name, cnt in one_fail_data.get("subjects_summary", {}).items():
        c1 = ws_1fail.cell(row=r_sum, column=1, value=subj_name)
        c2 = ws_1fail.cell(row=r_sum, column=2, value=cnt)
        c1.font = bold_font
        c2.font = fail_font
        c1.border = thin_border
        c2.border = thin_border
        c2.alignment = align_center
        r_sum += 1

    # Detailed Student Roster
    start_roster = r_sum + 2
    ws_1fail.cell(row=start_roster, column=1, value="STUDENT ROSTER (1-SUBJECT FAIL)").font = Font(name="Calibri", size=12, bold=True, color="0F172A")

    roster_headers = ["Roll No", "Reg No", "Student Name", "Group", "Candidate Type", "GPA (With 4th)", "GPA (Without 4th)", "Failed Subject (F)", "Passed Subject Grades", "Remarks"]
    for col_idx, h_text in enumerate(roster_headers, start=1):
        c = ws_1fail.cell(row=start_roster + 1, column=col_idx, value=h_text)
        c.fill = header_fill
        c.font = header_font
        c.alignment = align_center
        c.border = thin_border

    curr_r = start_roster + 2
    for st in one_fail_data.get("students", []):
        row_vals = [
            st["Roll"],
            st["Registration"],
            st["Name"],
            st["Group"],
            st.get("Type", "REGULAR"),
            st.get("GPA_With_4th", st.get("GPA", "0.00")),
            st.get("GPA_Without_4th", "0.00"),
            st["Failed_Subject"],
            st["Passed_Grades_Summary"],
            "Eligible for 1-Subject Improvement"
        ]
        for col_idx, val in enumerate(row_vals, start=1):
            c = ws_1fail.cell(row=curr_r, column=col_idx, value=val)
            c.font = normal_font
            c.border = thin_border
            c.alignment = align_center if col_idx in [1, 2, 4, 5, 6, 7, 10] else align_left
            if col_idx == 8:
                c.fill = fail_badge_fill
                c.font = fail_font
        curr_r += 1

    for col in ws_1fail.columns:
        max_len = max(len(str(cell.value or "")) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws_1fail.column_dimensions[col_letter].width = min(max(max_len + 3, 12), 45)

    # =========================================================================
    # SHEET 5: TWO SUBJECT FAIL BREAKDOWN
    # =========================================================================
    ws_2fail = wb.create_sheet(title="2 Subject Fail")
    ws_2fail.views.sheetView[0].showGridLines = True

    two_fail_data = analysis.get("two_subject_fail", {})

    ws_2fail.cell(row=1, column=1, value="TWO-SUBJECT FAIL SUMMARY (TOP FAILED SUBJECTS)").font = Font(name="Calibri", size=12, bold=True, color="0F172A")
    ws_2fail.cell(row=2, column=1, value="Subject").fill = header_fill
    ws_2fail.cell(row=2, column=1).font = header_font
    ws_2fail.cell(row=2, column=2, value="Occurrences in 2-Fail").fill = header_fill
    ws_2fail.cell(row=2, column=2).font = header_font

    r_sum = 3
    for subj_name, cnt in two_fail_data.get("subjects_summary", {}).items():
        c1 = ws_2fail.cell(row=r_sum, column=1, value=subj_name)
        c2 = ws_2fail.cell(row=r_sum, column=2, value=cnt)
        c1.font = bold_font
        c2.font = fail_font
        c1.border = thin_border
        c2.border = thin_border
        c2.alignment = align_center
        r_sum += 1

    # Roster
    start_roster = r_sum + 2
    ws_2fail.cell(row=start_roster, column=1, value="STUDENT ROSTER (2-SUBJECT FAIL)").font = Font(name="Calibri", size=12, bold=True, color="0F172A")

    roster2_headers = ["Roll No", "Reg No", "Student Name", "Group", "Candidate Type", "GPA (With 4th)", "GPA (Without 4th)", "Failed Subject 1 (F)", "Failed Subject 2 (F)", "Passed Subject Grades"]
    for col_idx, h_text in enumerate(roster2_headers, start=1):
        c = ws_2fail.cell(row=start_roster + 1, column=col_idx, value=h_text)
        c.fill = header_fill
        c.font = header_font
        c.alignment = align_center
        c.border = thin_border

    curr_r = start_roster + 2
    for st in two_fail_data.get("students", []):
        f_list = st.get("Failed_Subjects", [])
        f1 = f_list[0] if len(f_list) > 0 else "-"
        f2 = f_list[1] if len(f_list) > 1 else "-"

        row_vals = [
            st["Roll"],
            st["Registration"],
            st["Name"],
            st["Group"],
            st.get("Type", "REGULAR"),
            st.get("GPA_With_4th", st.get("GPA", "0.00")),
            st.get("GPA_Without_4th", "0.00"),
            f1,
            f2,
            st["Passed_Grades_Summary"]
        ]
        for col_idx, val in enumerate(row_vals, start=1):
            c = ws_2fail.cell(row=curr_r, column=col_idx, value=val)
            c.font = normal_font
            c.border = thin_border
            c.alignment = align_center if col_idx in [1, 2, 4, 5, 6, 7] else align_left
            if col_idx in [8, 9]:
                c.fill = fail_badge_fill
                c.font = fail_font
        curr_r += 1

    for col in ws_2fail.columns:
        max_len = max(len(str(cell.value or "")) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws_2fail.column_dimensions[col_letter].width = min(max(max_len + 3, 12), 45)

    # =========================================================================
    # SHEET 6: 3+ SUBJECT FAIL (MULTI-FAIL)
    # =========================================================================
    ws_mfail = wb.create_sheet(title="3+ Subject Fail")
    ws_mfail.views.sheetView[0].showGridLines = True

    multi_fail_data = analysis.get("multi_subject_fail", {})
    ws_mfail.cell(row=1, column=1, value="STUDENTS FAILING 3 OR MORE SUBJECTS").font = Font(name="Calibri", size=12, bold=True, color="0F172A")

    roster3_headers = ["Roll No", "Student Name", "Group", "Candidate Type", "GPA (With 4th)", "GPA (Without 4th)", "Number of Failed Subjects", "Failed Subjects List", "Passed Grades Summary"]
    for col_idx, h_text in enumerate(roster3_headers, start=1):
        c = ws_mfail.cell(row=2, column=col_idx, value=h_text)
        c.fill = header_fill
        c.font = header_font
        c.alignment = align_center
        c.border = thin_border

    curr_r = 3
    for st in multi_fail_data.get("students", []):
        f_list = st.get("Failed_Subjects", [])
        row_vals = [
            st["Roll"],
            st["Name"],
            st["Group"],
            st.get("Type", "REGULAR"),
            st.get("GPA_With_4th", st.get("GPA", "0.00")),
            st.get("GPA_Without_4th", "0.00"),
            st["Failed_Count"],
            ", ".join(f_list),
            st["Passed_Grades_Summary"]
        ]
        for col_idx, val in enumerate(row_vals, start=1):
            c = ws_mfail.cell(row=curr_r, column=col_idx, value=val)
            c.font = normal_font
            c.border = thin_border
            c.alignment = align_center if col_idx in [1, 3, 4, 5, 6, 7] else align_left
            if col_idx == 7:
                c.fill = fail_badge_fill
                c.font = fail_font
        curr_r += 1

    for col in ws_mfail.columns:
        max_len = max(len(str(cell.value or "")) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws_mfail.column_dimensions[col_letter].width = min(max(max_len + 3, 12), 45)

    # =========================================================================
    # SHEET 7: ABSENT STUDENTS ANALYSIS & ROSTER
    # =========================================================================
    ws_absent = wb.create_sheet(title="Absent Students")
    ws_absent.views.sheetView[0].showGridLines = True

    absent_data = analysis.get("absent_analysis", {})

    # Title
    ws_absent.merge_cells("A1:I1")
    ws_absent["A1"] = "ABSENT STUDENTS ANALYSIS & ROSTER"
    ws_absent["A1"].font = Font(name="Calibri", size=14, bold=True, color="FFFFFF")
    ws_absent["A1"].fill = header_fill
    ws_absent["A1"].alignment = align_center
    ws_absent.row_dimensions[1].height = 30

    # Section 1: Summary Statistics
    ws_absent.cell(row=3, column=1, value="ABSENT SUMMARY METRICS").font = Font(name="Calibri", size=12, bold=True, color="0F172A")
    absent_kpis = [
        ("Total Registered Students", summary.get("total_students", 0)),
        ("Students Appeared", summary.get("appeared", 0)),
        ("Total Absent Candidates", absent_data.get("total_absent", summary.get("absent", 0))),
        ("Absent Rate (%)", f"{absent_data.get('absent_rate', summary.get('absent_rate', 0))}%"),
        ("Regular Candidates Absent", absent_data.get("regular_absent", 0)),
        ("Irregular Candidates Absent", absent_data.get("irregular_absent", 0)),
    ]
    r_abs = 4
    for label, val in absent_kpis:
        c1 = ws_absent.cell(row=r_abs, column=1, value=label)
        c2 = ws_absent.cell(row=r_abs, column=2, value=val)
        c1.font = bold_font
        c2.font = normal_font
        c1.border = thin_border
        c2.border = thin_border
        c2.alignment = align_center
        r_abs += 1

    # Section 2: Group-wise Absent Breakdown
    r_grp_abs_title = r_abs + 1
    ws_absent.cell(row=r_grp_abs_title, column=1, value="GROUP-WISE ABSENT BREAKDOWN").font = Font(name="Calibri", size=12, bold=True, color="0F172A")
    grp_abs_headers = ["Group Name", "Total Registered", "Absent Count", "Absent Rate (%)"]
    for col_idx, h_text in enumerate(grp_abs_headers, start=1):
        c = ws_absent.cell(row=r_grp_abs_title + 1, column=col_idx, value=h_text)
        c.fill = header_fill
        c.font = header_font
        c.alignment = align_center
        c.border = thin_border

    curr_r_grp_abs = r_grp_abs_title + 2
    for grp_name, g_info in absent_data.get("groups_absent", {}).items():
        row_vals = [
            grp_name,
            g_info.get("total", 0),
            g_info.get("absent", 0),
            f"{g_info.get('rate', 0)}%"
        ]
        for col_idx, val in enumerate(row_vals, start=1):
            c = ws_absent.cell(row=curr_r_grp_abs, column=col_idx, value=val)
            c.font = bold_font if col_idx == 1 else normal_font
            c.border = thin_border
            c.alignment = align_left if col_idx == 1 else align_center
        curr_r_grp_abs += 1

    # Section 3: Detailed Absent Student Roster
    start_roster_abs = curr_r_grp_abs + 2
    ws_absent.cell(row=start_roster_abs, column=1, value="ABSENT STUDENT ROSTER").font = Font(name="Calibri", size=12, bold=True, color="0F172A")

    roster_abs_headers = ["Roll No", "Reg No", "Student Name", "Father's Name", "Group", "Candidate Type", "Institute", "Status", "Remarks"]
    for col_idx, h_text in enumerate(roster_abs_headers, start=1):
        c = ws_absent.cell(row=start_roster_abs + 1, column=col_idx, value=h_text)
        c.fill = header_fill
        c.font = header_font
        c.alignment = align_center
        c.border = thin_border

    curr_r = start_roster_abs + 2
    abs_students = absent_data.get("students", [])
    if not abs_students:
        c = ws_absent.cell(row=curr_r, column=1, value="No absent students recorded.")
        c.font = normal_font
        ws_absent.merge_cells(start_row=curr_r, start_column=1, end_row=curr_r, end_column=9)
    else:
        for st in abs_students:
            row_vals = [
                st.get("Roll", "-"),
                st.get("Registration", "-"),
                st.get("Name", "N/A"),
                st.get("Father Name", "-"),
                st.get("Group", "GENERAL"),
                st.get("Type", "REGULAR"),
                st.get("Institute", "-"),
                "ABSENT",
                st.get("Remarks", "Did not appear")
            ]
            for col_idx, val in enumerate(row_vals, start=1):
                c = ws_absent.cell(row=curr_r, column=col_idx, value=val)
                c.font = normal_font
                c.border = thin_border
                c.alignment = align_center if col_idx in [1, 2, 5, 6, 8] else align_left
                if col_idx == 8:
                    c.fill = fail_badge_fill
                    c.font = fail_font
            curr_r += 1

    for col in ws_absent.columns:
        max_len = max(len(str(cell.value or "")) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws_absent.column_dimensions[col_letter].width = min(max(max_len + 3, 12), 40)

    # =========================================================================
    # SHEET 8: STUDENT MERIT LIST
    # =========================================================================
    ws_merit = wb.create_sheet(title="Student Merit List")
    ws_merit.views.sheetView[0].showGridLines = True

    merit_data = analysis.get("merit_list", {})
    merit_students = merit_data.get("students", [])

    ws_merit.cell(row=1, column=1, value="STUDENT MERIT LIST (ACADEMIC RANKING)").font = Font(name="Calibri", size=12, bold=True, color="0F172A")
    merit_headers = ["Merit Rank", "Group Rank", "Roll No", "Reg No", "Student Name", "Group", "Candidate Type", "GPA (With 4th)", "GPA (Without 4th)", "Grades Summary", "Passed Subjects Detail"]
    for col_idx, h_text in enumerate(merit_headers, start=1):
        c = ws_merit.cell(row=2, column=col_idx, value=h_text)
        c.fill = header_fill
        c.font = header_font
        c.alignment = align_center
        c.border = thin_border

    curr_r_m = 3
    if not merit_students:
        c = ws_merit.cell(row=curr_r_m, column=1, value="No passed students to rank.")
        c.font = normal_font
        ws_merit.merge_cells(start_row=curr_r_m, start_column=1, end_row=curr_r_m, end_column=11)
    else:
        for st in merit_students:
            a_plus = st.get("A_Plus_Count", 0)
            a_cnt = st.get("A_Count", 0)
            grades_brief = f"{a_plus} A+, {a_cnt} A" if a_plus or a_cnt else "Passed"
            row_vals = [
                st.get("Merit_Position", "-"),
                st.get("Group_Merit_Position", "-"),
                st.get("Roll", "-"),
                st.get("Registration", "-"),
                st.get("Name", "N/A"),
                st.get("Group", "-"),
                st.get("Type", "REGULAR"),
                st.get("GPA_With_4th", "-"),
                st.get("GPA_Without_4th", "-"),
                grades_brief,
                st.get("Passed_Grades_Summary", "")
            ]
            for col_idx, val in enumerate(row_vals, start=1):
                c = ws_merit.cell(row=curr_r_m, column=col_idx, value=val)
                c.font = normal_font
                c.border = thin_border
                c.alignment = align_center if col_idx in [1, 2, 3, 4, 6, 7, 8, 9, 10] else align_left
                if col_idx == 1 and val in [1, 2, 3]:
                    c.fill = PatternFill(start_color="FEF08A", end_color="FEF08A", fill_type="solid")
                    c.font = bold_font
            curr_r_m += 1

    for col in ws_merit.columns:
        max_len = max(len(str(cell.value or "")) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws_merit.column_dimensions[col_letter].width = min(max(max_len + 3, 12), 45)

    # =========================================================================
    # SHEET 9: SUBJECT-WISE FAILED LIST (ASCENDING ORDER)
    # =========================================================================
    ws_fail_subj = wb.create_sheet(title="Subject Wise Failed List")
    ws_fail_subj.views.sheetView[0].showGridLines = True

    fail_subj_data = analysis.get("failed_list_subject_wise", {})
    fail_entries = fail_subj_data.get("entries", [])

    ws_fail_subj.cell(row=1, column=1, value="FAILED LIST SUBJECT WISE (ASCENDING ORDER)").font = Font(name="Calibri", size=12, bold=True, color="0F172A")
    fail_subj_headers = ["Failed Subject", "Roll No", "Reg No", "Student Name", "Group", "Candidate Type", "Fails Count", "All Failed Subjects", "Passed Subject Grades"]
    for col_idx, h_text in enumerate(fail_subj_headers, start=1):
        c = ws_fail_subj.cell(row=2, column=col_idx, value=h_text)
        c.fill = header_fill
        c.font = header_font
        c.alignment = align_center
        c.border = thin_border

    curr_r_f = 3
    if not fail_entries:
        c = ws_fail_subj.cell(row=curr_r_f, column=1, value="No failed students recorded.")
        c.font = normal_font
        ws_fail_subj.merge_cells(start_row=curr_r_f, start_column=1, end_row=curr_r_f, end_column=9)
    else:
        for item in fail_entries:
            all_f = item.get("All_Failed_Subjects", [])
            row_vals = [
                item.get("Failed_Subject", "-"),
                item.get("Roll", "-"),
                item.get("Registration", "-"),
                item.get("Name", "N/A"),
                item.get("Group", "-"),
                item.get("Type", "REGULAR"),
                item.get("Failed_Count", 1),
                ", ".join(all_f) if isinstance(all_f, list) else str(all_f),
                item.get("Passed_Grades_Summary", "")
            ]
            for col_idx, val in enumerate(row_vals, start=1):
                c = ws_fail_subj.cell(row=curr_r_f, column=col_idx, value=val)
                c.font = normal_font
                c.border = thin_border
                c.alignment = align_center if col_idx in [2, 3, 5, 6, 7] else align_left
                if col_idx == 1:
                    c.font = bold_font
                    c.fill = fail_badge_fill
            curr_r_f += 1

    for col in ws_fail_subj.columns:
        max_len = max(len(str(cell.value or "")) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws_fail_subj.column_dimensions[col_letter].width = min(max(max_len + 3, 12), 45)

    # =========================================================================
    # SHEET 10: FAIL ANALYSIS (SUBJECT-WISE: REGULAR, IRREGULAR, TOTAL)
    # =========================================================================
    ws_fail_tables = wb.create_sheet(title="Fail Analysis")
    ws_fail_tables.views.sheetView[0].showGridLines = True

    fail_tables_data = analysis.get("subject_fail_tables", {})
    one_table_all = fail_tables_data.get("one_subject_fail", {}).get("all", {})
    one_table_groups = fail_tables_data.get("one_subject_fail", {}).get("by_group", {})

    two_table_all = fail_tables_data.get("two_subject_fail", {}).get("all", {})
    two_table_groups = fail_tables_data.get("two_subject_fail", {}).get("by_group", {})

    three_table_all = fail_tables_data.get("three_plus_subject_fail", {}).get("all", {})
    three_table_groups = fail_tables_data.get("three_plus_subject_fail", {}).get("by_group", {})

    all_groups_list = sorted(list(analysis.get("groups_summary", {}).keys()))

    def write_excel_fail_block(ws, start_col, start_r, scope_title, cat_title, subj_dict):
        # Row 1: Scope (All / Group: ...) across 4 columns
        c_scope = ws.cell(row=start_r, column=start_col, value=scope_title)
        c_scope.font = Font(name="Calibri", size=11, bold=True, color="0F172A")
        c_scope.fill = PatternFill(start_color="E2E8F0", end_color="E2E8F0", fill_type="solid")
        c_scope.alignment = align_center
        ws.merge_cells(start_row=start_r, start_column=start_col, end_row=start_r, end_column=start_col + 3)
        for ci in range(start_col, start_col + 4):
            ws.cell(row=start_r, column=ci).border = thin_border

        # Row 2: Category (e.g. One Subject Fail (Without 4th Subject)) across 4 columns
        c_cat = ws.cell(row=start_r + 1, column=start_col, value=cat_title)
        c_cat.font = Font(name="Calibri", size=10, bold=True, color="1E293B")
        c_cat.fill = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
        c_cat.alignment = align_center
        ws.merge_cells(start_row=start_r + 1, start_column=start_col, end_row=start_r + 1, end_column=start_col + 3)
        for ci in range(start_col, start_col + 4):
            ws.cell(row=start_r + 1, column=ci).border = thin_border

        # Row 3: 4 Headers: Subject | Regular | Irregular | Total
        headers = ["Subject", "Regular", "Irregular", "Total"]
        for h_idx, h_text in enumerate(headers):
            h_cell = ws.cell(row=start_r + 2, column=start_col + h_idx, value=h_text)
            h_cell.font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
            h_cell.fill = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
            h_cell.alignment = align_center
            h_cell.border = thin_border

        # Rows
        r_idx = start_r + 3
        tot_reg = 0
        tot_irreg = 0
        tot_all = 0

        if not subj_dict:
            e1 = ws.cell(row=r_idx, column=start_col, value="No failures recorded")
            e1.font = Font(name="Calibri", size=10, italic=True, color="64748B")
            e1.alignment = align_left
            for ci in range(start_col, start_col + 4):
                c = ws.cell(row=r_idx, column=ci)
                c.border = thin_border
                if ci > start_col:
                    c.value = 0
                    c.font = Font(name="Calibri", size=10, color="64748B")
                    c.alignment = align_center
            r_idx += 1
        else:
            for s_name, item_val in subj_dict.items():
                if isinstance(item_val, dict):
                    reg_c = item_val.get("regular", 0)
                    irreg_c = item_val.get("irregular", 0)
                    total_c = item_val.get("total", reg_c + irreg_c)
                else:
                    reg_c = item_val
                    irreg_c = 0
                    total_c = item_val

                s_c = ws.cell(row=r_idx, column=start_col, value=s_name)
                c_reg = ws.cell(row=r_idx, column=start_col + 1, value=reg_c)
                c_irreg = ws.cell(row=r_idx, column=start_col + 2, value=irreg_c)
                c_tot = ws.cell(row=r_idx, column=start_col + 3, value=total_c)

                s_c.font = normal_font
                c_reg.font = Font(name="Calibri", size=10, color="2563EB")
                c_irreg.font = Font(name="Calibri", size=10, color="D97706")
                c_tot.font = Font(name="Calibri", size=10, bold=True, color="DC2626")

                for cell_item in [s_c, c_reg, c_irreg, c_tot]:
                    cell_item.border = thin_border
                s_c.alignment = align_left
                c_reg.alignment = align_center
                c_irreg.alignment = align_center
                c_tot.alignment = align_center

                tot_reg += reg_c
                tot_irreg += irreg_c
                tot_all += total_c
                r_idx += 1

        # Total Row
        t1 = ws.cell(row=r_idx, column=start_col, value="Total")
        t_reg = ws.cell(row=r_idx, column=start_col + 1, value=tot_reg)
        t_irreg = ws.cell(row=r_idx, column=start_col + 2, value=tot_irreg)
        t_all = ws.cell(row=r_idx, column=start_col + 3, value=tot_all)

        for t_cell in [t1, t_reg, t_irreg, t_all]:
            t_cell.font = Font(name="Calibri", size=10, bold=True, color="0F172A")
            t_cell.fill = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
            t_cell.border = thin_border
            t_cell.alignment = align_center
        t1.alignment = align_left

        return r_idx + 2

    # =========================================================================
    # PART 1: FAIL SUBJECT SUMMARY MATRIX (AT FIRST)
    # =========================================================================
    matrix_data = analysis.get("fail_subject_summary_matrix") or build_fail_subject_summary_matrix([])
    m_groups = matrix_data.get("groups", ["Science", "Business Studies", "Humanities"])
    m_rows = matrix_data.get("rows", [])
    m_subtotal_row = matrix_data.get("subtotal_row", {})
    m_total_row = matrix_data.get("total_row", {})

    matrix_total_cols = 1 + (len(m_groups) * 2) + 2 + 1  # 10 columns for 3 groups

    # Row 2: Title Banner
    c_m_title = ws_fail_tables.cell(row=2, column=1, value="FAIL SUBJECT SUMMARY ANALYSIS (GROUP & CANDIDATE TYPE BREAKDOWN - WITHOUT 4TH SUBJECT)")
    c_m_title.font = Font(name="Calibri", size=12, bold=True, color="FFFFFF")
    c_m_title.fill = PatternFill(start_color="0F172A", end_color="0F172A", fill_type="solid")
    c_m_title.alignment = align_center
    ws_fail_tables.merge_cells(start_row=2, start_column=1, end_row=2, end_column=matrix_total_cols)
    ws_fail_tables.row_dimensions[2].height = 26
    for ci in range(1, matrix_total_cols + 1):
        ws_fail_tables.cell(row=2, column=ci).border = thin_border

    # Row 3 & 4: Headers
    c_fs = ws_fail_tables.cell(row=3, column=1, value="Fail Subject")
    c_fs.font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    c_fs.fill = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
    c_fs.alignment = align_center
    ws_fail_tables.merge_cells(start_row=3, start_column=1, end_row=4, end_column=1)
    ws_fail_tables.cell(row=3, column=1).border = thin_border
    ws_fail_tables.cell(row=4, column=1).border = thin_border

    col_cur = 2
    for grp in m_groups:
        c_g = ws_fail_tables.cell(row=3, column=col_cur, value=grp)
        c_g.font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
        g_fill_hex = "1E3A8A" if "Sci" in grp else ("0369A1" if "Bus" in grp or "Com" in grp else ("0F766E" if "Hum" in grp or "Art" in grp else "334155"))
        c_g.fill = PatternFill(start_color=g_fill_hex, end_color=g_fill_hex, fill_type="solid")
        c_g.alignment = align_center
        ws_fail_tables.merge_cells(start_row=3, start_column=col_cur, end_row=3, end_column=col_cur + 1)
        ws_fail_tables.cell(row=3, column=col_cur).border = thin_border
        ws_fail_tables.cell(row=3, column=col_cur + 1).border = thin_border

        c_reg = ws_fail_tables.cell(row=4, column=col_cur, value="Regular")
        c_reg.font = Font(name="Calibri", size=9, bold=True, color="1E3A8A")
        c_reg.fill = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
        c_reg.alignment = align_center
        c_reg.border = thin_border

        c_irreg = ws_fail_tables.cell(row=4, column=col_cur + 1, value="Irregular")
        c_irreg.font = Font(name="Calibri", size=9, bold=True, color="D97706")
        c_irreg.fill = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
        c_irreg.alignment = align_center
        c_irreg.border = thin_border
        col_cur += 2

    # SubTotal columns in Row 3-4
    c_subtot = ws_fail_tables.cell(row=3, column=col_cur, value="SubTotal")
    c_subtot.font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    c_subtot.fill = PatternFill(start_color="334155", end_color="334155", fill_type="solid")
    c_subtot.alignment = align_center
    ws_fail_tables.merge_cells(start_row=3, start_column=col_cur, end_row=3, end_column=col_cur + 1)
    ws_fail_tables.cell(row=3, column=col_cur).border = thin_border
    ws_fail_tables.cell(row=3, column=col_cur + 1).border = thin_border

    c_st_reg = ws_fail_tables.cell(row=4, column=col_cur, value="Regular")
    c_st_reg.font = Font(name="Calibri", size=9, bold=True, color="0F172A")
    c_st_reg.fill = PatternFill(start_color="E2E8F0", end_color="E2E8F0", fill_type="solid")
    c_st_reg.alignment = align_center
    c_st_reg.border = thin_border

    c_st_irreg = ws_fail_tables.cell(row=4, column=col_cur + 1, value="Irregular")
    c_st_irreg.font = Font(name="Calibri", size=9, bold=True, color="0F172A")
    c_st_irreg.fill = PatternFill(start_color="E2E8F0", end_color="E2E8F0", fill_type="solid")
    c_st_irreg.alignment = align_center
    c_st_irreg.border = thin_border
    col_cur += 2

    # Total column in Row 3-4 (merged)
    c_tot_hdr = ws_fail_tables.cell(row=3, column=col_cur, value="Total")
    c_tot_hdr.font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    c_tot_hdr.fill = PatternFill(start_color="991B1B", end_color="991B1B", fill_type="solid")
    c_tot_hdr.alignment = align_center
    ws_fail_tables.merge_cells(start_row=3, start_column=col_cur, end_row=4, end_column=col_cur)
    ws_fail_tables.cell(row=3, column=col_cur).border = thin_border
    ws_fail_tables.cell(row=4, column=col_cur).border = thin_border

    # Data Rows (Row 5 to 9)
    r_idx = 5
    for row_item in m_rows:
        ws_fail_tables.row_dimensions[r_idx].height = 20
        c_lbl = ws_fail_tables.cell(row=r_idx, column=1, value=row_item.get("label", ""))
        c_lbl.font = Font(name="Calibri", size=10, bold=True, color="1E293B")
        c_lbl.alignment = align_left
        c_lbl.border = thin_border

        c_col = 2
        for grp in m_groups:
            grp_counts = row_item.get("counts", {}).get(grp, {"regular": 0, "irregular": 0, "total": 0})
            cr = ws_fail_tables.cell(row=r_idx, column=c_col, value=grp_counts.get("regular", 0))
            cr.font = Font(name="Calibri", size=10, color="2563EB")
            cr.alignment = align_center
            cr.border = thin_border

            ci = ws_fail_tables.cell(row=r_idx, column=c_col + 1, value=grp_counts.get("irregular", 0))
            ci.font = Font(name="Calibri", size=10, color="D97706")
            ci.alignment = align_center
            ci.border = thin_border
            c_col += 2

        subt = row_item.get("subtotal", {"regular": 0, "irregular": 0, "total": 0})
        csr = ws_fail_tables.cell(row=r_idx, column=c_col, value=subt.get("regular", 0))
        csr.font = Font(name="Calibri", size=10, bold=True, color="0F172A")
        csr.fill = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
        csr.alignment = align_center
        csr.border = thin_border

        csi = ws_fail_tables.cell(row=r_idx, column=c_col + 1, value=subt.get("irregular", 0))
        csi.font = Font(name="Calibri", size=10, bold=True, color="0F172A")
        csi.fill = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
        csi.alignment = align_center
        csi.border = thin_border
        c_col += 2

        ct = ws_fail_tables.cell(row=r_idx, column=c_col, value=row_item.get("total", 0))
        ct.font = Font(name="Calibri", size=10, bold=True, color="DC2626")
        ct.fill = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid")
        ct.alignment = align_center
        ct.border = thin_border
        r_idx += 1

    # Row 10: SubTotal Row
    ws_fail_tables.row_dimensions[r_idx].height = 22
    c_st_lbl = ws_fail_tables.cell(row=r_idx, column=1, value="SubTotal")
    c_st_lbl.font = Font(name="Calibri", size=10, bold=True, color="0F172A")
    c_st_lbl.fill = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
    c_st_lbl.alignment = align_left
    c_st_lbl.border = thin_border

    c_col = 2
    for grp in m_groups:
        g_st = m_subtotal_row.get("counts", {}).get(grp, {"regular": 0, "irregular": 0, "total": 0})
        cr = ws_fail_tables.cell(row=r_idx, column=c_col, value=g_st.get("regular", 0))
        cr.font = Font(name="Calibri", size=10, bold=True, color="1E3A8A")
        cr.fill = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
        cr.alignment = align_center
        cr.border = thin_border

        ci = ws_fail_tables.cell(row=r_idx, column=c_col + 1, value=g_st.get("irregular", 0))
        ci.font = Font(name="Calibri", size=10, bold=True, color="D97706")
        ci.fill = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
        ci.alignment = align_center
        ci.border = thin_border
        c_col += 2

    st_all = m_subtotal_row.get("subtotal", {"regular": 0, "irregular": 0, "total": 0})
    csr = ws_fail_tables.cell(row=r_idx, column=c_col, value=st_all.get("regular", 0))
    csr.font = Font(name="Calibri", size=10, bold=True, color="0F172A")
    csr.fill = PatternFill(start_color="E2E8F0", end_color="E2E8F0", fill_type="solid")
    csr.alignment = align_center
    csr.border = thin_border

    csi = ws_fail_tables.cell(row=r_idx, column=c_col + 1, value=st_all.get("irregular", 0))
    csi.font = Font(name="Calibri", size=10, bold=True, color="0F172A")
    csi.fill = PatternFill(start_color="E2E8F0", end_color="E2E8F0", fill_type="solid")
    csi.alignment = align_center
    csi.border = thin_border
    c_col += 2

    c_st_tot = ws_fail_tables.cell(row=r_idx, column=c_col, value=m_subtotal_row.get("total", 0))
    c_st_tot.font = Font(name="Calibri", size=11, bold=True, color="DC2626")
    c_st_tot.fill = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid")
    c_st_tot.alignment = align_center
    c_st_tot.border = thin_border
    r_idx += 1

    # Row 11: Total Row (Merged across Regular + Irregular per group)
    ws_fail_tables.row_dimensions[r_idx].height = 24
    c_tot_lbl = ws_fail_tables.cell(row=r_idx, column=1, value="Total")
    c_tot_lbl.font = Font(name="Calibri", size=11, bold=True, color="0F172A")
    c_tot_lbl.fill = PatternFill(start_color="E2E8F0", end_color="E2E8F0", fill_type="solid")
    c_tot_lbl.alignment = align_left
    c_tot_lbl.border = thin_border

    c_col = 2
    for grp in m_groups:
        grp_tot = m_total_row.get("counts", {}).get(grp, 0)
        c_gt = ws_fail_tables.cell(row=r_idx, column=c_col, value=grp_tot)
        c_gt.font = Font(name="Calibri", size=10, bold=True, color="0F172A")
        c_gt.fill = PatternFill(start_color="E2E8F0", end_color="E2E8F0", fill_type="solid")
        c_gt.alignment = align_center
        ws_fail_tables.merge_cells(start_row=r_idx, start_column=c_col, end_row=r_idx, end_column=c_col + 1)
        ws_fail_tables.cell(row=r_idx, column=c_col).border = thin_border
        ws_fail_tables.cell(row=r_idx, column=c_col + 1).border = thin_border
        c_col += 2

    # SubTotal total merged
    c_st_merged = ws_fail_tables.cell(row=r_idx, column=c_col, value=m_total_row.get("subtotal", 0))
    c_st_merged.font = Font(name="Calibri", size=10, bold=True, color="0F172A")
    c_st_merged.fill = PatternFill(start_color="CBD5E1", end_color="CBD5E1", fill_type="solid")
    c_st_merged.alignment = align_center
    ws_fail_tables.merge_cells(start_row=r_idx, start_column=c_col, end_row=r_idx, end_column=c_col + 1)
    ws_fail_tables.cell(row=r_idx, column=c_col).border = thin_border
    ws_fail_tables.cell(row=r_idx, column=c_col + 1).border = thin_border
    c_col += 2

    # Grand Total
    c_grand_tot = ws_fail_tables.cell(row=r_idx, column=c_col, value=m_total_row.get("total", 0))
    c_grand_tot.font = Font(name="Calibri", size=11, bold=True, color="DC2626")
    c_grand_tot.fill = PatternFill(start_color="FECACA", end_color="FECACA", fill_type="solid")
    c_grand_tot.alignment = align_center
    c_grand_tot.border = thin_border
    r_idx += 1

    # Row 13: Section Header for Detailed Subject-Wise Tables
    r_idx += 1
    c_sec_hdr = ws_fail_tables.cell(row=r_idx, column=1, value="DETAILED SUBJECT-WISE FAILURE TABLES (WITHOUT 4TH SUBJECT)")
    c_sec_hdr.font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    c_sec_hdr.fill = PatternFill(start_color="0F172A", end_color="0F172A", fill_type="solid")
    c_sec_hdr.alignment = align_center
    ws_fail_tables.merge_cells(start_row=r_idx, start_column=1, end_row=r_idx, end_column=14)
    ws_fail_tables.row_dimensions[r_idx].height = 24
    for ci in range(1, 15):
        ws_fail_tables.cell(row=r_idx, column=ci).border = thin_border

    start_detail_r = r_idx + 2

    # =========================================================================
    # PART 2: DETAILED SUBJECT-WISE TABLES (SIDE-BY-SIDE CATEGORIES)
    # =========================================================================
    # Render Category 1: One Subject Fail (Cols A-D, start_col = 1)
    curr_r_1 = start_detail_r
    curr_r_1 = write_excel_fail_block(ws_fail_tables, 1, curr_r_1, "All", "One Subject Fail (Without 4th Subject)", one_table_all)
    for grp in all_groups_list:
        curr_r_1 = write_excel_fail_block(ws_fail_tables, 1, curr_r_1, f"Group: {grp}", "One Subject Fail (Without 4th Subject)", one_table_groups.get(grp, {}))

    # Render Category 2: Two Subject Fail (Cols F-I, start_col = 6)
    curr_r_2 = start_detail_r
    curr_r_2 = write_excel_fail_block(ws_fail_tables, 6, curr_r_2, "All", "Two Subject Fail (Without 4th Subject)", two_table_all)
    for grp in all_groups_list:
        curr_r_2 = write_excel_fail_block(ws_fail_tables, 6, curr_r_2, f"Group: {grp}", "Two Subject Fail (Without 4th Subject)", two_table_groups.get(grp, {}))

    # Render Category 3: Three or More Subject Fail (Cols K-N, start_col = 11)
    curr_r_3 = start_detail_r
    curr_r_3 = write_excel_fail_block(ws_fail_tables, 11, curr_r_3, "All", "Three or More Subject Fail (Without 4th Subject)", three_table_all)
    for grp in all_groups_list:
        curr_r_3 = write_excel_fail_block(ws_fail_tables, 11, curr_r_3, f"Group: {grp}", "Three or More Subject Fail (Without 4th Subject)", three_table_groups.get(grp, {}))

    # Auto-fit columns
    for col in ws_fail_tables.columns:
        vals = [str(cell.value or "") for cell in col if cell.value is not None]
        max_len = max((len(v) for v in vals), default=0)
        col_letter = get_column_letter(col[0].column)
        if col[0].column in [2, 3, 4, 5, 6, 7, 8, 9, 10, 12, 13, 14]:
            ws_fail_tables.column_dimensions[col_letter].width = max(max_len + 3, 12)
        else:
            ws_fail_tables.column_dimensions[col_letter].width = max(max_len + 3, 20)

    stream = io.BytesIO()
    wb.save(stream)
    stream.seek(0)
    return stream


def generate_demo_dataset(count: int = 50) -> List[Dict[str, Any]]:
    """
    Generates a realistic mock batch of student result records for demonstration,
    featuring Science, Business Studies, and Humanities groups, Regular and Irregular candidates,
    realistic GPA distributions (GPA 5.0, 4.x, 3.x), 1-subject fails, 2-subject fails, and absentees.
    """
    import random

    names = [
        "MD. TANVIR HASAN", "NUSRAT JAHAN", "RAKIBUL ISLAM", "FATEMA AKTER",
        "SAKIB AL HASAN", "SADIA SULTANA", "ARIFUR RAHMAN", "MIM AKTER",
        "MEHEDI HASAN", "JANNATUL MAOWA", "KAMRUN NAHAR", "DIPA AKTER",
        "MAHIN AHMED", "SUMAIYA AKTER", "FARHAN TANVIR", "NAFISA ISLAM",
        "SHAKIL KHAN", "TASNIA RAHMAN", "AL AMIN", "SABRINA JAHAN",
        "ZAHID HASAN", "SUMONA PARVIN", "MAHMUDUL HASAN", "ANIKA TAHSIN"
    ]
    groups = ["SCIENCE", "BUSINESS STUDIES", "HUMANITIES"]
    types = ["REGULAR", "REGULAR", "REGULAR", "IRREGULAR"]

    curricula = {
        "SCIENCE": [
            ("101", "BANGLA"), ("107", "ENGLISH"), ("275", "INFORMATION & COMMUNICATION TECHNOLOGY"),
            ("174", "PHYSICS"), ("175", "CHEMISTRY"), ("176", "BIOLOGY"), ("129", "HIGHER MATHEMATICS (4TH)")
        ],
        "BUSINESS STUDIES": [
            ("101", "BANGLA"), ("107", "ENGLISH"), ("275", "INFORMATION & COMMUNICATION TECHNOLOGY"),
            ("253", "ACCOUNTING"), ("277", "BUSINESS ORGANIZATION & MANAGEMENT"),
            ("286", "PRODUCTION MANAGEMENT & MARKETING"), ("292", "FINANCE, BANKING & INSURANCE (4TH)")
        ],
        "HUMANITIES": [
            ("101", "BANGLA"), ("107", "ENGLISH"), ("275", "INFORMATION & COMMUNICATION TECHNOLOGY"),
            ("109", "ECONOMICS"), ("269", "CIVICS & GOOD GOVERNANCE"), ("304", "SOCIOLOGY"),
            ("121", "LOGIC (4TH)")
        ]
    }

    records = []
    for i in range(count):
        grp = groups[i % len(groups)]
        stype = types[i % len(types)]
        roll = str(160000 + i)
        reg = str(1512345600 + i)
        name = names[i % len(names)] + (f" ({i+1})" if i >= len(names) else "")
        subjects = curricula[grp]

        # Define performance categories
        if i == count - 1:
            # 1 Absent student
            rec = {
                "Roll": roll,
                "Registration": reg,
                "Name": name,
                "Father Name": "MD. ABDUR RAHIM",
                "Mother Name": "MRS. RAHIMA BEGUM",
                "Institute": "CUMILLA SHIKKHA BOARD GOVT. MODEL COLLEGE (132092)",
                "Board": "Cumilla",
                "Group": grp,
                "Type": stype,
                "Exam": "HSC",
                "Year": "2026",
                "Result": "ABSENT",
                "GPA": "0.00",
                "Subject_Grades": {},
                "Subject_Grades_Summary": "ABSENT"
            }
            records.append(rec)
            continue

        subj_grades = {}
        fail_count = 0
        if i in [3, 9, 17, 26, 35]:
            target_fails = 1
        elif i in [7, 21, 39]:
            target_fails = 2
        elif i in [14, 31]:
            target_fails = 3
        elif i in [28]:
            target_fails = 4
        elif i in [44]:
            target_fails = 5
        else:
            target_fails = 0

        for s_idx, (code, sname) in enumerate(subjects):
            full_name = f"{code} - {sname}"
            if target_fails > 0 and s_idx == 3:
                subj_grades[full_name] = "F"
                fail_count += 1
            elif target_fails > 1 and s_idx == 4:
                subj_grades[full_name] = "F"
                fail_count += 1
            elif target_fails > 2 and s_idx == 5:
                subj_grades[full_name] = "F"
                fail_count += 1
            elif target_fails > 3 and s_idx == 1:
                subj_grades[full_name] = "F"
                fail_count += 1
            elif target_fails > 4 and s_idx == 2:
                subj_grades[full_name] = "F"
                fail_count += 1
            else:
                if i % 6 == 0:
                    subj_grades[full_name] = "A+"
                elif i % 6 in [1, 2]:
                    subj_grades[full_name] = "A" if s_idx % 2 == 0 else "A+"
                elif i % 6 in [4]:
                    subj_grades[full_name] = "A-" if s_idx % 2 == 0 else "B"
                else:
                    subj_grades[full_name] = "B" if s_idx % 2 == 0 else "C"

        if fail_count > 0:
            result = f"FAILED IN {fail_count} SUBJECT" + ("S" if fail_count > 1 else "")
            gpa = "0.00"
        else:
            result = "PASSED"
            if all(g == "A+" for g in subj_grades.values()):
                gpa = "5.00"
            elif i % 6 == 1:
                gpa = "4.83"
            elif i % 6 == 2:
                gpa = "4.58"
            elif i % 6 == 4:
                gpa = "3.75"
            else:
                gpa = "3.25"

        summary_parts = [f"{k}: {v}" for k, v in subj_grades.items()]
        rec = {
            "Roll": roll,
            "Registration": reg,
            "Name": name,
            "Father Name": "MD. ANWAR HOSSAIN",
            "Mother Name": "MRS. FARIDA YASMIN",
            "Institute": "CUMILLA SHIKKHA BOARD GOVT. MODEL COLLEGE (132092)",
            "Board": "Cumilla",
            "Group": grp,
            "Type": stype,
            "Exam": "HSC",
            "Year": "2026",
            "Result": result,
            "GPA": gpa,
            "Subject_Grades": subj_grades,
            "Subject_Grades_Summary": ", ".join(summary_parts)
        }
        records.append(rec)

    return records

