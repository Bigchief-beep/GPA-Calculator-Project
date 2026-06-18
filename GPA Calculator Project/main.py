import pandas as pd

# ======================================
# GPA CALCULATOR USING PANDAS
# ======================================

# Excel file name
file_name = "AcademicData.xlsx"

# ======================================
# 1. READ DATA FROM ALL SHEETS
# ======================================

students_df = pd.read_excel(file_name, sheet_name="Students")
subjects_df = pd.read_excel(file_name, sheet_name="Subjects")
scores_df = pd.read_excel(file_name, sheet_name="RawScores")

print("\n========== STUDENTS ==========")
print(students_df)

print("\n========== SUBJECTS ==========")
print(subjects_df)

print("\n========== RAW SCORES ==========")
print(scores_df)

# ======================================
# 2. JOIN DATA USING StudentID
# ======================================

merged_df = pd.merge(
    scores_df,
    students_df,
    on="StudentID"
)

# ======================================
# 3. JOIN DATA USING SubjectCode
# ======================================

merged_df = pd.merge(
    merged_df,
    subjects_df,
    on="SubjectCode"
)

print("\n========== MERGED DATA ==========")
print(merged_df)

# ======================================
# 4. CONVERT SCORE TO LETTER GRADE
# ======================================

def get_grade(score):
    if score >= 85:
        return "A"
    elif score >= 80:
        return "A-"
    elif score >= 75:
        return "B+"
    elif score >= 70:
        return "B"
    elif score >= 65:
        return "B-"
    elif score >= 60:
        return "C+"
    elif score >= 55:
        return "C"
    elif score >= 50:
        return "D"
    else:
        return "E"

merged_df["Grade"] = merged_df["Score"].apply(get_grade)

# ======================================
# 5. CONVERT LETTER GRADE TO GRADE POINT
# ======================================

grade_points = {
    "A": 4.0,
    "A-": 3.7,
    "B+": 3.3,
    "B": 3.0,
    "B-": 2.7,
    "C+": 2.3,
    "C": 2.0,
    "D": 1.0,
    "E": 0.0
}

merged_df["GradePoint"] = merged_df["Grade"].map(grade_points)

# ======================================
# 6. CALCULATE WEIGHTED POINTS
# ======================================

merged_df["WeightedPoint"] = (
    merged_df["GradePoint"] * merged_df["SKS"]
)

# ======================================
# 7. CALCULATE GPA FOR EACH STUDENT
# ======================================

gpa_report = (
    merged_df
    .groupby(["StudentID", "StudentName"])
    .agg(
        TotalWeightedPoints=("WeightedPoint", "sum"),
        TotalSKS=("SKS", "sum")
    )
    .reset_index()
)

gpa_report["GPA"] = (
    gpa_report["TotalWeightedPoints"]
    / gpa_report["TotalSKS"]
).round(2)

# ======================================
# 8. DISPLAY GPA OF EACH STUDENT
# ======================================

print("\n========== GPA REPORT ==========")

for _, row in gpa_report.iterrows():
    print(
        f"{row['StudentID']} - "
        f"{row['StudentName']} : GPA = {row['GPA']}"
    )

# ======================================
# 9. CLASS AVERAGE GPA
# ======================================

class_average = round(
    gpa_report["GPA"].mean(),
    2
)

print("\nClass Average GPA:", class_average)

# ======================================
# 10. SAVE REPORT TO NEW SHEET
# ======================================

with pd.ExcelWriter(
    file_name,
    engine="openpyxl",
    mode="a",
    if_sheet_exists="replace"
) as writer:
    gpa_report.to_excel(
        writer,
        sheet_name="GPA_Report",
        index=False
    )

print("\nGPA report successfully saved to sheet 'GPA_Report'")