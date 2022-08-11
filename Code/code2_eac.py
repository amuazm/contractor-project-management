from openpyxl import load_workbook
import shutil

# Copy file
src = "./Excel Files/Output/Result.xlsx"
dest = "./Excel Files/Output/Result_EAC.xlsx"
shutil.copyfile(src, dest)

# Result WB and WS
wb_result = load_workbook(dest)
ws_reports = wb_result["Reports"]
ws_budget = wb_result["Budget"]

# Months Passed
ws_reports.insert_cols(3)
ws_reports["C1"] = "Months Passed"
i = 1
current_project = ""
for row in ws_reports.iter_rows(min_row=2):
    if row[0].value != current_project:
        current_project = row[0].value
        i = 1
    row[2].value = i
    i += 1

# Overall Budgets dict
overall_budgets = {}
for row in ws_budget.iter_rows(min_row=2):
    overall_budgets[row[0].value] = row[4].value

# EAC/CPI = ACWP + (Overall Budget - BCWP) / CPI
ws_reports.insert_cols(12)
ws_reports["L1"].value = "EAC"
for row in ws_reports.iter_rows(min_row=2):
    row[11].value = row[4].value + (overall_budgets[row[0].value] - row[5].value) / row[7].value
    row[11].style = "Currency"

# Project durations dict
project_durations = {}
for row in ws_budget.iter_rows(min_row=2):
    project_durations[row[0].value] = row[2].value

# EAC(t)(ED)/SPI = Months Passed + (max(Duration (Months), Months Passed) - ED) / SPI
# ED = Months Passed * SPI
ws_reports.insert_cols(13)
ws_reports["M1"].value = "EAC(t)"
for row in ws_reports.iter_rows(min_row=2):
    row[12].value = row[2].value + (max(project_durations[row[0].value], row[2].value) - row[2].value * row[9].value) / row[9].value

# VAC = Overall Budget - EAC
ws_reports.insert_cols(14)
ws_reports["N1"].value = "VAC"
for row in ws_reports.iter_rows(min_row=2):
    row[13].value = overall_budgets[row[0].value] - row[11].value
    row[13].style = "Currency"

# VAC(t) = Duration (Months) - EAC(t)
ws_reports.insert_cols(15)
ws_reports["O1"].value = "VAC(t)"
for row in ws_reports.iter_rows(min_row=2):
    row[14].value = project_durations[row[0].value] - row[12].value

wb_result.save(dest)