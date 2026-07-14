import json
import pandas as pd

def load_json(json_file):
    with open(json_file, "r") as file:
        config = json.load(file)
    return config

def load_excel(excel_file):
    workbook = pd.ExcelFile(excel_file)
    return workbook

def validate_sheet_list(config, workbook):

    results = []

    json_sheets = config["sheetlist"]
    excel_sheets = workbook.sheet_names

    print("\n========== SHEET LIST VALIDATION ==========\n")

    for sheet in json_sheets:

        if sheet in excel_sheets:
            status = "P"
            print(f"{sheet} - PASS")
        else:
            status = "F"
            print(f"{sheet} - FAIL")

        results.append({
            "Sheet Name": sheet,
            "Field": "Sheet",
            "Type": "Existence Testing",
            "Status (P/F)": status,
            "Failed Count": 0 if status == "P" else 1
        })

    return results

def validate_sheet_columns(config, excel_file):

    results = []

    sheet_validation = config["sheetvalidation"]

    print("\n========== SHEET COLUMN VALIDATION ==========\n")

    for sheet_name, sheet_info in sheet_validation.items():

        df = pd.read_excel(excel_file, sheet_name=sheet_name)

        excel_columns = []

        for col in df.columns:
            clean_col = (
                str(col)
                .strip()
                .lower()
                .replace(" ", "")
                .replace("_", "")
            )
            excel_columns.append(clean_col)

        print(f"\n{sheet_name}")

        for field in sheet_info["fields"]:

            json_column = (
                field["name"]
                .strip()
                .lower()
                .replace(" ", "")
                .replace("_", "")
            )

            if json_column in excel_columns:
                status = "P"
                print(f"{field['name']} - PASS")
            else:
                status = "F"
                print(f"{field['name']} - FAIL")

            results.append({
                "Sheet Name": sheet_name,
                "Field": field["name"],
                "Type": "Existence Testing",
                "Status (P/F)": status,
                "Failed Count": 0 if status == "P" else 1
            })

    return results
def validate_duplicate_check(config, excel_file):

    results = []

    duplicate_validation = config["dataqualityvalidation"]["duplicatecheck"]

    print("\n========== DUPLICATE CHECK VALIDATION ==========\n")

    for rule in duplicate_validation:

        sheet_name = rule["sheetname"]
        columns = rule["columns"]

        df = pd.read_excel(excel_file, sheet_name=sheet_name)

        duplicate_rows = df[df.duplicated(subset=columns, keep=False)]

        duplicate_count = len(duplicate_rows)

        if duplicate_count == 0:
            status = "P"
            print(f"{sheet_name} - PASS")
        else:
            status = "F"
            print(f"{sheet_name} - FAIL ({duplicate_count} duplicate rows)")

        for column in columns:
            results.append({
                "Sheet Name": sheet_name,
                "Field": column,
                "Type": "Duplicate Check",
                "Status (P/F)": status,
                "Failed Count": duplicate_count
            })

    return results

def validate_null_check(config, excel_file):
    results = []

    null_validation = config["dataqualityvalidation"]["nullcheck"]

    print("\n========== NULL CHECK VALIDATION ==========\n")

    for rule in null_validation:

        sheet_name = rule["sheetname"]
        columns = rule["columns"]

        df = pd.read_excel(excel_file, sheet_name=sheet_name)

        for column in columns:

            null_count = df[column].isnull().sum()

            if null_count == 0:
                status = "P"
                print(f"{sheet_name} - {column} - PASS")
            else:
                status = "F"
                print(f"{sheet_name} - {column} - FAIL ({null_count} null values)")

            results.append({
                "Sheet Name": sheet_name,
                "Field": column,
                "Type": "Null Check",
                "Status (P/F)": status,
                "Failed Count": null_count
            })

    return results

def validate_formula_check(config, excel_file):
    results = []

    formula_validation = config["formulavalidation"]

    print("\n========== FORMULA VALIDATION ==========\n")

    df = pd.read_excel(excel_file, sheet_name="KPI")

    for rule in formula_validation:

        formula_name = rule["name"]

        if formula_name == "ACR Calculation":

            acr = ((df["AssociatedPayments"] +
                    df["AssociatedAdjustments"]) /
                    df["Current_Charges"]) * 100

            failed_count = (acr <= 100).sum()

        elif formula_name == "GCR Calculation":

            gcr = (df["AssociatedPayments"] /
                   df["Current_Charges"]) * 100

            failed_count = (gcr <= 100).sum()

        elif formula_name == "Gross AR Calculation":

            calculated_gross_ar = df["InsuranceAR"] + df["PatientAR"]

            failed_count = (calculated_gross_ar != df["GrossAR"]).sum()

        if failed_count == 0:
            status = "P"
            print(f"{formula_name} - PASS")
        else:
            status = "F"
            print(f"{formula_name} - FAIL ({failed_count} rows)")

        results.append({
            "Sheet Name": "KPI",
            "Field": formula_name,
            "Type": "Formula Validation",
            "Status (P/F)": status,
            "Failed Count": failed_count
        })

    return results

def main():

    json_file = "Kpi_automation_phase1.json"
    excel_file = "KPI_Sample.xlsx"

    config = load_json(json_file)
    workbook = load_excel(excel_file)

    print("JSON Loaded Successfully")
    print("Excel Loaded Successfully")

    sheet_results = validate_sheet_list(config, workbook)
    column_results = validate_sheet_columns(config, excel_file)
    duplicate_results = validate_duplicate_check(config, excel_file)
    null_results = validate_null_check(config, excel_file)
    formula_results = validate_formula_check(config, excel_file)

    final_results = sheet_results + column_results + duplicate_results + null_results + formula_results

    report = pd.DataFrame(final_results)

    report.to_excel("Validation_Report.xlsx", index=False)

    print("\nValidation_Report.xlsx created successfully.")


if __name__ == "__main__":
    main()