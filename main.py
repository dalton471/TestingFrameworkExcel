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

            failed_count = (acr > 100).sum()

        elif formula_name == "GCR Calculation":

            gcr = (df["AssociatedPayments"] /
                   df["Current_Charges"]) * 100

            failed_count = (gcr > 100).sum()

        elif formula_name == "Gross AR Calculation":
            calculated_gross_ar = df["InsuranceAR"] + df["PatientAR"]
            comparison = ~(
            calculated_gross_ar.round(2).fillna(0)
            .eq(df["GrossAR"].round(2).fillna(0))
            ) 
            failed_rows = df.loc[
                comparison,
                ["InsuranceAR", "PatientAR", "GrossAR"]
                ].copy()

            failed_rows["CalculatedGrossAR"] = calculated_gross_ar[comparison]
            failed_count = len(failed_rows)

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
def validate_numeric_precision(config, excel_file):

    results = []

    precision_validation = config["dataqualityvalidation"]["numericprecision"]

    print("\n========== NUMERIC PRECISION VALIDATION ==========\n")

    for rule in precision_validation:

        sheet_name = rule["sheetname"]
        precision = rule["precision"]

        df = pd.read_excel(excel_file, sheet_name=sheet_name)

        numeric_columns = df.select_dtypes(include="number").columns

        for column in numeric_columns:

            failed_count = 0

            for value in df[column].dropna():

                decimal_places = str(value).split(".")

                if len(decimal_places) == 2:
                    if len(decimal_places[1]) > precision:
                        failed_count += 1

            if failed_count == 0:
                status = "P"
                print(f"{sheet_name} - {column} - PASS")
            else:
                status = "F"
                print(f"{sheet_name} - {column} - FAIL ({failed_count} rows)")

            results.append({
                "Sheet Name": sheet_name,
                "Field": column,
                "Type": "Numeric Precision",
                "Status (P/F)": status,
                "Failed Count": failed_count
            })

    return results

def validate_business_rule(config, excel_file):

    results = []

    business_rules = config["businessrulevalidation"]

    print("\n========== BUSINESS RULE VALIDATION ==========\n")

    data_currency_df = pd.read_excel(excel_file, sheet_name="Data Currency")
    practice_detail_df = pd.read_excel(excel_file, sheet_name="Practice Detail")
    kpi_df = pd.read_excel(excel_file, sheet_name="KPI")

    data_currency_df["LatestTransactionDate"] = pd.to_datetime(
        data_currency_df["LatestTransactionDate"],
        errors="coerce"
    )

    data_currency_df["LastLoadedDate"] = pd.to_datetime(
        data_currency_df["LastLoadedDate"],
        errors="coerce"
    )

    for rule in business_rules:

        rule_name = rule["name"]

        # Rule 1
        if rule_name == "Latest Transaction Date should be less than Latest Loaded Date":

            failed_count = (
                data_currency_df["LatestTransactionDate"]
                <
                data_currency_df["LastLoadedDate"]
            ).sum()

            if failed_count == 0:
                status = "P"
                print(f"{rule_name} - PASS")
            else:
                status = "F"
                print(f"{rule_name} - FAIL ({failed_count} rows)")

            results.append({
                "Sheet Name": "Data Currency",
                "Field": rule_name,
                "Type": "Business Rule Validation",
                "Status (P/F)": status,
                "Failed Count": failed_count
            })

        elif rule_name == "Latest Transaction Date older than 3 months should display No LatestTransactionDate or NULL":

            today = pd.Timestamp.today()

            three_months_old = today - pd.DateOffset(months=3)

            failed_count = (
                (data_currency_df["LatestTransactionDate"] < three_months_old)
                &
                (
                    ~data_currency_df["LatestTransactionDate"].isna()
                )
            ).sum()

            if failed_count == 0:
                status = "P"
                print(f"{rule_name} - PASS")
            else:
                status = "F"
                print(f"{rule_name} - FAIL ({failed_count} rows)")

            results.append({
                "Sheet Name": "Data Currency",
                "Field": rule_name,
                "Type": "Business Rule Validation",
                "Status (P/F)": status,
                "Failed Count": failed_count
            })

        elif rule_name == "Practice Status of KPI sheet should match with Problem Reason of Practice Detail sheet":

            merged_df = kpi_df.merge(
                practice_detail_df,
                left_on="PracticeCode",
                right_on="practicecode",
                how="left"
            )

            comparison = (
                merged_df["Practice Status"]
                !=
                merged_df["ProblemReason"]
            )

            failed_rows = merged_df.loc[comparison,["PracticeCode", "Practice Status", "ProblemReason"]]
            print("\nFAILED ROWS:")
            print(failed_rows)

            failed_count = len(failed_rows)

            if failed_count == 0:
                status = "P"
                print(f"{rule_name} - PASS")
            else:
                status = "F"
                print(f"{rule_name} - FAIL ({failed_count} rows)")

            results.append({
                "Sheet Name": "KPI",
                "Field": rule_name,
                "Type": "Business Rule Validation",
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
    numeric_precision_results = validate_numeric_precision(config, excel_file)
    business_rule_results = validate_business_rule(config, excel_file)

    final_results = sheet_results + column_results + duplicate_results + null_results + formula_results + numeric_precision_results + business_rule_results

    report = pd.DataFrame(final_results)

    report.to_excel("Validation_Report.xlsx", index=False)

    print("\nValidation_Report.xlsx created successfully.")


if __name__ == "__main__":
    main()