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
            "Test Type": "Sheet Validation",
            "Test Name": "Sheet Existence Testing",
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
                "Test Type": "Sheet Validation",
                "Test Name": "Field Existence Testing",
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

        field_name = ", ".join(columns)

        results.append({
            "Sheet Name": sheet_name,
            "Field": field_name,
            "Test Type": "Data Quality Validation",
            "Test Name": "Duplicate Check",
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

        status = "P"
        failed_count = 0

        field_name = ", ".join(columns)

        for column in columns:

            if column not in df.columns:
                print(f"{sheet_name} - {column} - Column Not Found")
                status = "F"
                continue

            null_count = df[column].isnull().sum()

            if null_count > 0:
                print(f"{sheet_name} - {column} - FAIL ({null_count} null values)")
                status = "F"
                failed_count += null_count
            else:
                print(f"{sheet_name} - {column} - PASS")

                print("Appending one row...")

        results.append({
            "Sheet Name": sheet_name,
            "Field": field_name,
            "Test Type": "Data Quality Validation",
            "Test Name": "Null Check",
            "Status (P/F)": status,
            "Failed Count": failed_count
        })

    return results

def validate_formula_check(config, excel_file):

    results = []

    formula_rules = config["formulavalidation"]

    print("\n========== FORMULA VALIDATION ==========\n")

    df = pd.read_excel(excel_file, sheet_name="KPI")

    for rule in formula_rules:

        formula_name = rule["name"]
        formula = rule["formula"]

        try:

            calculated = df.eval(formula)

            if "targetcolumn" in rule:

                target_column = rule["targetcolumn"]

                if target_column not in df.columns:
                    raise ValueError(f"{target_column} column not found")

                failed_mask = calculated != df[target_column]

            elif "threshold" in rule and "operator" in rule:

                threshold = rule["threshold"]
                operator = rule["operator"]

                if operator == "<=":
                    failed_mask = calculated > threshold

                elif operator == "<":
                    failed_mask = calculated >= threshold

                elif operator == ">=":
                    failed_mask = calculated < threshold

                elif operator == ">":
                    failed_mask = calculated <= threshold

                elif operator == "==":
                    failed_mask = calculated != threshold

                elif operator == "!=":
                    failed_mask = calculated == threshold

                else:
                    raise ValueError(f"Unsupported operator: {operator}")

            else:
                raise ValueError("Invalid formula validation rule in JSON")

            failed_count = int(failed_mask.sum())

            if failed_count == 0:
                status = "P"
                print(f"{formula_name} - PASS")
            else:
                status = "F"
                print(f"{formula_name} - FAIL ({failed_count} rows)")

        except Exception as e:

            status = "F"
            failed_count = 1

            print(f"{formula_name} - FAIL ({str(e)})")

        results.append({
            "Sheet Name": "KPI",
            "Field": formula_name,
            "Test Type": "Formula Validation",
            "Test Name": "Formula Validation",
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

                decimal_part = str(value).split(".")

                if len(decimal_part) == 2:

                    if len(decimal_part[1]) > precision:
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
                "Test Type": "Data Quality Validation",
                "Test Name": "Numeric Precision",
                "Status (P/F)": status,
                "Failed Count": failed_count
            })

    return results

def validate_business_rule(config, excel_file):

    results = []

    business_rules = config["businessrulevalidation"]

    print("\n========== BUSINESS RULE VALIDATION ==========\n")

    for rule in business_rules:

        rule_name = rule["name"]
        sheet_name = rule["sheetname"]

        df = pd.read_excel(excel_file, sheet_name=sheet_name)

        if rule_name == "Latest Transaction Date should be less than Latest Loaded Date":

            left_column = rule["leftcolumn"]
            right_column = rule["rightcolumn"]
            distinct_column = rule["distinctcolumn"]

            left_dates = pd.to_datetime(df[left_column], errors="coerce")
            right_dates = pd.to_datetime(df[right_column], errors="coerce")

            left_text = df[left_column].astype(str).str.strip().str.lower()
            right_text = df[right_column].astype(str).str.strip().str.lower()

            comparison = (
                (left_dates >= right_dates)
                |
                (
                    left_dates.isna()
                    &
                    (left_text != "not applicable")
                )
                |
                (
                    right_dates.isna()
                    &
                    (right_text != "not applicable")
                )
            )

            failed_rows = df.loc[
                comparison,
                [distinct_column, left_column, right_column]
            ]

            print("\nFAILED ROWS:")
            print(failed_rows)

            failed_count = len(failed_rows)

            if failed_count == 0:
                status = "P"
                print(f"{rule_name} - PASS")
            else:
                status = "F"
                print(f"{rule_name} - FAIL ({failed_count} Practice(s))")

            results.append({
                "Sheet Name": sheet_name,
                "Field": rule_name,
                "Test Type": "Business Rule Validation",
                "Test Name": "Business Rule Validation",
                "Status (P/F)": status,
                "Failed Count": failed_count
            })


        elif rule_name == "Latest Transaction Date older than 3 months should display No LatestTransactionDate or NULL":

            column_name = rule["column"]
            months = rule["months"]

            df[column_name] = pd.to_datetime(
            df[column_name],
            errors="coerce"
            )

            today = pd.Timestamp.today()
            three_months_old = today - pd.DateOffset(months=months)

            comparison = (
            (df[column_name] < three_months_old)
            &
            (~df[column_name].isna())
            )

            failed_rows = df.loc[
            comparison,
            [column_name]
            ]

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
            "Sheet Name": sheet_name,
            "Field": rule_name,
            "Test Type": "Business Rule Validation",
            "Test Name": "Business Rule Validation",
            "Status (P/F)": status,
            "Failed Count": failed_count
            })

        elif rule_name == "Practice Status of KPI sheet should match with Problem Reason of Practice Detail sheet":

            left_column = rule["leftcolumn"]
            lookup_sheet = rule["lookupsheet"]
            lookup_key_left = rule["lookupkeyleft"]
            lookup_key_right = rule["lookupkeyright"]
            right_column = rule["rightcolumn"]

            lookup_df = pd.read_excel(excel_file,sheet_name=lookup_sheet)

            merged_df = df.merge(lookup_df,left_on=lookup_key_left,right_on=lookup_key_right,how="left")

            comparison = (
            merged_df[left_column].astype(str).str.strip().str.lower()
            !=
            merged_df[right_column].astype(str).str.strip().str.lower())

            failed_rows = merged_df.loc[
            comparison,
            [
            lookup_key_left,
            left_column,
            right_column
            ]]

            failed_rows = failed_rows.drop_duplicates(subset=[lookup_key_left])

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
            "Sheet Name": sheet_name,
            "Field": rule_name,
            "Test Type": "Business Rule Validation",
            "Test Name": "Business Rule Validation",
            "Status (P/F)": status,
            "Failed Count": failed_count
            })

    return results

def validate_reconciliation(config, excel_file):

    results = []

    reconciliation_rules = config["reconciliationvalidation"]

    print("\n========== RECONCILIATION VALIDATION ==========\n")

    for rule in reconciliation_rules:

        rule_name = rule["name"]
        source_sheet = rule["sourcesheet"]
        target_sheet = rule["targetsheet"]
        match_columns = rule["matchcolumns"]
        comparisons = rule["comparisons"]
        distinct_column = rule["distinctcolumn"]

        source_df = pd.read_excel(excel_file, sheet_name=source_sheet)
        target_df = pd.read_excel(excel_file, sheet_name=target_sheet)

        source_df.columns = [
            str(col).strip().replace(" ", "_")
            for col in source_df.columns
        ]

        target_df.columns = [
            str(col).strip().replace(" ", "_")
            for col in target_df.columns
        ]

        match_columns = [
            col.replace(" ", "_")
            for col in match_columns
        ]

        distinct_column = distinct_column.replace(" ", "_")

        merged_df = pd.merge(
            source_df,
            target_df,
            on=match_columns,
            how="inner"
        )

        print(f"\n{rule_name}")
        print(f"Matched Rows : {len(merged_df)}")

        total_failed = 0

        for compare in comparisons:

            source_column = compare["sourcecolumn"].replace(" ", "_")
            target_column = compare["targetcolumn"].replace(" ", "_")

            if source_column == "Gross_AR":
                source_column = "GrossAR"

            print(f"\nComparing {source_column} with {target_column}")

            if source_column not in merged_df.columns:
                print(f"Source column '{source_column}' not found.")
                continue

            if target_column not in merged_df.columns:
                print(f"Target column '{target_column}' not found.")
                continue

            source_values = pd.to_numeric(
                merged_df[source_column],
                errors="coerce"
            ).fillna(0)

            target_values = pd.to_numeric(
                merged_df[target_column],
                errors="coerce"
            ).fillna(0)

            comparison = source_values != target_values

            display_columns = list(dict.fromkeys(
                match_columns +
                [distinct_column, source_column, target_column]
            ))

            failed_rows = merged_df.loc[
                comparison,
                display_columns
            ]

            failed_rows = failed_rows.drop_duplicates(
                subset=[distinct_column]
            )

            failed_count = len(failed_rows)
            total_failed += failed_count

            if failed_count == 0:
                print("PASS")
            else:
                print(f"FAIL ({failed_count} rows)")
                print(failed_rows)

        if total_failed == 0:
            status = "P"
        else:
            status = "F"

        results.append({
            "Sheet Name": source_sheet,
            "Field": rule_name,
            "Test Type": "Reconciliation Validation",
            "Test Name": "Reconciliation Validation",
            "Status (P/F)": status,
            "Failed Count": total_failed
        })

    return results

def validate_cross_sheet(config, excel_file):
    results = []

    rules = config.get("crosssheetvalidation", [])

    print("\n========== Cross Sheet Validation ==========")

    for rule in rules:
        rule_name = rule.get("name")
        source_sheet = rule.get("sourcesheet")
        target_sheet = rule.get("targetsheet")
        match_column = rule.get("matchcolumn")
        validation_type = rule.get("validationtype")

        source_df = pd.read_excel(excel_file, sheet_name=source_sheet)
        target_df = pd.read_excel(excel_file, sheet_name=target_sheet)

        source_df.columns = source_df.columns.str.strip().str.lower()
        target_df.columns = target_df.columns.str.strip().str.lower()

        match_col = match_column.strip().lower()

        if validation_type == "valuecomparison":

            source_column = rule.get("sourcecolumn").strip().lower()
            target_column = rule.get("targetcolumn").strip().lower()

            merged_df = pd.merge(
                source_df[[match_col, source_column]],
                target_df[[match_col, target_column]],
                on=match_col,
                how="inner"
            )

            failed_count = (
                merged_df[source_column]
                .fillna("")
                .astype(str)
                .str.strip()
                .str.lower()
                !=
                merged_df[target_column]
                .fillna("")
                .astype(str)
                .str.strip()
                .str.lower()
            ).sum()

        elif validation_type == "existencecheck":

            source_keys = set(
                source_df[match_col]
                .dropna()
                .astype(str)
                .str.strip()
            )

            target_keys = set(
                target_df[match_col]
                .dropna()
                .astype(str)
                .str.strip()
            )

            missing_keys = source_keys - target_keys
            failed_count = len(missing_keys)

        else:
            continue

        status = "PASS" if failed_count == 0 else "FAIL"

        print(f"{rule_name}: {status} ({failed_count} failures)")

        results.append({
            "Sheet Name": source_sheet,
            "Field": rule_name,
            "Test Type": "Cross Sheet Validation",
            "Test Name": "Cross Sheet Validation",
            "Status (P/F)": status,
            "Failed Count": failed_count
        })

    return results

def validate_trend(config, excel_file):

    print("\n====================================")
    print("Trend Validation")
    print("====================================")

    results = []

    rules = config.get("trendvalidation", [])

    for rule in rules:

        sheet_name = rule["sheet"]
        metric = rule["metric"]
        threshold = rule["threshold"]

        df = pd.read_excel(excel_file, sheet_name=sheet_name)

        values = pd.to_numeric(df["YOUR_METRIC_COLUMN"], errors="coerce")

        failed_count = (values > threshold).sum()

        status = "PASS" if failed_count == 0 else "FAIL"

        print(f"{metric}: {status} ({failed_count} failures)")

        results.append({
            "Sheet Name": sheet_name,
            "Field": metric,
            "Test Type": "Trend Validation",
            "Test Name": "Trend Validation",
            "Status (P/F)": status,
            "Failed Count": int(failed_count)
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
    reconciliation_results = validate_reconciliation(config, excel_file)
    cross_sheet_results = validate_cross_sheet(config, excel_file)
    trend_results = validate_trend(config, excel_file)

    final_results = sheet_results + column_results + duplicate_results + null_results + formula_results + numeric_precision_results + business_rule_results + reconciliation_results + cross_sheet_results + trend_results

    report = pd.DataFrame(final_results)

    report.to_excel("Validation_Report.xlsx", index=False)

    print("\nValidation_Report.xlsx created successfully.")


if __name__ == "__main__":
    main()