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
                "Test Type": "Field Validation",
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

    df.columns = (
        df.columns.str.strip()
        .str.replace(" ", "_")
        .str.replace("/", "_")
    )

    for rule in formula_rules:

        formula_name = rule["name"]
        formula = rule["formula"]
        validation_type = rule["validationtype"].lower()
        aggregate = rule.get("aggregate", False)

        try:

            if validation_type == "threshold":

                threshold = rule["threshold"]
                operator = rule["operator"]

                if aggregate:

                    group_columns = [
                        col.strip().replace(" ", "_")
                        for col in rule["groupbycolumns"]
                    ]

                    agg_dict = {}

                    for col in rule.get("sourcecolumns", []):
                        agg_dict[col.strip().replace(" ", "_")] = "sum"

                    grouped_df = (
                        df.groupby(
                            group_columns,
                            as_index=False
                        )
                        .agg(agg_dict)
                    )

                    calculated = grouped_df.eval(formula)

                else:

                    grouped_df = df.copy()

                    calculated = grouped_df.eval(formula)

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
                    raise ValueError(
                        f"Unsupported operator: {operator}"
                    )

                failed_count = int(failed_mask.sum())

            elif validation_type == "comparison":

                target_column = (
                    rule["targetcolumn"]
                    .strip()
                    .replace(" ", "_")
                )

                source_columns = [
                    col.strip().replace(" ", "_")
                    for col in rule["sourcecolumns"]
                ]

                if aggregate:

                    group_columns = [
                        col.strip().replace(" ", "_")
                        for col in rule["groupbycolumns"]
                    ]

                    agg_dict = {}

                    for col in source_columns:
                        agg_dict[col] = "sum"

                    agg_dict[target_column] = "sum"

                    grouped_df = (
                        df.groupby(
                            group_columns,
                            as_index=False
                        )
                        .agg(agg_dict)
                    )

                else:

                    grouped_df = df.copy()

                calculated = grouped_df.eval(formula)

                target_values = grouped_df[target_column]

                failed_mask = (
                    calculated.round(2)
                    !=
                    target_values.round(2)
                )

                failed_count = int(failed_mask.sum())

                if failed_count > 0:

                    print("\nFAILED ROWS:")

                    display_columns = []

                    if aggregate:
                        display_columns.extend(group_columns)

                    display_columns.extend(source_columns)
                    display_columns.append(target_column)

                    print(
                        grouped_df.loc[
                            failed_mask,
                            display_columns
                        ]
                    )

            else:

                raise ValueError(
                    f"Unsupported validation type: {validation_type}"
                )

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
            "Test Type": "Data Quality Validation",
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

                value = round(float(value), precision)

                value_str = str(value)

                decimal_part = value_str.split(".")

                if len(decimal_part) == 2:

                    decimal_places = len(
                        decimal_part[1].rstrip("0")
                    )

                    if decimal_places > precision:
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

        print("\nRule :", rule_name)
        print("Sheet:", sheet_name)
        print("Columns:")
        print(df.columns.tolist())

        if rule_name == "Latest Transaction Date should be less than Latest Loaded Date":

            left_column = rule["leftcolumn"]
            right_column = rule["rightcolumn"]
            distinct_column = rule["distinctcolumn"]

            failed_rows = []

            for _, row in df.iterrows():

                left_value = row[left_column]
                right_value = row[right_column]

                left_text = str(left_value).strip().lower()
                right_text = str(right_value).strip().lower()

                if (
                    left_text == "not applicable"
                    and
                    right_text == "not applicable"
                ):
                    failed_rows.append({
                        distinct_column:
                row[distinct_column],
                        left_column: left_value,
                        right_column: right_value
                    })
                    continue

                if (
                    (left_text == "not applicable" and right_text != "not applicable")
                    or
                    (right_text == "not applicable" and left_text != "not applicable")
                ):
                    failed_rows.append({
                        distinct_column: row[distinct_column],
                        left_column: left_value,
                        right_column: right_value
                    })
                    continue

                left_date = pd.to_datetime(left_value, errors="coerce")
                right_date = pd.to_datetime(right_value, errors="coerce")

                if pd.isna(left_date) or pd.isna(right_date):
                    failed_rows.append({
                        distinct_column: row[distinct_column],
                        left_column: left_value,
                        right_column: right_value
                    })
                    continue

                if left_date >= right_date:
                    failed_rows.append({
                        distinct_column: row[distinct_column],
                        left_column: left_value,
                        right_column: right_value
                    })

            failed_df = pd.DataFrame(
                failed_rows,
                columns=[
                    distinct_column,
                    left_column,
                    right_column
                ]
            )

            print("\nFAILED ROWS:")
            print(failed_df)

            failed_count = len(failed_df)

            if failed_count == 0:
                status = "P"
                print(f"{rule_name} - PASS")
            else:
                status = "F"
                print(f"{rule_name} - FAIL ({failed_count} Practice(s))")

            results.append({
                "Sheet Name": sheet_name,
                "Field": rule_name,
                "Test Type": "Data Validation",
                "Test Name": "Business Rule Validation",
                "Status (P/F)": status,
                "Failed Count": failed_count
            })

        elif rule_name == "Latest Transaction Date older than 3 months should display No LatestTransactionDate or NULL":

            column_name = rule["column"]
            reference_column = rule["referencecolumn"]
            months = rule["months"]

            df[column_name] = pd.to_datetime(df[column_name],errors="coerce")

            df[reference_column] = pd.to_datetime(df[reference_column],errors="coerce")

            threshold_date = df[reference_column] - pd.DateOffset(months=months)

            comparison = ((df[column_name] < threshold_date) & (~df[column_name].isna()))

            failed_rows = df.loc[comparison,[column_name, reference_column]]

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
                "Test Type": "Data Validation",
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
            "Test Type": "Data Validation",
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
        source_column = rule["sourcecolumn"]
        target_column = rule["targetcolumn"]
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
            col.strip().replace(" ", "_")
            for col in match_columns
        ]

        source_column = source_column.strip().replace(" ", "_")
        target_column = target_column.strip().replace(" ", "_")
        distinct_column = distinct_column.strip().replace(" ", "_")

        if source_column not in source_df.columns:
            print(f"Source column '{source_column}' not found.")
            continue

        if target_column not in target_df.columns:
            print(f"Target column '{target_column}' not found.")
            continue

        source_df[source_column] = pd.to_numeric(
            source_df[source_column],
            errors="coerce"
        ).fillna(0)

        target_df[target_column] = pd.to_numeric(
            target_df[target_column],
            errors="coerce"
        ).fillna(0)

        source_df = (
            source_df
            .groupby(match_columns, as_index=False)
            .agg({source_column: "sum"})
        )

        target_df = (
            target_df
            .groupby(match_columns, as_index=False)
            .agg({target_column: "sum"})
        )

        merged_df = pd.merge(
            source_df,
            target_df,
            on=match_columns,
            how="outer"
        ).fillna(0)

        print(f"\n{rule_name}")
        print(f"Matched Rows : {len(merged_df)}")

        print(f"\nComparing {source_column} with {target_column}")

        comparison = (merged_df[source_column] != merged_df[target_column])

        failed_rows = merged_df.loc[comparison].copy()

        if distinct_column in failed_rows.columns:
            failed_rows = failed_rows.drop_duplicates(
                subset=[distinct_column]
            )

        display_columns = list(dict.fromkeys(
            match_columns +
            [source_column, target_column]
        ))

        failed_rows = failed_rows[display_columns]

        failed_count = len(failed_rows)

        if failed_count == 0:
            status = "P"
            print("PASS")
        else:
            status = "F"
            print(f"FAIL ({failed_count} rows)")
            print(failed_rows)

        results.append({
            "Sheet Name": source_sheet,
            "Field": rule_name,
            "Test Type": "Data Validation",
            "Test Name": "Reconciliation Validation",
            "Status (P/F)": status,
            "Failed Count": failed_count
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

        status = "P" if failed_count == 0 else "F"

        print(f"{rule_name}: {status} ({failed_count} failures)")

        results.append({
            "Sheet Name": source_sheet,
            "Field": rule_name,
            "Test Type": "Data Validation",
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
        metric_name = rule["name"]
        metric = rule["metric"]
        previous_metric = rule["previousmetric"]

        operator = rule["operator"]
        threshold = rule["threshold"]

        aggregate = rule.get("aggregate", False)
        groupby_columns = rule.get("groupbycolumns", [])

        distinct_column = rule.get("distinctcolumn", "PracticeCode")

        df = pd.read_excel(excel_file, sheet_name=sheet_name)

        df.columns = (
            df.columns.str.strip()
            .str.replace(" ", "_")
            .str.replace("/", "_")
        )

        failed_count = 0

        if aggregate:

            df[metric] = pd.to_numeric(
                df[metric],
                errors="coerce"
            ).fillna(0)

            df = (
                df.groupby(
                    groupby_columns,
                    as_index=False
                ).agg({metric: "sum"})
            )

            df = df.sort_values(
                by=groupby_columns
            ).reset_index(drop=True)

            previous_values = (
                df.groupby(groupby_columns[0])[metric]
                .shift(1)
            )

            for index, (current, previous) in enumerate(
                zip(df[metric], previous_values)
            ):

                if pd.isna(previous):
                    continue

                if previous == 0:
                    continue

                deviation = ((current - previous) / previous) * 100

                if operator == ">":
                    condition = abs(deviation) > threshold

                elif operator == "<":
                    condition = abs(deviation) < threshold

                elif operator == ">=":
                    condition = abs(deviation) >= threshold

                elif operator == "<=":
                    condition = abs(deviation) <= threshold

                elif operator == "==":
                    condition = abs(deviation) == threshold

                else:
                    raise ValueError(
                        f"Unsupported operator: {operator}"
                    )
                if condition:

                    failed_count += 1

                    print(
                        f"{metric_name} -> "
                        f"{df.loc[index, 'PracticeCode']}"
                    )

        else:

            df = (
                df.groupby(
                    ["PracticeCode",
            "FiscalYear", "FiscalMonth"],
                    as_index=False
                )
                .agg({
                    metric: "sum",
                    previous_metric: "sum"
                })
            )

            current_values = pd.to_numeric(
                df[metric],
                errors="coerce"
            )

            previous_values = pd.to_numeric(
                df[previous_metric],
                errors="coerce"
            )

            for index, (current, previous) in enumerate(
                zip(current_values, previous_values)
            ):

                if pd.isna(current) or pd.isna(previous):
                    continue

                if previous == 0:
                    continue

                deviation = (
                    ((current - previous) / previous) * 100
                )

                condition = abs(deviation) > threshold

                status_text = "FAIL" if condition else "PASS"

                print(
                    df.loc[index, distinct_column],
                    current,
                    previous,
                    deviation,
                    threshold,
                    status_text
                )

                if condition:

                    failed_count += 1

                    print(
                        f"{metric_name} -> "
                        f"{df.loc[index, distinct_column]}"
                    )

        status = "P" if failed_count == 0 else "F"

        print(f"{metric_name}: {status} ({failed_count} failures)")

        results.append({
            "Sheet Name": sheet_name,
            "Field": metric_name,
            "Test Type": "Data Validation",
            "Test Name": "Trend Validation",
            "Status (P/F)": status,
            "Failed Count": int(failed_count)
        })

    print("\nTrend Validation Completed.")

    return results

def main():

    json_file = "Kpi_automation_phase1.json"
    excel_file = "KPI_Sample (2).xlsx"

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