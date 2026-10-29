from datetime import datetime



from pathlib import Path







import math



import pandas as pd



from openpyxl.styles import Alignment, Font, PatternFill



from openpyxl.utils import get_column_letter







from config_loader import load_config



from file_discovery import discover_input_files



from input_validator import validate_input_file



from telemetry_standardizer import standardize_telemetry_file











CORE_TELEMETRY_COLUMNS = [



    "Timestamp",



    "Latitude",



    "Longitude",



    "Speed(m/s)",



    "Speed(km/h)",



    "Address",



]







REMOVED_ROW_COLUMNS = [



    "File Name",



    "Source File",



    "Source Schema",



    "Source Sheet",



    "Source Excel Row",



    "Removal Stage",



    "Removal Reason",



    *CORE_TELEMETRY_COLUMNS,



]







SPEED_AUDIT_COLUMNS = [



    "File Name",



    "Source File",



    "Source Schema",



    "Source Sheet",



    "Source Excel Row",



    "Timestamp",



    "Latitude",



    "Longitude",



    "Original Speed(m/s)",



    "Original Speed(km/h)",



    "Address",



    "Speed(m/s)",



    "Speed(km/h)",



    "Speed Recovery Status",



]







FILE_SUMMARY_COLUMNS = [



    "file_name",



    "source_file",



    "source_schema",



    "source_sheet",



    "processing_status",



    "original_rows",



    "invalid_speeds_recovered",



    "invalid_speeds_unresolved",



    "unrecoverable_rows_removed",



    "duplicate_rows_removed",



    "total_rows_removed",



    "cleaned_rows",



    "notes",



]















def _clean_text(value) -> str:



    if value is None:



        return ""



    try:



        if pd.isna(value):



            return ""



    except (TypeError, ValueError):



        pass



    return str(value).strip()











def _normalise_key(value) -> str:



    text = _clean_text(value).casefold()



    return "".join(character for character in text if character.isalnum())











def _read_raw_address_by_excel_row(file_path: Path) -> dict[int, str]:



    """Read the raw Address column and return {Excel row number: Address}."""







    with pd.ExcelFile(file_path, engine="openpyxl") as excel:



        if not excel.sheet_names:



            raise ValueError(f"No worksheet found in {file_path}")



        raw = pd.read_excel(



            excel,



            sheet_name=excel.sheet_names[0],



        )







    lookup = {



        _normalise_key(column): column



        for column in raw.columns



    }







    address_column = None



    for alias in (



        "Address",



        "Location Address",



        "Full Address",



        "Formatted Address",



    ):



        actual = lookup.get(_normalise_key(alias))



        if actual is not None:



            address_column = actual



            break







    if address_column is None:



        # Keep the pipeline compatible with source files that genuinely do not



        # contain Address. Those rows receive a blank Address.



        return {}







    addresses = (



        raw[address_column]



        .astype("string")



        .fillna("")



        .str.strip()



    )







    # pandas row 0 corresponds to Excel row 2 because row 1 is the header.



    return {



        excel_row: _clean_text(address)



        for excel_row, address in enumerate(addresses, start=2)



    }















def _haversine_metres(latitude_1, longitude_1, latitude_2, longitude_2):



    """Return great-circle distance in metres between two GPS coordinates."""



    earth_radius_m = 6_371_008.8



    lat_1 = math.radians(float(latitude_1))



    lat_2 = math.radians(float(latitude_2))



    delta_lat = lat_2 - lat_1



    delta_lon = math.radians(float(longitude_2) - float(longitude_1))



    haversine_value = (



        math.sin(delta_lat / 2.0) ** 2



        + math.cos(lat_1)



        * math.cos(lat_2)



        * math.sin(delta_lon / 2.0) ** 2



    )



    haversine_value = min(1.0, max(0.0, haversine_value))



    return (



        earth_radius_m



        * 2.0



        * math.atan2(



            math.sqrt(haversine_value),



            math.sqrt(1.0 - haversine_value),



        )



    )






def repair_isolated_gps_spikes(dataframe, config):



    """



    Repair isolated impossible GPS teleports by time-based interpolation.



    A point is corrected only when both adjacent GPS intervals imply an



    impossible speed, while the direct previous-to-next movement remains



    physically plausible. Rows are retained; only Latitude and Longitude are



    replaced for confirmed isolated spikes.



    """



    dataframe = dataframe.copy()



    if len(dataframe) < 3:



        return dataframe



    cleaning = config.get("cleaning", {})



    minimum_jump_m = float(



        cleaning.get("gps_spike_min_jump_m", 500.0)



    )



    impossible_speed_kmh = float(



        cleaning.get("gps_spike_max_implied_speed_kmh", 200.0)



    )



    maximum_bridge_speed_kmh = float(



        cleaning.get("gps_spike_max_bridge_speed_kmh", 160.0)



    )



    maximum_surrounding_gap_seconds = float(



        cleaning.get("gps_spike_max_surrounding_gap_seconds", 60.0)



    )



    for position in range(1, len(dataframe) - 1):



        previous = dataframe.iloc[position - 1]



        current = dataframe.iloc[position]



        following = dataframe.iloc[position + 1]



        values = [



            previous["Timestamp"],



            current["Timestamp"],



            following["Timestamp"],



            previous["Latitude"],



            previous["Longitude"],



            current["Latitude"],



            current["Longitude"],



            following["Latitude"],



            following["Longitude"],



        ]



        if any(pd.isna(value) for value in values):



            continue



        seconds_before = (



            current["Timestamp"] - previous["Timestamp"]



        ).total_seconds()



        seconds_after = (



            following["Timestamp"] - current["Timestamp"]



        ).total_seconds()



        surrounding_seconds = seconds_before + seconds_after



        if (



            seconds_before <= 0



            or seconds_after <= 0



            or surrounding_seconds > maximum_surrounding_gap_seconds



        ):



            continue



        distance_before_m = _haversine_metres(



            previous["Latitude"],



            previous["Longitude"],



            current["Latitude"],



            current["Longitude"],



        )



        distance_after_m = _haversine_metres(



            current["Latitude"],



            current["Longitude"],



            following["Latitude"],



            following["Longitude"],



        )



        bridge_distance_m = _haversine_metres(



            previous["Latitude"],



            previous["Longitude"],



            following["Latitude"],



            following["Longitude"],



        )



        speed_before_kmh = (



            distance_before_m / seconds_before * 3.6



        )



        speed_after_kmh = (



            distance_after_m / seconds_after * 3.6



        )



        bridge_speed_kmh = (



            bridge_distance_m / surrounding_seconds * 3.6



        )



        isolated_spike = (



            distance_before_m >= minimum_jump_m



            and distance_after_m >= minimum_jump_m



            and speed_before_kmh > impossible_speed_kmh



            and speed_after_kmh > impossible_speed_kmh



            and bridge_speed_kmh <= maximum_bridge_speed_kmh



            and (distance_before_m + distance_after_m)



            >= 4.0 * max(bridge_distance_m, 1.0)



        )



        if not isolated_spike:



            continue



        interpolation_fraction = (



            seconds_before / surrounding_seconds



        )



        row_label = dataframe.index[position]



        dataframe.at[row_label, "Latitude"] = (



            float(previous["Latitude"])



            + interpolation_fraction



            * (float(following["Latitude"]) - float(previous["Latitude"]))



        )



        dataframe.at[row_label, "Longitude"] = (



            float(previous["Longitude"])



            + interpolation_fraction



            * (float(following["Longitude"]) - float(previous["Longitude"]))



        )



    return dataframe






def recover_invalid_speeds(dataframe, max_gap_seconds):



    """



    Recover negative or missing speeds by timestamp-based interpolation.







    An invalid block is recovered only when it has a valid observation on



    both sides and the full time gap between those observations is no more



    than max_gap_seconds. Unresolved speed rows are retained and flagged.



    """







    dataframe = dataframe.copy()







    dataframe["Original Speed(m/s)"] = dataframe["Speed(m/s)"]



    dataframe["Original Speed(km/h)"] = dataframe["Speed(km/h)"]



    dataframe["Speed Recovery Status"] = "ORIGINAL_VALID"







    invalid_speed_mask = (



        dataframe["Speed(m/s)"].isna()



        | dataframe["Speed(km/h)"].isna()



        | (dataframe["Speed(m/s)"] < 0)



        | (dataframe["Speed(km/h)"] < 0)



    )







    # Use floating NaN to preserve numeric column dtypes.



    dataframe.loc[



        invalid_speed_mask,



        ["Speed(m/s)", "Speed(km/h)"],



    ] = float("nan")







    dataframe.loc[



        invalid_speed_mask,



        "Speed Recovery Status",



    ] = "UNRESOLVED"







    invalid_positions = [



        position



        for position, is_invalid in enumerate(invalid_speed_mask)



        if is_invalid



    ]







    recovered_count = 0



    position_index = 0







    while position_index < len(invalid_positions):



        block_start = invalid_positions[position_index]



        block_end = block_start







        while (



            position_index + 1 < len(invalid_positions)



            and invalid_positions[position_index + 1] == block_end + 1



        ):



            position_index += 1



            block_end = invalid_positions[position_index]







        previous_position = block_start - 1



        next_position = block_end + 1







        previous_is_valid = (



            previous_position >= 0



            and pd.notna(



                dataframe.iloc[previous_position]["Speed(km/h)"]



            )



        )







        next_is_valid = (



            next_position < len(dataframe)



            and pd.notna(dataframe.iloc[next_position]["Speed(km/h)"])



        )







        if previous_is_valid and next_is_valid:



            previous_timestamp = dataframe.iloc[previous_position][



                "Timestamp"



            ]



            next_timestamp = dataframe.iloc[next_position]["Timestamp"]







            if pd.notna(previous_timestamp) and pd.notna(next_timestamp):



                total_gap_seconds = (



                    next_timestamp - previous_timestamp



                ).total_seconds()







                if 0 < total_gap_seconds <= max_gap_seconds:



                    previous_speed = float(



                        dataframe.iloc[previous_position]["Speed(km/h)"]



                    )



                    next_speed = float(



                        dataframe.iloc[next_position]["Speed(km/h)"]



                    )







                    for row_position in range(



                        block_start,



                        block_end + 1,



                    ):



                        current_timestamp = dataframe.iloc[row_position][



                            "Timestamp"



                        ]







                        if pd.isna(current_timestamp):



                            continue







                        elapsed_seconds = (



                            current_timestamp - previous_timestamp



                        ).total_seconds()



                        interpolation_fraction = (



                            elapsed_seconds / total_gap_seconds



                        )



                        recovered_speed_kmh = (



                            previous_speed



                            + interpolation_fraction



                            * (next_speed - previous_speed)



                        )



                        recovered_speed_kmh = max(



                            0.0,



                            recovered_speed_kmh,



                        )







                        row_label = dataframe.index[row_position]



                        dataframe.at[



                            row_label,



                            "Speed(km/h)",



                        ] = recovered_speed_kmh



                        dataframe.at[



                            row_label,



                            "Speed(m/s)",



                        ] = recovered_speed_kmh / 3.6



                        dataframe.at[



                            row_label,



                            "Speed Recovery Status",



                        ] = "TIME_INTERPOLATED"







                        recovered_count += 1







        position_index += 1







    unresolved_count = int(



        (dataframe["Speed Recovery Status"] == "UNRESOLVED").sum()



    )







    return dataframe, recovered_count, unresolved_count











def _collect_removal_reasons(dataframe, config):



    """Return core-invalid mask and a reason for every invalid source row."""







    latitude_min = config["cleaning"]["latitude_min"]



    latitude_max = config["cleaning"]["latitude_max"]



    longitude_min = config["cleaning"]["longitude_min"]



    longitude_max = config["cleaning"]["longitude_max"]







    checks = [



        (dataframe["Timestamp"].isna(), "Missing or invalid timestamp"),



        (dataframe["Latitude"].isna(), "Missing or non-numeric latitude"),



        (



            dataframe["Latitude"].notna()



            & ~dataframe["Latitude"].between(



                latitude_min,



                latitude_max,



            ),



            (



                "Latitude outside allowed range "



                f"[{latitude_min}, {latitude_max}]"



            ),



        ),



        (



            dataframe["Longitude"].isna(),



            "Missing or non-numeric longitude",



        ),



        (



            dataframe["Longitude"].notna()



            & ~dataframe["Longitude"].between(



                longitude_min,



                longitude_max,



            ),



            (



                "Longitude outside allowed range "



                f"[{longitude_min}, {longitude_max}]"



            ),



        ),



    ]







    reasons_by_index = {index: [] for index in dataframe.index}







    for mask, reason in checks:



        for index in dataframe.index[mask]:



            reasons_by_index[index].append(reason)







    invalid_mask = pd.Series(



        [bool(reasons_by_index[index]) for index in dataframe.index],



        index=dataframe.index,



        dtype=bool,



    )







    reasons = pd.Series(



        ["; ".join(reasons_by_index[index]) for index in dataframe.index],



        index=dataframe.index,



        dtype="string",



    )







    return invalid_mask, reasons











def _make_removed_rows(



    dataframe,



    mask,



    file_path,



    metadata,



    stage,



    reasons,



):



    """Create a report-ready table for rows removed at one stage."""







    if not bool(mask.any()):



        return pd.DataFrame(columns=REMOVED_ROW_COLUMNS)







    removed = dataframe.loc[



        mask,



        ["Source Excel Row", *CORE_TELEMETRY_COLUMNS],



    ].copy()







    if isinstance(reasons, str):



        removed["Removal Reason"] = reasons



    else:



        removed["Removal Reason"] = reasons.loc[mask].values







    removed["Removal Stage"] = stage



    removed["File Name"] = file_path.name



    removed["Source File"] = str(file_path)



    removed["Source Schema"] = metadata["source_schema"]



    removed["Source Sheet"] = metadata["sheet_name"]







    return removed[REMOVED_ROW_COLUMNS]











def _add_speed_audit_context(dataframe, file_path, metadata):



    """Return every recovered or unresolved speed row for the audit report."""







    audit_mask = dataframe["Speed Recovery Status"] != "ORIGINAL_VALID"







    if not bool(audit_mask.any()):



        return pd.DataFrame(columns=SPEED_AUDIT_COLUMNS)







    speed_audit = dataframe.loc[



        audit_mask,



        [



            "Source Excel Row",



            "Timestamp",



            "Latitude",



            "Longitude",



            "Original Speed(m/s)",



            "Original Speed(km/h)",



            "Address",



            "Speed(m/s)",



            "Speed(km/h)",



            "Speed Recovery Status",



        ],



    ].copy()







    speed_audit["File Name"] = file_path.name



    speed_audit["Source File"] = str(file_path)



    speed_audit["Source Schema"] = metadata["source_schema"]



    speed_audit["Source Sheet"] = metadata["sheet_name"]







    return speed_audit[SPEED_AUDIT_COLUMNS]











def clean_telemetry_data_with_audit(file_path, config):



    """



    Clean one standardized telemetry file and return full audit information.







    The original workbook is opened read-only by the pipeline and is never



    overwritten. Source Excel Row assumes the worksheet header is in row 1,



    so the first data record is Excel row 2.



    """







    file_path = Path(file_path)



    dataframe, metadata = standardize_telemetry_file(file_path)







    original_row_count = len(dataframe)



    dataframe = dataframe.copy()



    dataframe["Source Excel Row"] = range(2, len(dataframe) + 2)







    # Preserve the Address from the original input workbook.



    address_by_excel_row = _read_raw_address_by_excel_row(file_path)



    dataframe["Address"] = dataframe["Source Excel Row"].map(



        lambda row_number: address_by_excel_row.get(



            int(row_number),



            "",



        )



    )







    core_invalid_mask, core_removal_reasons = (



        _collect_removal_reasons(dataframe, config)



    )







    core_removed_rows = _make_removed_rows(



        dataframe=dataframe,



        mask=core_invalid_mask,



        file_path=file_path,



        metadata=metadata,



        stage="CORE_VALIDATION",



        reasons=core_removal_reasons,



    )







    unrecoverable_rows_removed = int(core_invalid_mask.sum())



    dataframe = dataframe.loc[~core_invalid_mask].copy()







    duplicate_rows_removed = 0



    duplicate_removed_rows = pd.DataFrame(



        columns=REMOVED_ROW_COLUMNS



    )







    if config["cleaning"]["remove_exact_duplicates"]:



        duplicate_mask = dataframe.duplicated(



            subset=CORE_TELEMETRY_COLUMNS,



            keep="first",



        )



        duplicate_rows_removed = int(duplicate_mask.sum())







        duplicate_removed_rows = _make_removed_rows(



            dataframe=dataframe,



            mask=duplicate_mask,



            file_path=file_path,



            metadata=metadata,



            stage="DUPLICATE_REMOVAL",



            reasons=(



                "Exact duplicate of an earlier standardized telemetry "



                "row; first occurrence retained"



            ),



        )







        dataframe = dataframe.loc[~duplicate_mask].copy()







    if config["cleaning"]["sort_by_timestamp"]:



        dataframe = dataframe.sort_values(



            by="Timestamp",



            kind="stable",



        )







    dataframe = dataframe.reset_index(drop=True)







    dataframe = repair_isolated_gps_spikes(



        dataframe,



        config,



    )







    max_gap_seconds = config["cleaning"].get(



        "speed_interpolation_max_gap_seconds",



        60,



    )







    (



        dataframe,



        invalid_speeds_recovered,



        invalid_speeds_unresolved,



    ) = recover_invalid_speeds(



        dataframe,



        max_gap_seconds,



    )







    removed_row_frames = [

        frame.dropna(axis=1, how="all")

        for frame in (core_removed_rows, duplicate_removed_rows)

        if frame is not None and not frame.empty

    ]



    if removed_row_frames:

        removed_rows = pd.concat(

            removed_row_frames,

            ignore_index=True,

        ).reindex(columns=REMOVED_ROW_COLUMNS)

    else:

        removed_rows = pd.DataFrame(

            columns=REMOVED_ROW_COLUMNS

        )







    speed_audit = _add_speed_audit_context(



        dataframe,



        file_path,



        metadata,



    )







    total_rows_removed = (



        unrecoverable_rows_removed + duplicate_rows_removed



    )







    quality_summary = {



        "file_name": file_path.name,



        "source_file": str(file_path),



        "source_schema": metadata["source_schema"],



        "source_sheet": metadata["sheet_name"],



        "processing_status": "CLEANED",



        "original_rows": original_row_count,



        "invalid_speeds_recovered": invalid_speeds_recovered,



        "invalid_speeds_unresolved": invalid_speeds_unresolved,



        "unrecoverable_rows_removed": unrecoverable_rows_removed,



        "duplicate_rows_removed": duplicate_rows_removed,



        "total_rows_removed": total_rows_removed,



        "cleaned_rows": len(dataframe),



        "notes": "",



    }







    return dataframe, quality_summary, removed_rows, speed_audit











def clean_telemetry_data(file_path, config):



    """



    Backward-compatible cleaner used by later pipeline modules.







    It intentionally returns the same two values as the earlier version.



    """







    dataframe, summary, _, _ = clean_telemetry_data_with_audit(



        file_path,



        config,



    )







    return dataframe, summary











def build_cleaning_rules(config):



    """Build the exact, human-readable rules recorded in the workbook."""







    cleaning = config["cleaning"]



    interpolation_gap = cleaning.get(



        "speed_interpolation_max_gap_seconds",



        60,



    )







    return pd.DataFrame(



        [



            {



                "Rule ID": "R01",



                "Stage": "Standardization",



                "Condition": "Supported source schema is detected",



                "Action": (



                    "Convert source fields to Timestamp, Latitude, "



                    "Longitude, Speed(m/s), and Speed(km/h)"



                ),



                "Removes Row": "No",



                "Setting / Threshold": (



                    "DEVICE_RAW, MOBILE_RAW, or MOBILE_PROCESSED"



                ),



            },



            {



                "Rule ID": "R02",



                "Stage": "Core validation",



                "Condition": "Timestamp is missing or cannot be parsed",



                "Action": "Remove row and record the reason",



                "Removes Row": "Yes",



                "Setting / Threshold": "Valid date-time required",



            },



            {



                "Rule ID": "R03",



                "Stage": "Core validation",



                "Condition": "Latitude is missing, non-numeric, or outside range",



                "Action": "Remove row and record the reason",



                "Removes Row": "Yes",



                "Setting / Threshold": (



                    f"{cleaning['latitude_min']} to "



                    f"{cleaning['latitude_max']} degrees"



                ),



            },



            {



                "Rule ID": "R04",



                "Stage": "Core validation",



                "Condition": "Longitude is missing, non-numeric, or outside range",



                "Action": "Remove row and record the reason",



                "Removes Row": "Yes",



                "Setting / Threshold": (



                    f"{cleaning['longitude_min']} to "



                    f"{cleaning['longitude_max']} degrees"



                ),



            },



            {



                "Rule ID": "R05",



                "Stage": "Duplicate removal",



                "Condition": (



                    "Timestamp, latitude, longitude, Speed(m/s), and "



                    "Speed(km/h) exactly match an earlier row"



                ),



                "Action": "Keep first occurrence; remove later occurrence",



                "Removes Row": (



                    "Yes" if cleaning["remove_exact_duplicates"] else "Disabled"



                ),



                "Setting / Threshold": (



                    f"remove_exact_duplicates = "



                    f"{cleaning['remove_exact_duplicates']}"



                ),



            },



            {



                "Rule ID": "R06",



                "Stage": "Ordering",



                "Condition": "Core validation and duplicate handling completed",



                "Action": (



                    "Sort records by Timestamp using a stable sort"



                    if cleaning["sort_by_timestamp"]



                    else "Keep source order"



                ),



                "Removes Row": "No",



                "Setting / Threshold": (



                    f"sort_by_timestamp = "



                    f"{cleaning['sort_by_timestamp']}"



                ),



            },



            {



                "Rule ID": "R07",



                "Stage": "Speed recovery",



                "Condition": "Speed is missing or negative",



                "Action": (



                    "Treat speed as unavailable; do not remove the row"



                ),



                "Removes Row": "No",



                "Setting / Threshold": "Speed < 0 or missing",



            },



            {



                "Rule ID": "R08",



                "Stage": "Speed recovery",



                "Condition": (



                    "Invalid speed block has valid speeds before and "



                    "after it, with a positive surrounding time gap "



                    "within the configured limit"



                ),



                "Action": (



                    "Linearly interpolate speed according to elapsed time"



                ),



                "Removes Row": "No",



                "Setting / Threshold": (



                    f"Maximum surrounding gap = {interpolation_gap} seconds"



                ),



            },



            {



                "Rule ID": "R09",



                "Stage": "Speed recovery",



                "Condition": "Invalid speed does not satisfy Rule R08",



                "Action": "Retain row and flag speed as UNRESOLVED",



                "Removes Row": "No",



                "Setting / Threshold": "No automatic replacement",



            },



            {



                "Rule ID": "R10",



                "Stage": "File acceptance",



                "Condition": "Rows remaining after cleaning are below minimum",



                "Action": "Mark the file as skipped",



                "Removes Row": "No additional row removal",



                "Setting / Threshold": (



                    f"minimum_valid_rows = "



                    f"{cleaning['minimum_valid_rows']}"



                ),



            },



            {



                "Rule ID": "R11",



                "Stage": "Traceability",



                "Condition": "A row is reported",



                "Action": (



                    "Record source workbook, source sheet, and original "



                    "Excel row number; assumes headers are in Excel row 1"



                ),



                "Removes Row": "No",



                "Setting / Threshold": "First data record = Excel row 2",



            },



            {



                "Rule ID": "R12",



                "Stage": "Data protection",



                "Condition": "Any cleaning or audit operation runs",



                "Action": (



                    "Read source files only; write the audit report under "



                    "the configured Results folder"



                ),



                "Removes Row": "No",



                "Setting / Threshold": "Raw workbooks are never overwritten",



            },



        ]



    )











def _format_audit_workbook(writer):



    """Apply readable formatting to every audit worksheet."""







    header_fill = PatternFill(



        fill_type="solid",



        fgColor="1F4E78",



    )



    header_font = Font(color="FFFFFF", bold=True)







    for worksheet in writer.book.worksheets:



        worksheet.sheet_view.showGridLines = False







        if worksheet.title != "Run Summary":



            worksheet.freeze_panes = "A2"



            if worksheet.max_row >= 1 and worksheet.max_column >= 1:



                worksheet.auto_filter.ref = worksheet.dimensions







        for cell in worksheet[1]:



            cell.fill = header_fill



            cell.font = header_font



            cell.alignment = Alignment(



                horizontal="center",



                vertical="center",



                wrap_text=True,



            )







        worksheet.row_dimensions[1].height = 32







        for column_number in range(1, worksheet.max_column + 1):



            column_letter = get_column_letter(column_number)



            values = [



                str(worksheet.cell(row=row, column=column_number).value or "")



                for row in range(1, min(worksheet.max_row, 250) + 1)



            ]



            longest = max((len(value) for value in values), default=0)



            width = min(max(longest + 2, 12), 45)



            worksheet.column_dimensions[column_letter].width = width







        for row in worksheet.iter_rows(min_row=2):



            for cell in row:



                if isinstance(cell.value, datetime):



                    cell.number_format = "yyyy-mm-dd hh:mm:ss"



                cell.alignment = Alignment(



                    vertical="top",



                    wrap_text=(worksheet.title == "Cleaning Rules"),



                )







    if "Cleaning Rules" in writer.book.sheetnames:



        rules_sheet = writer.book["Cleaning Rules"]



        rules_sheet.column_dimensions["C"].width = 48



        rules_sheet.column_dimensions["D"].width = 48



        rules_sheet.column_dimensions["F"].width = 38







    if "Removed Rows" in writer.book.sheetnames:



        removed_sheet = writer.book["Removed Rows"]



        removed_sheet.column_dimensions["B"].width = 45



        removed_sheet.column_dimensions["G"].width = 55







    if "File Summary" in writer.book.sheetnames:



        file_sheet = writer.book["File Summary"]



        file_sheet.column_dimensions["B"].width = 45



        file_sheet.column_dimensions["M"].width = 55











def write_cleaning_audit_report(



    config,



    file_summaries,



    removed_rows,



    speed_audit_rows,



    files_discovered,



    cleaned_count,



    skipped_count,



    failed_count,



):



    """Write the complete cleaning audit workbook under Results."""







    results_folder = Path(config["project"]["results_folder"])



    results_folder.mkdir(parents=True, exist_ok=True)







    report_name = config.get("outputs", {}).get(



        "cleaning_audit_report_name",



        "Data_Cleaning_Audit.xlsx",



    )



    report_path = results_folder / report_name







    file_summary_dataframe = pd.DataFrame(



        file_summaries,



        columns=FILE_SUMMARY_COLUMNS,



    )







    removed_row_frames = [

        frame.dropna(axis=1, how="all")

        for frame in removed_rows

        if frame is not None and not frame.empty

    ]



    if removed_row_frames:

        removed_rows_dataframe = pd.concat(

            removed_row_frames,

            ignore_index=True,

        ).reindex(columns=REMOVED_ROW_COLUMNS)

    else:

        removed_rows_dataframe = pd.DataFrame(

            columns=REMOVED_ROW_COLUMNS

        )







    speed_audit_frames = [

        frame.dropna(axis=1, how="all")

        for frame in speed_audit_rows

        if frame is not None and not frame.empty

    ]



    if speed_audit_frames:

        speed_audit_dataframe = pd.concat(

            speed_audit_frames,

            ignore_index=True,

        ).reindex(columns=SPEED_AUDIT_COLUMNS)

    else:

        speed_audit_dataframe = pd.DataFrame(

            columns=SPEED_AUDIT_COLUMNS

        )







    numeric_summary_columns = [



        "original_rows",



        "invalid_speeds_recovered",



        "invalid_speeds_unresolved",



        "unrecoverable_rows_removed",



        "duplicate_rows_removed",



        "total_rows_removed",



        "cleaned_rows",



    ]







    totals = {



        column: int(



            pd.to_numeric(



                file_summary_dataframe[column],



                errors="coerce",



            )



            .fillna(0)



            .sum()



        )



        for column in numeric_summary_columns



    }







    run_summary = pd.DataFrame(



        [



            ("Generated at", datetime.now()),



            ("Input folder", config["project"]["input_folder"]),



            ("Audit report", str(report_path)),



            ("Files discovered", files_discovered),



            ("Files cleaned successfully", cleaned_count),



            ("Files skipped", skipped_count),



            ("Files failed", failed_count),



            ("Original rows", totals["original_rows"]),



            (



                "Core-invalid rows removed",



                totals["unrecoverable_rows_removed"],



            ),



            (



                "Exact duplicate rows removed",



                totals["duplicate_rows_removed"],



            ),



            ("Total rows removed", totals["total_rows_removed"]),



            (



                "Invalid speeds recovered",



                totals["invalid_speeds_recovered"],



            ),



            (



                "Invalid speeds unresolved",



                totals["invalid_speeds_unresolved"],



            ),



            ("Final retained rows", totals["cleaned_rows"]),



        ],



        columns=["Metric", "Value"],



    )







    cleaning_rules = build_cleaning_rules(config)







    with pd.ExcelWriter(



        report_path,



        engine="openpyxl",



    ) as writer:



        run_summary.to_excel(



            writer,



            sheet_name="Run Summary",



            index=False,



        )



        file_summary_dataframe.to_excel(



            writer,



            sheet_name="File Summary",



            index=False,



        )



        removed_rows_dataframe.to_excel(



            writer,



            sheet_name="Removed Rows",



            index=False,



        )



        speed_audit_dataframe.to_excel(



            writer,



            sheet_name="Speed Audit",



            index=False,



        )



        cleaning_rules.to_excel(



            writer,



            sheet_name="Cleaning Rules",



            index=False,



        )







        _format_audit_workbook(writer)







    return report_path











def _validation_failure_summary(item, validation):



    """Create one File Summary record for a validator-rejected file."""







    return {



        "file_name": item["file_name"],



        "source_file": str(item["file_path"]),



        "source_schema": validation.get("source_schema"),



        "source_sheet": None,



        "processing_status": "SKIPPED_VALIDATION",



        "original_rows": validation.get("row_count", 0),



        "invalid_speeds_recovered": 0,



        "invalid_speeds_unresolved": validation.get(



            "invalid_speeds",



            0,



        ),



        "unrecoverable_rows_removed": 0,



        "duplicate_rows_removed": 0,



        "total_rows_removed": 0,



        "cleaned_rows": 0,



        "notes": "; ".join(validation.get("issues", [])),



    }











if __name__ == "__main__":



    settings = load_config()



    discovered_files = discover_input_files(settings)







    print("Data cleaning started")



    print(f"Files discovered: {len(discovered_files)}")







    cleaned_file_count = 0



    skipped_file_count = 0



    failed_file_count = 0







    file_summaries = []



    all_removed_rows = []



    all_speed_audit_rows = []







    for item in discovered_files:



        try:



            validation = validate_input_file(



                item["file_path"],



                settings,



            )







            if validation["status"] != "VALID":



                skipped_file_count += 1



                file_summaries.append(



                    _validation_failure_summary(item, validation)



                )







                print(



                    f"SKIPPED | {item['file_name']} | "



                    "Input validation failed"



                )



                continue







            (



                cleaned_data,



                summary,



                removed_rows,



                speed_audit,



            ) = clean_telemetry_data_with_audit(



                item["file_path"],



                settings,



            )







            minimum_valid_rows = settings["cleaning"][



                "minimum_valid_rows"



            ]







            if len(cleaned_data) < minimum_valid_rows:



                skipped_file_count += 1



                summary["processing_status"] = "SKIPPED_TOO_FEW_ROWS"



                summary["notes"] = (



                    "Insufficient valid rows after cleaning: "



                    f"{len(cleaned_data)} < {minimum_valid_rows}"



                )



                file_summaries.append(summary)



                all_removed_rows.append(removed_rows)



                all_speed_audit_rows.append(speed_audit)







                print(



                    f"SKIPPED | {summary['file_name']} | "



                    "Insufficient valid rows after cleaning"



                )



                continue







            cleaned_file_count += 1



            file_summaries.append(summary)



            all_removed_rows.append(removed_rows)



            all_speed_audit_rows.append(speed_audit)







            print(



                f"CLEANED | {summary['file_name']} | "



                f"Schema: {summary['source_schema']} | "



                f"Original: {summary['original_rows']} | "



                f"Speeds recovered: "



                f"{summary['invalid_speeds_recovered']} | "



                f"Unresolved speeds: "



                f"{summary['invalid_speeds_unresolved']} | "



                f"Core rows removed: "



                f"{summary['unrecoverable_rows_removed']} | "



                f"Duplicates removed: "



                f"{summary['duplicate_rows_removed']} | "



                f"Final rows: {summary['cleaned_rows']}"



            )







        except Exception as error:



            failed_file_count += 1



            file_summaries.append(



                {



                    "file_name": item["file_name"],



                    "source_file": str(item["file_path"]),



                    "source_schema": None,



                    "source_sheet": None,



                    "processing_status": "FAILED",



                    "original_rows": 0,



                    "invalid_speeds_recovered": 0,



                    "invalid_speeds_unresolved": 0,



                    "unrecoverable_rows_removed": 0,



                    "duplicate_rows_removed": 0,



                    "total_rows_removed": 0,



                    "cleaned_rows": 0,



                    "notes": str(error),



                }



            )







            print(f"FAILED | {item['file_name']} | {error}")







    print(f"Files cleaned successfully: {cleaned_file_count}")



    print(f"Files skipped: {skipped_file_count}")



    print(f"Files failed: {failed_file_count}")







    try:



        audit_report_path = write_cleaning_audit_report(



            config=settings,



            file_summaries=file_summaries,



            removed_rows=all_removed_rows,



            speed_audit_rows=all_speed_audit_rows,



            files_discovered=len(discovered_files),



            cleaned_count=cleaned_file_count,



            skipped_count=skipped_file_count,



            failed_count=failed_file_count,



        )







        print(f"Cleaning audit saved: {audit_report_path}")







    except PermissionError:



        print(



            "AUDIT REPORT NOT SAVED | Close "



            "Data_Cleaning_Audit.xlsx if it is open in Excel, "



            "then run the cleaner again."



        )



    except Exception as error:



        print(f"AUDIT REPORT FAILED | {error}")