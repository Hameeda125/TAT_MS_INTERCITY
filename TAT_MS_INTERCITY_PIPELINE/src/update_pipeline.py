"""



TAT COMPLETE PIPELINE UPDATER



Version: 2026-09-13 INTERCITY_9_STAGE_V6







IMPORTANT



\\---------



This is the complete updater, not the old cleaning/index-only updater.







Save this file as:



    D:/Hameeda_Projects/TAT_MS_INTERCITY/TAT_MS_INTERCITY_PIPELINE/src/update_pipeline.py







Normal run:



    python src/update_pipeline.py







Check stage files only:



    python src/update_pipeline.py --check-only







Optional diagnostic run that attempts later stages even if one fails:



    python src/update_pipeline.py --continue-on-error







DEFAULT PIPELINE ORDER



\\----------------------



Stage 1/9  Cleaning + permanent trip index



Stage 2/9  Manual trip-log reconciliation



Stage 3/9  Candidate EV-ICE pairing



Stage 4/9  Pair validation + common-section cropping



Stage 5/9  Post-cropping feature engineering



Stage 6/9  Phase segmentation



Stage 7/9  TAT computation



Stage 8/9  Results visualization



Stage 9/9  Master INTERCITY summary + ML training datasets







EXPLICITLY EXCLUDED



\\-------------------



reverse_geocoding.py



plot_unique_routes_and_update_summary.py







Stage 1 internally uses the existing helper modules through:



    cleaned_file_writer.write_all_cleaned_files()



    trip_index.update_trip_index()







The normal/safe behavior is to stop if an upstream stage fails.



Use --continue-on-error only for diagnosis, because downstream outputs may



not be valid when an earlier stage failed.



"""







from __future__ import annotations







import argparse



import os



import subprocess



import sys



import time



import traceback



from datetime import datetime



from pathlib import Path



from typing import Any







from config_loader import load_config



from cleaned_file_writer import write_all_cleaned_files



from trip_index import update_trip_index











# =============================================================================



# VERSION MARKER



# =============================================================================







VERSION = "2026-09-29_INTERCITY_9_STAGE_V7_PATHS_STRICT_GUARD"











# =============================================================================



# PROJECT PATHS



# =============================================================================







PROJECT_ROOT = Path(__file__).resolve().parent.parent



SRC_DIR = PROJECT_ROOT / "src"







RESULTS_ROOT = PROJECT_ROOT.parent / "Data" / "Results"



SYSTEM_DIR = RESULTS_ROOT / "System"







TOTAL_STAGES = 9



RUN_VISUALIZATION = True











# =============================================================================



# ACTUAL STANDALONE PIPELINE STAGES



# =============================================================================







STAGE_SCRIPTS: list[tuple[int, str, Path]] = [



    (



        2,



        "Manual trip-log reconciliation",



        SRC_DIR / "trip_log_reconciliation.py",



    ),



    (



        3,



        "Candidate EV-ICE pairing",



        SRC_DIR / "candidate_pairing.py",



    ),



    (



        4,



        "Pair validation + common-section cropping",



        SRC_DIR / "pair_cropping.py",



    ),



    (



        5,



        "Post-cropping feature engineering",



        SRC_DIR / "feature_engineering.py",



    ),



    (



        6,



        "Phase segmentation",



        SRC_DIR / "phase_segmentation.py",



    ),



    (



        7,



        "TAT computation",



        SRC_DIR / "tat_computation.py",



    ),



    (



        8,



        "Results visualization",



        SRC_DIR / "results_visualization.py",



    ),



    (



        9,



        "Master INTERCITY summary + ML training datasets",



        SRC_DIR / "master_summary_intercity.py",



    ),



]







EXCLUDED_FILES = [



    SRC_DIR / "reverse_geocoding.py",



    SRC_DIR / "plot_unique_routes_and_update_summary.py",



]











# =============================================================================



# PIPELINE DEPENDENCY / ORDER RULES



# =============================================================================



#



# The verified trip index is produced by trip-log reconciliation. Therefore



# reconciliation MUST happen before EV-ICE candidate pairing.



#



STAGE_DEPENDENCIES: dict[int, tuple[int, ...]] = {



    2: (1,),



    3: (2,),



    4: (3,),



    5: (4,),



    6: (5,),



    7: (6,),



    8: (7,),



    9: (8,),



}











# =============================================================================



# GENERAL HELPERS



# =============================================================================







def format_duration(seconds: float) -> str:



    seconds = max(0, int(round(seconds)))



    hours, remainder = divmod(seconds, 3600)



    minutes, seconds = divmod(remainder, 60)



    return f"{hours:02d}:{minutes:02d}:{seconds:02d}"











def append_log(log_path: Path, text: str) -> None:



    log_path.parent.mkdir(parents=True, exist_ok=True)







    with log_path.open("a", encoding="utf-8") as handle:



        handle.write(text)



        if not text.endswith("\n"):



            handle.write("\n")











def stage_header(



    stage_number: int,



    stage_name: str,



    log_path: Path,



) -> None:



    bar = "=" * 100







    text = (



        f"\n{bar}\n"



        f"STAGE {stage_number}/{TOTAL_STAGES} | {stage_name}\n"



        f"{bar}\n"



    )







    print(text, end="", flush=True)



    append_log(log_path, text)











def validate_stage_order() -> None:



    """



    Validate the complete pipeline dependency order before any outputs are



    changed.







    Critical dependency:



        Stage 1 -> Stage 2 reconciliation -> Stage 3 pairing







    This prevents candidate pairing from running against a stale



    verified_trip_index.csv.



    """



    configured_numbers = [stage_no for stage_no, _, _ in STAGE_SCRIPTS]



    expected_numbers = list(range(2, TOTAL_STAGES + 1))







    if configured_numbers != expected_numbers:



        raise RuntimeError(



            "Invalid pipeline stage order. "



            f"Expected stages {expected_numbers}, "



            f"but configured stages are {configured_numbers}."



        )







    completed_numbers = {1}







    for stage_no, _, _ in STAGE_SCRIPTS:



        required = STAGE_DEPENDENCIES.get(stage_no, ())



        missing_dependencies = [



            dependency



            for dependency in required



            if dependency not in completed_numbers



        ]







        if missing_dependencies:



            raise RuntimeError(



                f"Invalid pipeline dependency for Stage {stage_no}. "



                f"Missing prerequisite stage(s): {missing_dependencies}."



            )







        completed_numbers.add(stage_no)











def validate_stage_files() -> None:



    """



    Check all standalone stage files before changing any pipeline outputs.



    """



    missing: list[str] = []







    for stage_no, stage_name, script_path in STAGE_SCRIPTS:



        if not script_path.is_file():



            missing.append(



                f"Stage {stage_no}/{TOTAL_STAGES} | {stage_name}\n"



                f"  Missing: {script_path}"



            )







    if missing:



        raise FileNotFoundError(



            "Required pipeline stage files are missing:\n\n"



            + "\n\n".join(missing)



        )











def validate_strict_pairing_config(config: Any) -> None:
    """Validate settings.yaml against the agreed STRICT temporal rule."""
    pairing = config.get("pairing", {}) or {}
    problems: list[str] = []

    if not bool(pairing.get("strict_pair_same_day", False)):
        problems.append("pairing.strict_pair_same_day must be true")

    if not bool(pairing.get("strict_pair_similar_start_time", False)):
        problems.append("pairing.strict_pair_similar_start_time must be true")

    try:
        tolerance = float(pairing.get("strict_start_time_tolerance_minutes"))
    except (TypeError, ValueError):
        tolerance = None

    if tolerance != 120.0:
        problems.append(
            "pairing.strict_start_time_tolerance_minutes must be 120"
        )

    if problems:
        raise RuntimeError(
            "settings.yaml does not match the agreed STRICT-pair rule:\n  - "
            + "\n  - ".join(problems)
        )


# =============================================================================



# STAGE 1



# Cleaning + permanent trip index



# =============================================================================







def run_stage_1(



    config: Any,



    log_path: Path,



) -> dict[str, Any]:



    """



    Run the original update functions as Stage 1.







    1A. Clean/audit/standardize/synchronize cleaned workbooks.



    1B. Synchronize the permanent trip index.



    """







    stage_name = "Cleaning + permanent trip index"







    stage_header(



        1,



        stage_name,



        log_path,



    )







    started = time.monotonic()







    # -------------------------------------------------------------------------



    # 1A - cleaning



    # -------------------------------------------------------------------------



    print(



        "1A | Clean, audit, standardize, and synchronize cleaned files",



        flush=True,



    )







    cleaned_summary = write_all_cleaned_files(config)







    if not cleaned_summary.get("sync_succeeded", False):



        raise RuntimeError(



            "Stage 1A failed: cleaned-file synchronization did not "



            "finish safely. Check Data_Cleaning_Audit.xlsx and the "



            "terminal messages."



        )







    # -------------------------------------------------------------------------



    # 1B - permanent trip index



    # -------------------------------------------------------------------------



    print(



        "\n1B | Synchronize permanent trip index",



        flush=True,



    )







    index_summary = update_trip_index(config)







    if not index_summary.get("index_saved", False):



        raise RuntimeError(



            "Stage 1B failed: the permanent trip index was not saved."



        )







    elapsed = time.monotonic() - started







    message = (



        f"\nSTAGE 1/{TOTAL_STAGES} COMPLETED | "



        f"{stage_name} | {format_duration(elapsed)}\n"



    )







    print(message, flush=True)



    append_log(log_path, message)







    return {



        "cleaned": cleaned_summary,



        "index": index_summary,



        "elapsed": elapsed,



    }











# =============================================================================



# STAGES 2-9



# =============================================================================







def run_python_stage(



    stage_number: int,



    stage_name: str,



    script_path: Path,



    log_path: Path,



) -> dict[str, Any]:



    """



    Execute one standalone stage using the active conda interpreter.







    Returns a dictionary instead of immediately terminating so the caller can



    choose safe-stop behavior or diagnostic --continue-on-error behavior.



    """







    stage_header(



        stage_number,



        stage_name,



        log_path,



    )







    print(



        f"File: {script_path}",



        flush=True,



    )







    append_log(



        log_path,



        f"File: {script_path}\n",



    )







    command = [



        sys.executable,



        "-u",



        str(script_path),



    ]







    append_log(



        log_path,



        "Command: "



        + subprocess.list2cmdline(command)



        + "\n",



    )







    started = time.monotonic()







    process = subprocess.Popen(



        command,



        cwd=PROJECT_ROOT,



        stdout=subprocess.PIPE,



        stderr=subprocess.STDOUT,



        text=True,



        encoding="utf-8",



        errors="replace",



        bufsize=1,



        env=os.environ.copy(),



    )







    if process.stdout is None:



        raise RuntimeError(



            f"Unable to capture output from {stage_name}"



        )







    try:



        for line in process.stdout:



            print(



                line,



                end="",



                flush=True,



            )



            append_log(



                log_path,



                line,



            )







    except KeyboardInterrupt:



        print(



            "\nCTRL+C received. Stopping current stage...",



            flush=True,



        )







        process.terminate()







        try:



            process.wait(timeout=10)



        except subprocess.TimeoutExpired:



            process.kill()



            process.wait()







        raise







    return_code = process.wait()



    elapsed = time.monotonic() - started



    success = return_code == 0







    if success:



        message = (



            f"\nSTAGE {stage_number}/{TOTAL_STAGES} COMPLETED | "



            f"{stage_name} | {format_duration(elapsed)}\n"



        )



    else:



        message = (



            f"\nSTAGE {stage_number}/{TOTAL_STAGES} FAILED | "



            f"{stage_name}\n"



            f"Exit code: {return_code}\n"



            f"Elapsed: {format_duration(elapsed)}\n"



        )







    print(



        message,



        flush=True,



    )



    append_log(



        log_path,



        message,



    )







    return {



        "stage_number": stage_number,



        "stage_name": stage_name,



        "script": script_path,



        "success": success,



        "exit_code": return_code,



        "elapsed": elapsed,



    }











# =============================================================================



# CHECK ONLY



# =============================================================================







def check_only() -> None:
    validate_stage_order()
    validate_stage_files()

    validate_strict_pairing_config(config)
    validate_strict_pairing_config(load_config())







    print("=" * 100)



    print("TAT COMPLETE PIPELINE CHECK")



    print(f"Updater version: {VERSION}")



    print("=" * 100)







    print(



        "[READY] Stage 1/9 | Cleaning + permanent trip index"



    )







    for stage_no, stage_name, script_path in STAGE_SCRIPTS:



        print(



            f"[FOUND] Stage {stage_no}/{TOTAL_STAGES} | "



            f"{stage_name}"



        )



        print(



            f"        {script_path}"



        )







    print()







    for excluded in EXCLUDED_FILES:



        print(



            f"[EXCLUDED] {excluded.name}"



        )







    print()



    print("CHECK COMPLETED SUCCESSFULLY")



    print("Nothing was executed.")











# =============================================================================



# COMPLETE PIPELINE



# =============================================================================







def update_pipeline(



    config: Any,



    continue_on_error: bool = False,



) -> dict[str, Any]:



    """



    Execute the complete pipeline in dependency order.







    Critical order:



        Stage 1: cleaning + permanent trip index



        Stage 2: trip-log reconciliation -> verified_trip_index.csv



        Stage 3: EV-ICE candidate pairing -> reads the fresh verified index







    Default:



        Stop after the first failed stage.







    Diagnostic mode:



        --continue-on-error does not run a stage whose prerequisites failed.



    """







    # Validate order and stage files before touching pipeline outputs.



    validate_stage_order()



    validate_stage_files()







    SYSTEM_DIR.mkdir(



        parents=True,



        exist_ok=True,



    )







    timestamp = datetime.now().strftime(



        "%Y%m%d_%H%M%S"



    )







    log_path = (



        SYSTEM_DIR



        / f"update_pipeline_9_stage_{timestamp}.log"



    )







    overall_started = time.monotonic()







    print("=" * 100)



    print("TAT COMPLETE 9-STAGE PIPELINE")



    print(f"Updater version: {VERSION}")



    print("=" * 100)







    print(



        f"Python:       {sys.executable}"



    )



    print(



        f"Project root: {PROJECT_ROOT}"



    )



    print(



        f"Results root: {RESULTS_ROOT}"



    )



    print(



        f"Run log:      {log_path}"



    )







    print()







    print("Stages configured to run:")



    print("  1. Cleaning + permanent trip index")







    for stage_no, stage_name, _ in STAGE_SCRIPTS:



        print(



            f"  {stage_no}. {stage_name}"



        )







    print()







    print("Explicitly excluded:")



    for excluded in EXCLUDED_FILES:



        print(



            f"  - {excluded.name}"



        )







    print()







    if continue_on_error:



        print(



            "MODE: CONTINUE ON ERROR "



            "(diagnostic only; downstream outputs may be invalid)"



        )



    else:



        print(



            "MODE: SAFE STOP "



            "(pipeline stops if a stage fails)"



        )







    print()







    append_log(



        log_path,



        (



            "TAT COMPLETE 9-STAGE PIPELINE\n"



            f"Version: {VERSION}\n"



            f"Started: {datetime.now():%d-%m-%Y %H:%M:%S}\n"



            f"Python: {sys.executable}\n"



            f"Project root: {PROJECT_ROOT}\n"



            f"Continue on error: {continue_on_error}\n"



        ),



    )







    results: list[dict[str, Any]] = []



    stage_times: dict[str, float] = {}







    # -------------------------------------------------------------------------



    # Stage 1



    # -------------------------------------------------------------------------



    stage_1_summary = run_stage_1(



        config,



        log_path,



    )







    results.append(



        {



            "stage_number": 1,



            "stage_name": "Cleaning + permanent trip index",



            "success": True,



            "exit_code": 0,



            "elapsed": stage_1_summary["elapsed"],



        }



    )







    stage_times[



        "stage_1_cleaning_and_trip_index"



    ] = stage_1_summary[



        "elapsed"



    ]







    # -------------------------------------------------------------------------



    # Stages 2-9



    # -------------------------------------------------------------------------



    #



    # Stage 2 MUST finish before Stage 3 starts.



    # Stage 2 creates/updates verified_trip_index.csv and Stage 3 reads it.



    #



    completed_successfully: set[int] = {1}







    for stage_no, stage_name, script_path in STAGE_SCRIPTS:



        dependencies = STAGE_DEPENDENCIES.get(stage_no, ())



        failed_dependencies = [



            dependency



            for dependency in dependencies



            if dependency not in completed_successfully



        ]







        if failed_dependencies:



            skip_message = (



                f"\nSTAGE {stage_no}/{TOTAL_STAGES} | {stage_name} — SKIPPED\n"



                f"Prerequisite stage(s) did not complete successfully: "



                f"{failed_dependencies}\n"



            )







            print(skip_message, flush=True)



            append_log(log_path, skip_message)







            results.append(



                {



                    "stage_number": stage_no,



                    "stage_name": stage_name,



                    "success": False,



                    "skipped": True,



                    "exit_code": None,



                    "elapsed": 0.0,



                }



            )







            stage_times[



                f"stage_{stage_no}_{script_path.stem}"



            ] = 0.0



            continue







        if stage_no == 8 and not RUN_VISUALIZATION:



            skip_message = (



                f"\nSTAGE {stage_no}/{TOTAL_STAGES} | {stage_name} — SKIPPED\n"



            )



            print(skip_message, flush=True)



            append_log(log_path, skip_message)







            # If visualization is intentionally disabled, it is treated as an



            # optional presentation step so Stage 9 can still run.



            completed_successfully.add(stage_no)







            results.append(



                {



                    "stage_number": stage_no,



                    "stage_name": stage_name,



                    "success": True,



                    "skipped": True,



                    "exit_code": 0,



                    "elapsed": 0.0,



                }



            )







            stage_times[



                f"stage_{stage_no}_{script_path.stem}"



            ] = 0.0



            continue







        stage_result = run_python_stage(



            stage_no,



            stage_name,



            script_path,



            log_path,



        )







        stage_result["skipped"] = False



        results.append(stage_result)







        stage_times[



            f"stage_{stage_no}_{script_path.stem}"



        ] = stage_result[



            "elapsed"



        ]







        if stage_result["success"]:



            completed_successfully.add(stage_no)



        elif not continue_on_error:



            stop_message = (



                "\nPIPELINE STOPPED SAFELY.\n"



                f"Stage {stage_no}/{TOTAL_STAGES} failed: "



                f"{stage_name}\n"



                "Later stages were not executed because they depend on "



                "upstream outputs.\n"



                f"Run log: {log_path}\n"



            )







            print(



                stop_message,



                flush=True,



            )



            append_log(



                log_path,



                stop_message,



            )







            raise RuntimeError(



                f"Stage {stage_no}/{TOTAL_STAGES} failed: "



                f"{stage_name}"



            )







    total_elapsed = (



        time.monotonic()



        - overall_started



    )







    failures = [



        item



        for item in results



        if not item["success"]



    ]







    return {



        "cleaned": stage_1_summary["cleaned"],



        "index": stage_1_summary["index"],



        "results": results,



        "failures": failures,



        "stage_times": stage_times,



        "total_elapsed": total_elapsed,



        "log_path": log_path,



    }











# =============================================================================



# FINAL SUMMARY



# =============================================================================







def print_final_summary(



    summary: dict[str, Any],



) -> None:



    print()







    print("=" * 100)







    if summary["failures"]:



        print(



            "TAT 9-STAGE PIPELINE FINISHED WITH FAILURES"



        )



    else:



        print(



            "TAT COMPLETE 9-STAGE PIPELINE FINISHED SUCCESSFULLY"



        )







    print("=" * 100)







    for result in summary["results"]:



        if result.get("skipped", False):



            state = "SKIPPED"



        else:



            state = (



                "SUCCESS"



                if result["success"]



                else "FAILED"



            )







        print(



            f"Stage {result['stage_number']}/{TOTAL_STAGES} | "



            f"{state} | "



            f"{result['stage_name']} | "



            f"{format_duration(result['elapsed'])}"



        )







    print()







    print(



        f"Total elapsed: {format_duration(summary['total_elapsed'])}"



    )







    print(



        f"Run log: {summary['log_path']}"



    )







    print()







    print("Excluded:")



    for excluded in EXCLUDED_FILES:



        print(



            f"  - {excluded.name}"



        )







    print("=" * 100)











# =============================================================================



# MAIN



# =============================================================================







def main() -> None:



    parser = argparse.ArgumentParser(



        description=(



            "Run the complete TAT 9-stage update pipeline."



        )



    )







    parser.add_argument(



        "--check-only",



        action="store_true",



        help=(



            "Verify all required stage scripts without running the pipeline."



        ),



    )







    parser.add_argument(



        "--continue-on-error",



        action="store_true",



        help=(



            "Diagnostic mode: attempt later stages even after a stage fails. "



            "Not recommended for normal result generation."



        ),



    )







    args = parser.parse_args()







    if args.check_only:



        check_only()



        return







    config = load_config()







    summary = update_pipeline(



        config,



        continue_on_error=args.continue_on_error,



    )







    print_final_summary(



        summary



    )







    if summary["failures"]:



        raise SystemExit(



            1



        )











if __name__ == "__main__":



    try:



        main()







    except KeyboardInterrupt:



        print(



            "\nTAT PIPELINE CANCELLED BY USER (CTRL+C)"



        )



        raise SystemExit(



            130



        )







    except SystemExit:



        raise







    except Exception as error:



        print(



            f"\nTAT PIPELINE FAILED | {error}"



        )







        traceback.print_exc()







        raise SystemExit(



            1



        )