import io
import json
import os
import sys
import re
import time
import zipfile
from copy import deepcopy
from typing import Any, Dict, List, Optional, Union
import pandas as pd
import openpyxl
import pyodbc
import requests
from jinja2 import Environment, FileSystemLoader
from weasyprint import HTML

class PAFApplicationForms:
    """Orchestrates data collection from Qualtrics API endpoints for applicants,
    fetches academic records from SQL Data Warehouse, and
    compiles clean, print-ready PDF files for each student application.
    """

    # CLASS CONSTANTS
    AUTH_KEYS = ["client_id", "client_secret", "data_center_id"]
    SERVER = "ITAPPROREPP01"
    DATABASE = "UofC_DW_Student"

    CONN_STR = (
        f"DRIVER={{ODBC Driver 18 for SQL Server}};"
        f"SERVER={SERVER};"
        f"DATABASE={DATABASE};"
        f"Trusted_Connection=yes;"
        f"TrustServerCertificate=yes;"
    )

    REQUIRED_COLUMNS_UCID_DF = {"responseId", "UCID"}

    def __init__(
        self,
        auth_creds: Dict[str, str],
        survey_id: str,
        sql_file_path: str,
        html_template_dir: str,
        output_pdf_path: str,
        template_name: str = "survey_report.html",
        ucid_question: str = "UCID#:",
        exclude_questions: Optional[List[str]] = None,
        ucid_df: Optional[pd.DataFrame] = None,
        check_dw: bool = True,
    ):
        """Initializes the PAFApplicationForms orchestrator, validates input parameter setup,

        and initiates API connections for survey data and metadata retrieval.

        Args:
            auth_creds (Dict[str, str]): Dictionary containing authentication credentials with keys
                'client_id', 'client_secret', and 'data_center_id'.
            survey_id (str): Non-empty unique identifier of the target Qualtrics survey (e.g., 'SV_...').
            sql_file_path (str): Relative or absolute path to an existing .sql query file.
            html_template_dir (str): Relative or absolute path to the directory containing HTML templates.
            output_pdf_path (str): Non-empty destination directory path where output PDFs will be written.
            template_name (str): Name of the Jinja HTML template file. Defaults to 'survey_report.html'.
            ucid_question (str): The question text asking for UCID. Must match survey metadata label.
            exclude_questions (Optional[List[str]]): List of Question texts to exclude from rendering.
            ucid_df (Optional[pd.DataFrame]): DataFrame containing 'responseId' and 'UCID' for validation mapping.
            check_dw (bool): Flag indicating whether to fetch and append Data Warehouse student records.

        Raises:
            TypeError: If input parameters do not match required data types.
            KeyError: If required authentication keys are missing from auth_creds.
            ValueError: If string values are blank/empty, or required DataFrame columns are missing.
            FileNotFoundError: If sql_file_path or html_template_dir paths do not exist.
            RuntimeError: If initialization steps or API token requests fail during setup.
        """
        # --- 1. TYPE AND RANGE VALIDATIONS ---
        if not isinstance(auth_creds, dict):
            raise TypeError(f"auth_creds must be a dict, got {type(auth_creds).__name__}")
        if not isinstance(survey_id, str):
            raise TypeError(f"survey_id must be a str, got {type(survey_id).__name__}")
        if not isinstance(sql_file_path, str):
            raise TypeError(f"sql_file_path must be a str, got {type(sql_file_path).__name__}")
        if not isinstance(html_template_dir, str):
            raise TypeError(f"html_template_dir must be a str, got {type(html_template_dir).__name__}")
        if not isinstance(output_pdf_path, str):
            raise TypeError(f"output_pdf_path must be a str, got {type(output_pdf_path).__name__}")
        if not isinstance(ucid_question, str):
            raise TypeError(f"ucid_question must be a str, got {type(ucid_question).__name__}")
        if exclude_questions is not None and not isinstance(exclude_questions, list):
            raise TypeError(f"exclude_questions must be a list, got {type(exclude_questions).__name__}")
        if ucid_df is not None and not isinstance(ucid_df, pd.DataFrame):
            raise TypeError(f"ucid_df must be a pandas DataFrame, got {type(ucid_df).__name__}")
        if not isinstance(check_dw, bool):
            raise TypeError(f"check_dw must be a boolean, got {type(check_dw).__name__}")
        if not isinstance(template_name, str):
            raise TypeError(f"template_name must be a str, got {type(template_name).__name__}")
        if not template_name.strip():
            raise ValueError("template_name cannot be empty or whitespace.")

        # --- 2. VALUE & PATH CONTENT VALIDATIONS ---
        missing_keys = [k for k in self.AUTH_KEYS if k not in auth_creds]
        if missing_keys:
            raise KeyError(f"auth_creds dictionary is missing required keys: {missing_keys}")

        for k in self.AUTH_KEYS:
            val = auth_creds[k]
            if not isinstance(val, str):
                raise TypeError(f"Value for auth_creds key '{k}' must be a string.")
            if not val.strip():
                raise ValueError(f"Value for auth_creds key '{k}' cannot be empty or whitespace.")

        if not survey_id.strip():
            raise ValueError("survey_id cannot be empty or whitespace.")
        if not sql_file_path.strip():
            raise ValueError("sql_file_path cannot be empty or whitespace.")
        if not os.path.exists(sql_file_path):
            raise FileNotFoundError(f"The SQL file at '{sql_file_path}' does not exist.")
        if not html_template_dir.strip():
            raise ValueError("html_template_dir cannot be empty or whitespace.")
        if not os.path.exists(html_template_dir):
            raise FileNotFoundError(f"HTML template directory '{html_template_dir}' does not exist.")
        if not ucid_question.strip():
            raise ValueError("ucid_question cannot be empty or whitespace.")
        if not output_pdf_path.strip():
            raise ValueError("output_pdf_path cannot be empty or whitespace.")

        full_template_path = os.path.join(html_template_dir.strip(), template_name.strip())
        if not os.path.exists(full_template_path):
            raise FileNotFoundError(f"Template file '{template_name}' not found in directory '{html_template_dir}'.")

        if exclude_questions is not None:
            for item in exclude_questions:
                if not isinstance(item, str):
                    raise TypeError("All items in exclude_questions must be string types.")
                if not item.strip():
                    raise ValueError("Items in exclude_questions cannot be empty or whitespace-only.")

        if ucid_df is not None:
            if not self.REQUIRED_COLUMNS_UCID_DF.issubset(set(ucid_df.columns)):
                raise ValueError(
                    f"ucid_df must contain required columns: {self.REQUIRED_COLUMNS_UCID_DF}"
                )

        # --- 3. ATTRIBUTE ASSIGNMENT & API SETUP ---
        try:
            self.auth_creds = auth_creds
            self.survey_id = survey_id.strip()
            self.sql_file_path = sql_file_path.strip()
            self.html_template_dir = html_template_dir.strip()
            self.output_pdf_path = output_pdf_path.strip()

            self.ucid_df = ucid_df.copy() if ucid_df is not None else None
            self.check_dw = check_dw
            self.template_name = template_name.strip()

            self.client_id = self.auth_creds["client_id"].strip()
            self.client_secret = self.auth_creds["client_secret"].strip()
            self.data_center_id = self.auth_creds["data_center_id"].strip()

            self.access_token = self.get_oauth_token()
            self.survey_data = self.fetch_responses()
            self.survey_metadata = self.fetch_survey_metadata()

            # Resolve UCID Question Text to its QID key
            self.ucid_qid = self._get_question_key_by_exact_text(
                survey_metadata=self.survey_metadata,
                target_text=ucid_question.strip(),
                strip_html=True,
                required=True
            )

            # Resolve Exclusion Questions to QID keys
            self.exclude_qids = []
            if exclude_questions is not None:
                for text in exclude_questions:
                    qid_key = self._get_question_key_by_exact_text(
                        survey_metadata=self.survey_metadata,
                        target_text=text,
                        strip_html=True,
                        required=True
                    )
                    self.exclude_qids.append(qid_key)

                # Deduplicate resolved QIDs
                self.exclude_qids = list(set(self.exclude_qids))

        except Exception as init_err:
            raise RuntimeError(f"Initialization of PAFApplicationForms failed: {init_err}") from init_err

    def get_oauth_token(self) -> str:
        """Generates an OAuth access token using Qualtrics Client ID and Client Secret.

        Returns:
            str: A valid OAuth Bearer access token string.

        Raises:
            RuntimeError: If authentication HTTP request fails or payload is invalid.
        """
        base_url = f"https://{self.data_center_id}.qualtrics.com"
        url = f"{base_url}/oauth2/token"
        data = {"grant_type": "client_credentials", "scope": "read:survey_responses"}

        try:
            response = requests.post(
                url, auth=(self.client_id, self.client_secret), data=data, timeout=10
            )
            response.raise_for_status()
            token = response.json().get("access_token")
            if not token or not isinstance(token, str):
                raise ValueError("Response did not contain a valid string 'access_token'.")
            return token

        except requests.exceptions.RequestException as req_err:
            raise RuntimeError(f"Authentication API handshake failed: {req_err}") from req_err
        except Exception as err:
            raise RuntimeError(f"Failed to obtain OAuth access token: {err}") from err

    def fetch_responses(self) -> Dict[str, Any]:
        """Triggers, polls, and downloads survey response outputs from Qualtrics API in JSON format.

        Returns:
            Dict[str, Any]: Parsed JSON dataset dictionary matching downloaded participant responses.

        Raises:
            RuntimeError: If export job initialization, polling status, or download pipeline fails.
        """
        base_url = f"https://{self.data_center_id}.qualtrics.com"
        headers = {
            "Authorization": f"Bearer {self.access_token}",
            "Content-Type": "application/json",
        }

        try:
            # Step A: Initialize the export tracking process job
            start_url = f"{base_url}/API/v3/surveys/{self.survey_id}/export-responses"
            start_res = requests.post(
                start_url, json={"format": "json"}, headers=headers, timeout=10
            )
            start_res.raise_for_status()
            progress_id = start_res.json().get("result", {}).get("progressId")
            if not progress_id:
                raise ValueError("Failed to retrieve progressId from export initiation response.")

            # Step B: Poll process queue checking status
            progress_url = f"{base_url}/API/v3/surveys/{self.survey_id}/export-responses/{progress_id}"

            while True:
                prog_res = requests.get(progress_url, headers=headers, timeout=10)
                prog_res.raise_for_status()
                result_data = prog_res.json().get("result", {})
                status = result_data.get("status")

                if status == "complete":
                    file_id = result_data.get("fileId")
                    if not file_id:
                        raise ValueError("Export complete status received but missing fileId.")
                    break
                elif status == "failed":
                    raise RuntimeError("Qualtrics response export processing engine failed on backend.")

                time.sleep(2)

            # Step C: Stream payload zip download and extract internal JSON target data
            file_url = f"{base_url}/API/v3/surveys/{self.survey_id}/export-responses/{file_id}/file"
            file_res = requests.get(file_url, headers=headers, timeout=15)
            file_res.raise_for_status()

            with zipfile.ZipFile(io.BytesIO(file_res.content)) as archive:
                for file_name in archive.namelist():
                    if file_name.endswith(".json"):
                        with archive.open(file_name) as json_file:
                            return json.loads(json_file.read().decode("utf-8"))

            raise RuntimeError("Target zip archive contained no readable JSON components.")

        except requests.exceptions.RequestException as req_err:
            raise RuntimeError(f"Qualtrics HTTP Pipeline processing exception: {req_err}") from req_err
        except Exception as general_err:
            raise RuntimeError(f"Unexpected response extraction failure: {general_err}") from general_err

    def fetch_survey_metadata(self) -> Dict[str, Any]:
        """Downloads survey structure metadata and constructs a lookup dictionary for mapping QIDs.

        Returns:
            Dict[str, Any]: Structural map of Question IDs to their question texts, types, choices, and columns.

        Raises:
            RuntimeError: If survey metadata lookup request or payload parsing fails.
        """
        base_url = f"https://{self.data_center_id}.qualtrics.com"
        url = f"{base_url}/API/v3/surveys/{self.survey_id}"
        headers = {
            "Authorization": f"Bearer {self.access_token}",
            "Content-Type": "application/json",
        }

        try:
            response = requests.get(url, headers=headers, timeout=10)
            response.raise_for_status()

            survey_structure = response.json().get("result", {})
            questions_dict = survey_structure.get("questions", {})

            mapping = {}
            for qid, metadata in questions_dict.items():
                question_text = metadata.get("questionText", f"Question ({qid})")
                q_type = metadata.get("questionType", {}).get("type", "")

                # 1. Parse standard choices
                raw_choices = metadata.get("choices", {})
                choice_map = {
                    str(c_id): c_info.get("description", f"Choice {c_id}")
                    for c_id, c_info in raw_choices.items()
                }

                # 2. Parse Side-by-Side Matrix columns
                column_map = {}
                if "columns" in metadata:
                    for col_id, col_info in metadata["columns"].items():
                        col_text = col_info.get("questionText", f"Column {col_id}")
                        nested_choices = col_info.get("choices", {})
                        col_choice_map = {
                            str(c_id): c_info.get("description", f"Option {c_id}")
                            for c_id, c_info in nested_choices.items()
                        }
                        column_map[str(col_id)] = {
                            "text": col_text,
                            "choices": col_choice_map,
                        }

                mapping[qid] = {
                    "text": question_text,
                    "type": q_type,
                    "choices": choice_map,
                    "columns": column_map,
                }

            return mapping

        except requests.exceptions.RequestException as req_err:
            raise RuntimeError(f"Qualtrics Metadata Lookup Engine Failed: {req_err}") from req_err
        except Exception as general_err:
            raise RuntimeError(f"Failed parsing Qualtrics questions payload: {general_err}") from general_err

    def fetch_student_data_by_emplid(
        self, connection_string: str, emplid: Union[int, str]
    ) -> Dict[str, Any]:
        """Reads SQL query from file and fetches a single student database record matching the provided emplid/UCID.

        Args:
            connection_string (str): pyodbc SQL Server connection string. Must be non-empty string.
            emplid (Union[int, str]): Student identifier / UCID. Must be positive int or non-empty string.

        Returns:
            Dict[str, Any]: Mapping of database column names to values. Returns {} if no record exists.

        Raises:
            TypeError: If input parameter types are invalid.
            ValueError: If connection_string is blank or emplid range/value is invalid.
            FileNotFoundError: If sql_file_path does not exist on disk.
            RuntimeError: If database connection or query execution fails.
        """
        # Validate connection_string
        if not isinstance(connection_string, str):
            raise TypeError(f"connection_string must be a str, got {type(connection_string).__name__}")
        if not connection_string.strip():
            raise ValueError("connection_string cannot be empty or whitespace.")

        # Validate sql_file_path
        if not isinstance(self.sql_file_path, str):
            raise TypeError(f"sql_file_path must be a str, got {type(self.sql_file_path).__name__}")
        if not os.path.exists(self.sql_file_path):
            raise FileNotFoundError(f"The SQL file at '{self.sql_file_path}' does not exist.")

        # Validate emplid
        if not isinstance(emplid, (int, str)):
            raise TypeError(f"emplid must be an int or str, got {type(emplid).__name__}")
        if isinstance(emplid, int) and emplid <= 0:
            raise ValueError("If emplid is an integer, it must be a positive number (> 0).")
        if isinstance(emplid, str) and not emplid.strip():
            raise ValueError("If emplid is a string, it cannot be empty or whitespace.")

        conn = None
        cursor = None
        try:
            with open(self.sql_file_path, "r", encoding="utf-8") as file:
                sql_query = file.read()

            conn = pyodbc.connect(connection_string)
            cursor = conn.cursor()
            cursor.execute(sql_query, (str(emplid).strip(),))

            row = cursor.fetchone()
            if row:
                columns = [column[0] for column in cursor.description]
                return dict(zip(columns, row))

            return {}

        except pyodbc.Error as db_err:
            raise RuntimeError(f"Database query execution error: {db_err}") from db_err
        except Exception as err:
            raise RuntimeError(f"Unexpected database retrieval failure: {err}") from err
        finally:
            if cursor is not None:
                try:
                    cursor.close()
                except Exception:
                    pass
            if conn is not None:
                try:
                    conn.close()
                except Exception:
                    pass

    def generate_student_pdf(
        self,
        student_data: Dict[str, Any],
        target_pdf_path: str,
        dw_student_data: Optional[Dict[str, Any]] = None,
        override_ucid: Optional[str] = None,
    ) -> None:
        """Parses a verified student submission record and renders a print-ready PDF document artifact.

        Args:
            student_data (Dict[str, Any]): Single response payload containing 'responseId' and 'values'.
            target_pdf_path (str): Destination path for saving the output PDF file. Must be non-empty string.
            dw_student_data (Optional[Dict[str, Any]]): Data Warehouse record dictionary.
            override_ucid (Optional[str]): Standardized/Formatted zero-padded UCID to set in response values.

        Returns:
            None

        Raises:
            TypeError: If input parameters violate type constraints.
            ValueError: If target_pdf_path is empty or student_data lacks required keys.
            RuntimeError: If template parsing or PDF rendering fails.
        """
        # Parameter Type Checks
        if not isinstance(student_data, dict):
            raise TypeError(f"student_data must be a dict, got {type(student_data).__name__}")
        if not isinstance(target_pdf_path, str):
            raise TypeError(f"target_pdf_path must be a str, got {type(target_pdf_path).__name__}")
        if not target_pdf_path.strip():
            raise ValueError("target_pdf_path cannot be empty or whitespace.")
        if "values" not in student_data or "responseId" not in student_data:
            raise ValueError("student_data must contain valid 'values' and 'responseId' subkeys.")
        if dw_student_data is not None and not isinstance(dw_student_data, dict):
            raise TypeError(f"dw_student_data must be a dict, got {type(dw_student_data).__name__}")
        if override_ucid is not None and not isinstance(override_ucid, str):
            raise TypeError(f"override_ucid must be a str, got {type(override_ucid).__name__}")

        try:
            # Create working deepcopy to avoid mutating input survey_data in place
            working_data = deepcopy(student_data)
            response_id = working_data.get("responseId")
            values = working_data.get("values", {})

            # Apply override UCID if provided
            if override_ucid:
                values["UCID"] = override_ucid

            exclude_set = set(self.exclude_qids) if self.exclude_qids else set()
            for key in list(values.keys()):
                if key in exclude_set:
                    values.pop(key, None)

            parsed_elements: List[Dict[str, Any]] = []

            # Parse Qualtrics payload using survey metadata layout
            for qid, q_meta in self.survey_metadata.items():
                q_text = q_meta.get("text", "")
                q_type = q_meta.get("type", "")

                if q_type == "DB":
                    parsed_elements.append({"type": "DB", "text": q_text})
                    continue

                if q_type == "Draw":
                    continue

                if q_type == "SBS":
                    matrix_keys = [
                        k
                        for k in values.keys()
                        if k.startswith(f"{qid}#") and values[k] not in [None, "", [], {}]
                    ]
                    if not matrix_keys:
                        continue

                    headers = [
                        col_info["text"] for col_info in q_meta.get("columns", {}).values()
                    ]
                    temp_grid: Dict[str, Dict[str, List[Any]]] = {}

                    for m_key in matrix_keys:
                        _, structural_part = m_key.split("#")
                        parts = structural_part.split("_")

                        col_id = parts[0]
                        row_id = parts[1]
                        sub_input_id = parts[2] if len(parts) > 2 else "1"

                        col_meta = q_meta.get("columns", {}).get(col_id, {})
                        raw_val = values[m_key]

                        col_title_lower = col_meta.get("text", "").lower()
                        is_date_field = any(
                            term in col_title_lower for term in ["date", "year", "month"]
                        )

                        if is_date_field:
                            real_val = raw_val
                        else:
                            real_val = col_meta.get("choices", {}).get(str(raw_val), raw_val)

                        if row_id not in temp_grid:
                            temp_grid[row_id] = {}
                        if col_id not in temp_grid[row_id]:
                            temp_grid[row_id][col_id] = []

                        if sub_input_id == "1":
                            temp_grid[row_id][col_id].insert(0, real_val)
                        else:
                            temp_grid[row_id][col_id].append(real_val)

                    row_data: Dict[str, Dict[str, Any]] = {}
                    for row_id, cols in temp_grid.items():
                        row_data[row_id] = {}
                        for col_id, val_list in cols.items():
                            col_meta = q_meta.get("columns", {}).get(col_id, {})
                            col_title = col_meta.get("text", f"Column {col_id}")

                            unique_vals = []
                            for v in val_list:
                                if v not in unique_vals:
                                    unique_vals.append(v)

                            row_data[row_id][col_title] = " - ".join(map(str, unique_vals))

                    parsed_elements.append(
                        {
                            "type": "SBS",
                            "text": q_text,
                            "headers": headers,
                            "rows": row_data,
                        }
                    )

                elif q_type == "MC":
                    raw_val = values.get(qid)
                    if raw_val in [None, "", [], {}]:
                        continue

                    if isinstance(raw_val, list):
                        labels = [q_meta.get("choices", {}).get(str(c), c) for c in raw_val]
                        parsed_elements.append(
                            {
                                "type": "MC",
                                "text": q_text,
                                "is_list": True,
                                "answer": labels,
                            }
                        )
                    else:
                        label = q_meta.get("choices", {}).get(str(raw_val), raw_val)
                        parsed_elements.append(
                            {
                                "type": "MC",
                                "text": q_text,
                                "is_list": False,
                                "answer": label,
                            }
                        )

                elif q_type == "TE":
                    te_sub_choices = q_meta.get("choices", {})
                    if te_sub_choices:
                        populated_fields = {}
                        for sub_id, sub_label in te_sub_choices.items():
                            field_key = f"{qid}_{sub_id}"
                            field_val = values.get(field_key)
                            if field_val not in [None, "", "[No Data Provided]"]:
                                populated_fields[sub_label] = field_val

                        if not populated_fields:
                            continue

                        parsed_elements.append(
                            {
                                "type": "TE",
                                "text": q_text,
                                "is_composite": True,
                                "answer": populated_fields,
                            }
                        )
                    else:
                        raw_val = values.get(qid)
                        if raw_val in [None, "", "[No Data Provided]"]:
                            raw_val = values.get(f"{qid}_TEXT")
                        if raw_val in [None, "", "[No Data Provided]"]:
                            continue

                        parsed_elements.append(
                            {
                                "type": "TE",
                                "text": q_text,
                                "is_composite": False,
                                "answer": raw_val,
                            }
                        )

            # Map Data Warehouse elements if check_dw is True
            dw_elements: List[Dict[str, Any]] = []
            if self.check_dw and dw_student_data:
                term = dw_student_data.get("Term", "N/A")
                dw_header = [
                    {
                        "type": "DB",
                        "text": f"Student Related Data Pulled by FGS — Details as of {term} Term",
                        "is_dw_header": True,
                    }
                ]
                dw_student_details = [
                    {
                        "type": "TE",
                        "text": key.replace("_", " "),
                        "is_composite": False,
                        "answer": dw_student_data.get(key, "N/A"),
                    }
                    for key in dw_student_data.keys()
                    if key != "Term"
                ]
                dw_elements.extend(dw_header + dw_student_details)

            # Jinja2 HTML compilation and WeasyPrint PDF output
            env = Environment(loader=FileSystemLoader(self.html_template_dir))
            template = env.get_template(self.template_name)
            rendered_html = template.render(
                response_id=response_id,
                parsed_elements=parsed_elements,
                dw_elements=dw_elements,
            )

            HTML(string=rendered_html).write_pdf(target_pdf_path)

        except Exception as general_err:
            raise RuntimeError(f"PDF compilation execution failed: {general_err}") from general_err

    def batch_process_student_pdfs(self) -> None:
        """Iterates through all survey responses, validates/filters records against ucid_df (if provided),
        queries Data Warehouse records, and outputs dynamic individual PDFs.

        Returns:
            None

        Raises:
            RuntimeError: If batch directory preparation or pipeline execution crashes.
        """
        try:
            if not os.path.exists(self.output_pdf_path):
                os.makedirs(self.output_pdf_path, exist_ok=True)

            responses_list = self.survey_data.get("responses", [])
            if not responses_list:
                print("[⚠️ WARNING] No responses found in the provided survey data payload.")
                return

            # Construct responseId -> Formatted UCID lookup map from ucid_df if supplied
            valid_ucid_map: Dict[str, str] = {}
            if self.ucid_df is not None:
                for _, row in self.ucid_df.iterrows():
                    resp_id = str(row["responseId"]).strip()
                    raw_u = row["UCID"]
                    if pd.notna(raw_u):
                        formatted_u = str(int(raw_u) if isinstance(raw_u, float) else raw_u).strip().zfill(8)
                        valid_ucid_map[resp_id] = formatted_u

            print(f"Starting batch compilation pipeline for {len(responses_list)} records...\n")

            for record in responses_list:
                response_id = record.get("responseId", "UnknownID")
                values = record.get("values", {})
                survey_raw_ucid = str(values.get(self.ucid_qid, "")).strip()

                # Step 1: Pre-filter check against ucid_df mapping
                if self.ucid_df is not None:
                    if response_id not in valid_ucid_map:
                        print(f" -> Skipping responseId '{response_id}': Not present in ucid_df.")
                        continue
                    final_ucid = valid_ucid_map[response_id]
                else:
                    if survey_raw_ucid and survey_raw_ucid != "[No Data Provided]":
                        clean_u = re.sub(r"[^a-zA-Z0-9\-_]", "", survey_raw_ucid)
                        final_ucid = clean_u.zfill(8) if clean_u.isdigit() else clean_u
                    else:
                        final_ucid = ""

                # Step 2: Establish target filename using DataFrame ground-truth UCID
                filename = f"{final_ucid}.pdf" if final_ucid else f"UNKNOWN_UCID_{response_id}.pdf"
                target_pdf_path = os.path.join(self.output_pdf_path, filename)

                # Step 3: Optional Data Warehouse lookup
                db_record = None
                query_ucid = final_ucid or survey_raw_ucid
                if self.check_dw and self.CONN_STR and self.sql_file_path and query_ucid:
                    try:
                        db_record = self.fetch_student_data_by_emplid(self.CONN_STR, query_ucid)
                    except Exception as fetch_err:
                        print(
                            f"   [⚠️ DW WARNING] Could not retrieve DW record for UCID {query_ucid}: {fetch_err}"
                        )

                print(f" -> Compiling PDF for Student Identifier: {final_ucid or 'UNKNOWN'} (File: {filename})...")

                # Step 4: Call generator with individual path and formatted UCID
                self.generate_student_pdf(
                    student_data=record,
                    target_pdf_path=target_pdf_path,
                    dw_student_data=db_record,
                    override_ucid=final_ucid if final_ucid else None,
                )

            print(f"\nGeneration Complete! All PDF files saved to directory: '{self.output_pdf_path}'")

        except Exception as run_err:
            raise RuntimeError(f"Batch execution engine failed during processing: {run_err}") from run_err

    @staticmethod
    def _get_question_key_by_exact_text(
        survey_metadata: Dict[str, Any],
        target_text: str,
        strip_html: bool = True,
        required: bool = True,
    ) -> Optional[str]:
        """Finds a Question ID or composite key (e.g., 'QID4_1') by matching exact text in survey metadata.

        Args:
            survey_metadata (Dict[str, Any]): Dictionary mapping Question IDs to properties.
            target_text (str): Exact question or sub-choice text to match.
            strip_html (bool): Whether to strip HTML tags before comparing. Defaults to True.
            required (bool): Whether to raise an error if no match is found. Defaults to True.

        Returns:
            Optional[str]: The identified Question ID or composite key (e.g., 'QID4_1').
                Returns None only if required is False and no match is found.

        Raises:
            TypeError: If input argument types are invalid.
            ValueError: If metadata is empty, target_text is blank, or required target cannot be matched.
            RuntimeError: If an unexpected error occurs during matching execution.
        """
        try:
            # --- 1. Type & Range Validations ---
            if not isinstance(survey_metadata, dict):
                raise TypeError(
                    f"survey_metadata must be a dict, got {type(survey_metadata).__name__}"
                )
            if not survey_metadata:
                raise ValueError("survey_metadata dictionary cannot be empty.")

            if not isinstance(target_text, str):
                raise TypeError(
                    f"target_text must be a str, got {type(target_text).__name__}"
                )
            if not target_text.strip():
                raise ValueError("target_text cannot be empty or whitespace-only.")

            if not isinstance(strip_html, bool):
                raise TypeError(
                    f"strip_html must be a bool, got {type(strip_html).__name__}"
                )

            if not isinstance(required, bool):
                raise TypeError(
                    f"required must be a bool, got {type(required).__name__}"
                )

            html_strip_pattern = re.compile(r"<[^>]+>")

            def normalize(raw_text: Any) -> str:
                """Strips HTML (if enabled), normalizes whitespace, and lowercases for comparison."""
                if not isinstance(raw_text, str):
                    return ""
                text = raw_text
                if strip_html:
                    text = html_strip_pattern.sub(" ", text)
                return " ".join(text.split()).strip().lower()

            search_target = normalize(target_text)

            # --- 2. Exact Match Search Algorithm ---
            for qid, meta in survey_metadata.items():
                if not isinstance(meta, dict):
                    continue

                # Check sub-choices first (form fields, matrix questions like 'QID4_1')
                sub_choices = meta.get("choices", {}) or meta.get("subQuestions", {})
                if isinstance(sub_choices, dict):
                    for sub_id, sub_info in sub_choices.items():
                        if isinstance(sub_info, dict):
                            sub_text = (
                                sub_info.get("description")
                                or sub_info.get("choiceText")
                                or sub_info.get("text")
                            )
                        else:
                            sub_text = str(sub_info)

                        if normalize(sub_text) == search_target:
                            resolved_key = f"{qid}_{sub_id}"
                            if not resolved_key.strip():
                                break
                            return resolved_key

                # Check main question text
                main_text = (
                    meta.get("text")
                    or meta.get("questionText")
                    or meta.get("dataName")
                )
                if normalize(main_text) == search_target:
                    resolved_key = str(qid)
                    if not resolved_key.strip():
                        break
                    return resolved_key

            # --- 3. Unresolved Target Handling ---
            if required:
                raise ValueError(
                    f"Could not resolve QID for target question text: '{target_text}'. "
                    "Ensure the string matches the Qualtrics survey metadata exactly."
                )

            return None

        except (TypeError, ValueError):
            raise
        except Exception as match_err:
            raise RuntimeError(
                f"An error occurred during exact text key lookup for '{target_text}': {match_err}"
            ) from match_err