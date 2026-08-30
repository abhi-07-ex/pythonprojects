from datetime import datetime
import os
from flask import Flask, jsonify, render_template, request
import pandas as pd

app = Flask(__name__)
EXCEL_FILE = "employee_data.xlsx"

COLUMNS = [
    "Emp_ID",
    "FullName",
    "Department",
    "Designation",
    "Qualification",
    "Ex_Company",
    "Salary",
    "Address",
    "DateOfJoining",
    "WorkingDays",
]


# ==================== HELPER FUNCTIONS ====================
def load_data():
    """Excel sheet se data load karta hai aur missing columns/types fix karta hai."""
    if os.path.exists(EXCEL_FILE):
        df = pd.read_excel(EXCEL_FILE)

        # Ensure all required columns exist
        for col in COLUMNS:
            if col not in df.columns:
                df[col] = 0 if col == "Salary" else "N/A"

        # FIX 1: Emp_ID ko clean String format me standardize karna
        df["Emp_ID"] = (
            df["Emp_ID"]
            .astype(str)
            .str.replace(".0", "", regex=False)
            .str.strip()
        )
        return df[COLUMNS]

    return pd.DataFrame(columns=COLUMNS)


def save_data(df):
    """DataFrame ko Excel file mein save karta hai."""
    df.to_excel(EXCEL_FILE, index=False)


def parse_and_format_date(date_input):
    """Multiple formats (YYYY-MM-DD, DD/MM/YYYY) ko safety se parse karta hai."""
    if not date_input or pd.isna(date_input):
        return None, None

    try:
        # FIX 2: pd.to_datetime ko format detect karne ke liye robust banana
        date_obj = pd.to_datetime(date_input, errors="coerce").date()
        if pd.isna(date_obj):
            return None, None
        return date_obj, date_obj.strftime("%d/%m/%Y")
    except Exception:
        return None, None


def calculate_working_days(doj_str):
    """Total working days calculate karta hai."""
    date_obj, _ = parse_and_format_date(doj_str)
    if date_obj:
        return max(0, (datetime.today().date() - date_obj).days)
    return 0


def generate_emp_id(df):
    """Unique Employee ID (EMP1001, EMP1002...) generate karta hai."""
    if df.empty or df["Emp_ID"].dropna().empty:
        return "EMP1001"

    valid_ids = df["Emp_ID"].str.replace("EMP", "", regex=False)
    valid_ids = pd.to_numeric(valid_ids, errors="coerce").dropna()

    if valid_ids.empty:
        return "EMP1001"

    return f"EMP{int(valid_ids.max()) + 1}"


# ==================== ROUTES & APIs ====================
@app.route("/")
def home():
    return render_template("index.html")


@app.route("/api/employees", methods=["GET"])
def get_employees():
    df = load_data()

    if not df.empty:
        # Re-calculate Working Days dynamically for all existing records
        df["WorkingDays"] = df["DateOfJoining"].apply(calculate_working_days)
        # NaN values ko JSON safe conversion ke liye clean karein
        df = df.fillna("N/A")
        save_data(df)

    # Convert DataFrame to Python Dictionary/JSON format
    return jsonify(df.to_dict(orient="records"))


@app.route("/api/add", methods=["POST"])
def add_employee():
    data = request.json or {}
    df = load_data()

    name = data.get("FullName", "").strip()
    if not name.replace(" ", "").isalpha():
        return (
            jsonify(
                {
                    "success": False,
                    "message": "Invalid Name! Use only alphabets.",
                }
            ),
            400,
        )

    doj_input = data.get("DateOfJoining", "")
    date_obj, formatted_doj = parse_and_format_date(doj_input)
    if not date_obj:
        return (
            jsonify(
                {
                    "success": False,
                    "message": "Invalid Date format! Please provide a valid date.",
                }
            ),
            400,
        )

    emp_id = generate_emp_id(df)
    working_days = max(0, (datetime.today().date() - date_obj).days)

    try:
        salary_val = float(data.get("Salary", 0))
    except (ValueError, TypeError):
        salary_val = 0.0

    new_emp = pd.DataFrame(
        [
            {
                "Emp_ID": emp_id,
                "FullName": name,
                "Department": data.get("Department") or "N/A",
                "Designation": data.get("Designation") or "N/A",
                "Qualification": data.get("Qualification") or "N/A",
                "Ex_Company": data.get("Ex_Company") or "N/A",
                "Salary": salary_val,
                "Address": data.get("Address") or "N/A",
                "DateOfJoining": formatted_doj,
                "WorkingDays": working_days,
            }
        ]
    )

    df = pd.concat([df, new_emp], ignore_index=True)
    save_data(df)

    return jsonify(
        {
            "success": True,
            "message": f"Employee added successfully with ID: {emp_id}",
        }
    )


@app.route("/api/update", methods=["POST"])
def update_employee():
    data = request.json or {}
    df = load_data()

    emp_id = str(data.get("Emp_ID", "")).strip().upper()

    if emp_id not in df["Emp_ID"].values:
        return (
            jsonify({"success": False, "message": "Employee ID not found!"}),
            404,
        )

    idx = df[df["Emp_ID"] == emp_id].index[0]

    if data.get("FullName"):
        df.loc[idx, "FullName"] = str(data["FullName"]).strip()
    if data.get("Department"):
        df.loc[idx, "Department"] = str(data["Department"]).strip()
    if data.get("Designation"):
        df.loc[idx, "Designation"] = str(data["Designation"]).strip()
    if data.get("Qualification"):
        df.loc[idx, "Qualification"] = str(data["Qualification"]).strip()
    if data.get("Ex_Company"):
        df.loc[idx, "Ex_Company"] = str(data["Ex_Company"]).strip()
    if data.get("Salary") is not None and data.get("Salary") != "":
        try:
            df.loc[idx, "Salary"] = float(data["Salary"])
        except ValueError:
            pass
    if data.get("Address"):
        df.loc[idx, "Address"] = str(data["Address"]).strip()

    if data.get("DateOfJoining"):
        date_obj, formatted_doj = parse_and_format_date(data["DateOfJoining"])
        if date_obj:
            df.loc[idx, "DateOfJoining"] = formatted_doj
            df.loc[idx, "WorkingDays"] = max(
                0, (datetime.today().date() - date_obj).days
            )

    save_data(df)
    return jsonify(
        {"success": True, "message": f"Employee {emp_id} updated successfully!"}
    )


@app.route("/api/delete/<emp_id>", methods=["DELETE"])
def delete_employee(emp_id):
    df = load_data()
    emp_id_str = str(emp_id).strip().upper()

    if emp_id_str in df["Emp_ID"].values:
        df = df[df["Emp_ID"] != emp_id_str]
        save_data(df)
        return jsonify(
            {
                "success": True,
                "message": f"Employee {emp_id_str} deleted successfully!",
            }
        )
    return jsonify({"success": False, "message": "Employee ID not found!"}), 404


if __name__ == "__main__":
    app.run(debug=True, port=5000)