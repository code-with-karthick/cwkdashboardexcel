from flask import Flask, render_template, request, jsonify
import pandas as pd
from datetime import datetime
import os

app = Flask(__name__)

# --------------------------------------------------
# Configuration
# --------------------------------------------------

# Get the folder where app.py is located
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Excel file should be in the same folder as app.py
EXCEL_FILE = os.path.join(BASE_DIR, 'powerbi data.xlsx')


# --------------------------------------------------
# Login Page
# --------------------------------------------------

@app.route('/')
def login():
    return render_template('login.html')


# --------------------------------------------------
# Login API
# --------------------------------------------------

@app.route('/login', methods=['POST'])
def do_login():

    data = request.get_json() if request.is_json else request.form

    username = data.get('username')
    password = data.get('password')

    if username == 'admin' and password == 'admin123':
        return jsonify({
            'success': True,
            'message': 'Login Successful!'
        })

    return jsonify({
        'success': False,
        'message': 'Invalid Credentials!'
    }), 401


# --------------------------------------------------
# Dashboard Page
# --------------------------------------------------

@app.route('/dashboard')
def dashboard():
    return render_template('index.html')


# --------------------------------------------------
# Dashboard Data API
# --------------------------------------------------

@app.route('/api/dashboard-data')
def get_dashboard_data():

    start_date = request.args.get('start_date')
    end_date = request.args.get('end_date')

    try:

        # ------------------------------------------
        # Check Excel file
        # ------------------------------------------

        if not os.path.exists(EXCEL_FILE):
            raise FileNotFoundError(
                f"Excel file not found: {EXCEL_FILE}"
            )

        # ------------------------------------------
        # Read Excel
        # ------------------------------------------

        df = pd.read_excel(EXCEL_FILE)

        # Remove spaces from column names
        df.columns = df.columns.str.strip()

        # ------------------------------------------
        # Convert Admission Date
        # ------------------------------------------

        if 'Admission_Date' in df.columns:

            df['Admission_Date'] = pd.to_datetime(
                df['Admission_Date'],
                errors='coerce'
            ).dt.date

        else:

            df['Admission_Date'] = pd.NaT

        # ------------------------------------------
        # Total rows in Excel
        # ------------------------------------------

        total_rows_in_excel = len(df)

        # ------------------------------------------
        # Apply Date Filter
        # ------------------------------------------

        if start_date and end_date:

            try:

                s_date = datetime.strptime(
                    start_date,
                    '%Y-%m-%d'
                ).date()

                e_date = datetime.strptime(
                    end_date,
                    '%Y-%m-%d'
                ).date()

                df_filtered = df[
                    (df['Admission_Date'] >= s_date) &
                    (df['Admission_Date'] <= e_date)
                ]

            except ValueError:

                # Invalid date format
                df_filtered = df

        else:

            df_filtered = df

        # ------------------------------------------
        # Total Students
        # ------------------------------------------

        total_students = int(len(df_filtered))

        # ------------------------------------------
        # Active Students
        # ------------------------------------------

        if 'Status' in df_filtered.columns:

            active_students = int(
                len(
                    df_filtered[
                        df_filtered['Status']
                        .astype(str)
                        .str.strip()
                        .str.lower() == 'active'
                    ]
                )
            )

        else:

            active_students = 0

        # ------------------------------------------
        # Amount Paid
        # ------------------------------------------

        if 'Amount_Paid_INR' in df_filtered.columns:

            amount_paid = int(
                pd.to_numeric(
                    df_filtered['Amount_Paid_INR'],
                    errors='coerce'
                ).fillna(0).sum()
            )

        else:

            amount_paid = 0

        # ------------------------------------------
        # Balance Due
        # ------------------------------------------

        if 'Balance_Due_INR' in df_filtered.columns:

            balance_due = int(
                pd.to_numeric(
                    df_filtered['Balance_Due_INR'],
                    errors='coerce'
                ).fillna(0).sum()
            )

        else:

            balance_due = 0

        # ------------------------------------------
        # Course Counts
        # ------------------------------------------

        if (
            'Course_Enrolled' in df_filtered.columns
            and total_students > 0
        ):

            course_counts = (
                df_filtered['Course_Enrolled']
                .fillna('Unknown')
                .astype(str)
                .str.strip()
                .value_counts()
            )

        else:

            course_counts = pd.Series(dtype='int64')

        chart_labels = list(course_counts.index)

        chart_data = [
            int(value)
            for value in course_counts.values
        ]

        # ------------------------------------------
        # Certificate Issued
        # ------------------------------------------

        if 'Certificate_Issued' in df_filtered.columns:

            certificate_issued = int(
                len(
                    df_filtered[
                        df_filtered['Certificate_Issued']
                        .astype(str)
                        .str.strip()
                        .str.lower() == 'yes'
                    ]
                )
            )

        else:

            certificate_issued = 0

        # ------------------------------------------
        # Build Student List
        # ------------------------------------------

        students_list = []

        for _, row in df_filtered.iterrows():

            # Student name
            name = str(
                row.get('Student_Name', '')
            )

            # Course
            course = str(
                row.get('Course_Enrolled', '')
            )

            # Base fee
            base_fee = pd.to_numeric(
                row.get('Base_Fee_INR', 0),
                errors='coerce'
            )

            if pd.isna(base_fee):
                base_fee = 0

            # Amount paid
            paid = pd.to_numeric(
                row.get('Amount_Paid_INR', 0),
                errors='coerce'
            )

            if pd.isna(paid):
                paid = 0

            # Balance due
            due = pd.to_numeric(
                row.get('Balance_Due_INR', 0),
                errors='coerce'
            )

            if pd.isna(due):
                due = 0

            # Status
            status = str(
                row.get('Status', '')
            )

            # Admission date
            admission_date = row.get(
                'Admission_Date',
                ''
            )

            students_list.append({

                "name": name,

                "course": course,

                "base_fee": int(base_fee),

                "paid": int(paid),

                "due": int(due),

                "status": status,

                "date": str(admission_date)

            })

        # ------------------------------------------
        # Final Dashboard Data
        # ------------------------------------------

        data = {

            "total_students":
                total_students,

            "active_students":
                active_students,

            "today_walkins":
                total_students,

            "today_admissions":
                total_students,

            "this_month_join":
                total_students,

            "prev_month_join":
                0,

            "this_year_join":
                total_rows_in_excel,

            "certificate_issued":
                certificate_issued,

            "pending_revenue":
                balance_due,

            "study_material_distributed":
                total_students,

            "amount_paid":
                amount_paid,

            "balance_due":
                balance_due,

            "chart_labels":
                chart_labels,

            "chart_data":
                chart_data,

            "students_list":
                students_list
        }

    # ----------------------------------------------
    # Error Handling
    # ----------------------------------------------

    except Exception as e:

        print(
            "Error reading Excel:",
            str(e)
        )

        data = {

            "total_students": 0,

            "active_students": 0,

            "today_walkins": 0,

            "today_admissions": 0,

            "this_month_join": 0,

            "prev_month_join": 0,

            "this_year_join": 0,

            "certificate_issued": 0,

            "pending_revenue": 0,

            "study_material_distributed": 0,

            "amount_paid": 0,

            "balance_due": 0,

            "chart_labels": [],

            "chart_data": [],

            "students_list": [],

            "error": str(e)
        }

    return jsonify(data)


# --------------------------------------------------
# Run Flask Application
# --------------------------------------------------

if __name__ == '__main__':

    app.run(
        host='127.0.0.1',
        port=5000,
        debug=True,
        use_reloader=False
    )




