from flask import Flask, render_template, request, redirect, url_for,  make_response
from weasyprint import HTML
import shutil
import pandas as pd
import os

app = Flask(__name__)


    # ======================================================
    # Function Definition 
    # ======================================================
def format_value(v):

    if isinstance(v, (int,)) or (isinstance(v, float) and v.is_integer()):
        return int(v)

    if isinstance(v, float):
        return round(v, 2)

    return v

def get_result_data(file_path):

    df_raw = pd.read_excel(
        file_path,
        sheet_name="FinalResult",
        header=None
    )

    row = df_raw.iloc[7]

    def fix(v):
        if isinstance(v, str):
            return float(v.replace(",", "."))
        return round(float(v), 2)

    values = [
        fix(row[1]),  # Total Questions
        fix(row[2]),  # Applicable
        fix(row[3]),  # Conforme
        fix(row[4]),  # Non Conforme
        fix(row[5]),  # Non Applicable
        fix(row[6]),  # Score Final
    ]

    maturity_levels = [
        {"label": "Incomplet", "min": 0, "max": 16},
        {"label": "Exécuté", "min": 17, "max": 33},
        {"label": "Maîtrisé", "min": 34, "max": 51},
        {"label": "Établi", "min": 52, "max": 68},
        {"label": "Prévisible", "min": 69, "max": 85},
        {"label": "Optimisé", "min": 86, "max": 100},
    ]

    scoremat = df_raw.iloc[6, 6] *100
    print(scoremat)

    maturity = "N/A"

    for level in maturity_levels:
        if level["min"] <= scoremat <= level["max"]:
            maturity = level["label"]
            break

    return {
        "total_questions": int(values[0]),
        "applicable": int(values[1]),
        "conforme": int(values[2]),
        "non_conforme": int(values[3]),
        "non_applicable": int(values[4]),
        "score_final": round(values[5], 1),
        "maturity": maturity
    }

    # ======================================================
    # CONST
    # ======================================================
ICON_MAP = {
    "Nombre Des Questions": "img/1.png",
    "Questions Applicable": "img/3.png",
    "Conforme": "img/2.png",
    "Non Conforme": "img/4.png",
    "Questions non-applicable": "img/8.png",
    "Score": "img/5.png",
    "Ponderation ": "img/6.png",
    "Score Final": "img/9.png"
}

EXCEL_FOLDER = "Data"
TEMPLATE_FILE = "Data/Default_Template.xlsx"

# Default route
@app.route("/")
def index():
    return redirect(url_for("home"))

@app.route("/evaluation", methods=["POST"])
def open_evaluation():

    selected_file = request.form["excelFile"]

    return redirect(
        url_for(
            "dashboard",
            filename=selected_file
        )
    )


@app.route("/evaluation/<filename>")
def dashboard(filename):

    file_path = os.path.join("Data", filename)
    domain = request.args.get("domain", "Gouvernance")

    cards = []
    chart_data = {}
    score_domain = 0
    auditeur = None
    domain_scores = []
    maturity_label = ""

    df = pd.read_excel(
        file_path,
        sheet_name="FinalResult",
        header=0
    )

    df["Domaine"] = df["Domaine"].astype(str).str.strip()

    # ======================================================
    # RESULT TAB
    # ======================================================
    if domain == "Result":

        df_raw = pd.read_excel(
            file_path,
            sheet_name="FinalResult",
            header=None
        )

        row = df_raw.iloc[7]  # line 8

        def fix(v):
            if isinstance(v, str):
                return float(v.replace(",", "."))
                
            return round(v, 2)

        values = [
            fix(row[1]),
            fix(row[2]),
            fix(row[3]),
            fix(row[4]),
            fix(row[5]),
            fix(row[6]),
        ]

        titles = [
            "Total Questions",
            "Questions Applicable",
            "Conforme",
            "Non Conforme",
            "Questions non-applicable",
            "Score Final",
        ]

        for t, v in zip(titles, values):

            cards.append({
                "title": t,
                "value": int(v) if t != "Score Domain" else round(v, 2),
                "icon": ICON_MAP.get(t, "img/default.png"),
                "is_number": True
            })
        allowed_domains = [
        "Forensics",
        "Continuité d’activité",
        "Gestion de crise",
        "Gestion Incidents",
        "Confirmitee",
        "Gouvernance"
        ]

        for _, r in df.iterrows():

            domain_name = str(r["Domaine"]).strip()

            if domain_name in allowed_domains:

                score = r["ScoreDomain"]

                if isinstance(score, str):
                    score = float(score.replace(",", "."))

                score = round(float(score) * 100, 1)

                cards.append({
                    "title": domain_name,
                    "value": f"{score} %",
                    "icon": "img/5.png",
                    "is_number": False
                })
        score_domain = values[5]

        yes = values[2]
        no = values[3]
        na = values[4]

        total = yes + no + na

        chart_data = {
            "yes": round((yes / total) * 100, 1) if total else 0,
            "no": round((no / total) * 100, 1) if total else 0,
            "na": round((na / total) * 100, 1) if total else 0
        }

        maturity_levels = [
            {"label": "Incomplet", "min": 0, "max": 16},
            {"label": "Exécuté", "min": 17, "max": 33},
            {"label": "Maîtrisé", "min": 34, "max": 51},
            {"label": "Établi", "min": 52, "max": 68},
            {"label": "Prévisible", "min": 69, "max": 85},
            {"label": "Optimisé", "min": 86, "max": 100},
        ]

        maturity_label = "N/A"
        scoremat = df.iloc[6, 6]
        print("scoremat", scoremat)
        for level in maturity_levels:

            if level["min"] <= scoremat <= level["max"]:
                maturity_label = level["label"]
                break
    # ======================================================
    # DOMAIN TAB (ALL OTHER DOMAINS)
    # ======================================================
    else:

        domain_clean = domain.strip()
        score=0

        row = df[df["Domaine"] == domain_clean]

        if not row.empty:

            row = row.iloc[0]

            titles = [
                "Nombre Des Questions",
                "Questions Applicable",
                "Conforme",
                "Non Conforme",
                "Questions non-applicable",
                "Score",
                "Ponderation "
            ]

            values = [
                row["NbQuestion"],
                row["NbQuestionApp"],
                row["Yes"],
                row["No"],
                row["N/A"],
                row["ScoreDomain"],
                row["Ponderation"]
            ]

            auditeur = row.get("Auditeur", "")
            if pd.isna(auditeur):
                auditeur = "N/A"
            else:
                auditeur = str(auditeur).strip()

            for title, value in zip(titles, values):

                if title == "Score":
                    value = round(float(value) * 100, 1)
                    value = f"{value} %"

                cards.append({
                    "title": title,
                    "value": format_value(value),
                    "icon": ICON_MAP.get(title, "img/default.png")
                })

            cards.append({
                "title": "Auditeur",
                "value": auditeur,
                "icon": "img/7.png",
                "is_number": False
            })

            total = int(values[2]) + int(values[3]) + int(values[4])

            chart_data = {
                "yes": round((values[2] / total) * 100, 1) if total else 0,
                "no": round((values[3] / total) * 100, 1) if total else 0,
                "na": round((values[4] / total) * 100, 1) if total else 0
            }

            score_domain = round(float(values[5]) * 100, 1)


    # ======================================================
    # DOMAIN SCORES (ALL DOMAINS ROWS 2–7)
    # ======================================================

    subset = df.iloc[1:7]

    for _, r in subset.iterrows():

        sc = r["ScoreDomain"]

        if isinstance(sc, str):
            sc = float(sc.replace(",", "."))

        domain_scores.append({
            "domain": r["Domaine"],
            "score": round(float(sc) * 100, 1)
        })

    return render_template(
        "dashboard.html",
        cards=cards,
        active=domain,
        chart_data=chart_data,
        score_domain=score_domain,
        domain_scores=domain_scores,
        maturity_label=maturity_label,
        score=score,
        selected_file=filename
    )

@app.route("/Home")
def home():

    excel_files = [
        f for f in os.listdir(EXCEL_FOLDER)
        if f.endswith(".xlsx") or f.endswith(".xls")
    ]

    return render_template(
        "home.html",
        excel_files=excel_files
    )

@app.route("/create-evaluation", methods=["POST"])
def create_evaluation():

    entreprise = request.form.get("entreprise")
    date = request.form.get("date")
    auditor = request.form.get("auditor")

    entreprise = entreprise.replace(" ", "_")
    auditor = auditor.replace(" ", "_")

    filename = f"{entreprise}_{date}_{auditor}.xlsx"

    destination = os.path.join(EXCEL_FOLDER, filename)


    shutil.copy(TEMPLATE_FILE, destination)

    return redirect("/Home")

@app.route('/generate-report/<filename>')
def generate_report(filename):

    file_path = os.path.join("Data", filename)

    if not os.path.exists(file_path):
        return "File not found", 404

    # =====================================================
    # READ MAIN TABLE
    # =====================================================

    df = pd.read_excel(
        file_path,
        sheet_name="FinalResult",
        header=0
    )

    allowed_domains = [
        "Forensics",
        "Continuité d’activité",
        "Gestion de crise",
        "Gestion Incidents",
        "Confirmitee",
        "Gouvernance"
    ]

    domains = []

    for _, row in df.iterrows():

        domain = str(row["Domaine"]).strip()

        if domain in allowed_domains:

            score = row["ScoreDomain"]

            if isinstance(score, str):
                score = round(float(str(score).replace(",", ".")), 2)

            score_percent = round(float(score) * 100, 1)

            if score_percent <= 16:
                status = "Incomplet"
            elif score_percent <= 33:
                status = "Exécuté"
            elif score_percent <= 51:
                status = "Maîtrisé"
            elif score_percent <= 68:
                status = "Établi"
            elif score_percent <= 85:
                status = "Prévisible"
            else:
                status = "Optimisé"

            domains.append({
                "domain": domain,
                "nb_question": int(row["NbQuestion"]),
                "nb_app": int(row["NbQuestionApp"]),
                "yes": int(row["Yes"]),
                "no": int(row["No"]),
                "na": int(row["N/A"]),
                "score": score_percent,
                "ponderation": float(row["Ponderation"]),
                "auditeur": str(row["Auditeur"]),
                "status": status
            })

    # =====================================================
    # RESULT ROW
    # =====================================================

    result_row = df[
        df["Domaine"].astype(str).str.strip() == "Result"
    ].iloc[0]

    final_score = round(float(result_row["ScoreDomain"]), 2)

    if final_score <= 1:
        final_score = round(final_score * 100, 1)

    # =====================================================
    # READ MATURITY TABLE
    # =====================================================

    df_raw = pd.read_excel(
        file_path,
        sheet_name="FinalResult",
        header=None
    )

    maturity_label = "N/A"

    for i in range(len(df_raw)):

        try:

            label = str(df_raw.iloc[i, 1]).strip()

            if label in [
                "Incomplet",
                "Exécuté",
                "Maîtrisé",
                "Établi",
                "Prévisible",
                "Optimisé"
            ]:

                min_score = float(df_raw.iloc[i, 2])
                max_score = float(df_raw.iloc[i, 3])

                if min_score <= final_score <= max_score:
                    maturity_label = label
                    break

        except:
            pass

    # =====================================================
    # SUMMARY
    # =====================================================

    summary = {
        "total_questions": int(result_row["NbQuestion"]),
        "applicable": int(result_row["NbQuestionApp"]),
        "yes": int(result_row["Yes"]),
        "no": int(result_row["No"]),
        "na": int(result_row["N/A"]),
        "score_final": final_score
    }

    # =====================================================
    # GENERATE PDF
    # =====================================================

    rendered = render_template(
        "Report.html",
        filename=filename,
        summary=summary,
        domains=domains,
        maturity_label=maturity_label
    )

    pdf = HTML(string=rendered).write_pdf()

    response = make_response(pdf)
    response.headers["Content-Type"] = "application/pdf"
    response.headers["Content-Disposition"] = (
        f'inline; filename="{filename}_report.pdf"'
    )

    return response

if __name__ == "__main__":
    app.run(debug=True, host="127.0.0.1", port=8000)
    #app.run()