from flask import Flask, render_template, request
import pandas as pd
import os

app = Flask(__name__)

def format_value(v):

    if isinstance(v, (int,)) or (isinstance(v, float) and v.is_integer()):
        return int(v)

    if isinstance(v, float):
        return round(v, 2)

    return v


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


@app.route("/")
def dashboard():

    domain = request.args.get("domain", "Gouvernance")

    cards = []
    chart_data = {}
    score_domain = 0
    auditeur = None
    domain_scores = []
    maturity_label = ""

    df = pd.read_excel(
        "Data/LIVRABLE.xlsx",
        sheet_name="FinalResult",
        header=0
    )

    df["Domaine"] = df["Domaine"].astype(str).str.strip()

    # ======================================================
    # RESULT TAB
    # ======================================================
    if domain == "Result":

        df_raw = pd.read_excel(
            "Data/LIVRABLE.xlsx",
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
        score=score
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

if __name__ == "__main__":
    app.run(debug=True, host="127.0.0.1", port=8000)
    #app.run()