from flask import Flask, render_template, request
import pandas as pd

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
}


@app.route("/")
def dashboard():

    domain = request.args.get("domain", "Gouvernance")

    cards = []
    chart_data = {}
    values = []
    auditeur = None

    if domain != "Result":

        df = pd.read_excel(
            "Data/LIVRABLE.xlsx",
            sheet_name="FinalResult",
            header=0
        )

        # remove extra spaces in Excel
        df["Domaine"] = df["Domaine"].astype(str).str.strip()

        # same for URL parameter
        domain_clean = domain.strip()

        # find matching row
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
            # ADD THIS
            cards.append({
                "title": "Auditeur",
                "value": "N/A" if pd.isna(auditeur) else str(auditeur),
                "icon": "img/7.png",
                "is_number": False
            })
            print("FINAL AUDITEUR:", auditeur, type(auditeur))

            total = int(values[2]) + int(values[3]) + int(values[4])

            yes = int(values[2])
            no = int(values[3])
            na = int(values[4])

            chart_data = {
                "yes": round((yes / total) * 100, 1) if total else 0,
                "no": round((no / total) * 100, 1) if total else 0,
                "na": round((na / total) * 100, 1) if total else 0
            }

            score_domain = round(float(values[5]) * 100, 1)

        else:
            score_domain = 0

    if domain == "Result":

        df = pd.read_excel(
            "Data/LIVRABLE.xlsx",
            sheet_name="FinalResult",
            header=None
        )

        # line 8 = row index 7 (0-based indexing)
        row = df.iloc[7]

        final_score = row[6]  # adjust column if needed

        if pd.isna(final_score):
            final_score = 0
        else:
            final_score = round(final_score, 1)

        cards = [{
            "title": "Final Score",
            "value": f"{final_score} %",
            "icon": "img/score.png"
        }]

        chart_data = {}
        score_domain = final_score

    return render_template(
        "dashboard.html",
        cards=cards,
        active=domain,
        chart_data=chart_data,
        score_domain=score_domain,
    )


if __name__ == "__main__":
    app.run(debug=True, host="127.0.0.1", port=8000)