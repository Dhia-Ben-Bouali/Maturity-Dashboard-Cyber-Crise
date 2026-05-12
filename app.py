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

    if domain == "Gouvernance":

        df = pd.read_excel(
            "Data/LIVRABLE.xlsx",
            sheet_name="FinalResult",
            header=None
        )

        titles = [
            "Nombre Des Questions",
            "Questions Applicable",
            "Conforme",
            "Non Conforme",
            "Questions non-applicable",
            "Score",
            "Ponderation "
        ]

        values = df.iloc[6, 1:8].values


        # =========================
        # AUDITEUR (NEW)
        # Excel cell row 6, col 9
        # =========================
        for title, value in zip(titles, values):

            if title == "Score":
                value = round(float(value) * 100, 1)
                value = f"{value} %"

            cards.append({
                "title": title,
                "value": format_value(value),
                "icon": ICON_MAP.get(title, "img/default.png")
            })


        # PIE CHART VALUES
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

    return render_template(
        "dashboard.html",
        cards=cards,
        active=domain,
        chart_data=chart_data,
        score_domain=score_domain,
    )


if __name__ == "__main__":
    app.run(debug=True, host="127.0.0.1", port=8000)