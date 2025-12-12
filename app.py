import os, io
from flask import Flask, render_template_string, request, redirect, session, send_file
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from openpyxl.formatting.rule import FormulaRule
import pg8000
import secrets

app = Flask(__name__)
app.secret_key = secrets.token_hex(16)

DB_HOST = "mdm.inovaptt.com.br"
DB_PORT = 5432
DB_NAME = "hmdm"

def normalize(m):
    m = (m or "").strip().upper()
    if m.endswith("_O"):
        m = m[:-2]
    if m.startswith("TELOX_") or m.startswith("TELO_"):
        m = m.split("_", 1)[1] if "_" in m else m
    return m

def col_c(model_n):
    mu = model_n.upper()
    if mu.startswith("TE") or mu == "RG750" or mu.startswith("IS530"):
        return "HT"
    if mu.startswith("SM-") or mu == "RG935":
        return "TABLET"
    return ""

@app.route("/", methods=["GET", "POST"])
def login():
    form_html = """
    <h2>Login HMDM</h2>
    <form method="post">
        Usuário:<br><input name="user" required><br><br>
        Senha:<br><input type="password" name="pwd" required><br><br>
        <button>Entrar</button>
    </form>
    {% if msg %}<br><small style="color:red">{{msg}}</small>{% endif %}
    """
    if request.method == "POST":
        try:
            conn = pg8000.connect(host=DB_HOST, port=DB_PORT, database=DB_NAME,
                                  user=request.form["user"], password=request.form["pwd"])
            conn.close()
            session["user"] = request.form["user"]
            session["pwd"] = request.form["pwd"]
            return redirect("/busca")
        except Exception as e:
            return render_template_string(form_html, msg=str(e))
    return render_template_string(form_html)

@app.route("/busca", methods=["GET", "POST"])
def busca():
    if "user" not in session:
        return redirect("/")

    if request.method == "POST":
        sufixo = request.form["sufixo"].strip()

        conn = pg8000.connect(host=DB_HOST, port=DB_PORT, database=DB_NAME,
                              user=session["user"], password=session["pwd"])
        cur = conn.cursor()

        cur.execute("""
            SELECT 
                "number" AS linha_principal,

                -- CHIP 1
                NULLIF(TRIM(info::json->>'imei'), '') AS imei1,
                NULLIF(TRIM(info::json->>'iccid'), '') AS iccid1,
                NULLIF(TRIM(info::json->>'phone'), '') AS linha1,

                -- CHIP 2 (vários nomes possíveis nomes no HMDM)
                COALESCE(
                    NULLIF(TRIM(info::json->>'imei2'), ''),
                    NULLIF(TRIM(info::json->>'imeiSlot2'), ''),
                    NULLIF(TRIM(info::json->>'imei_2'), '')
                ) AS imei2,

                COALESCE(
                    NULLIF(TRIM(info::json->>'iccid2'), ''),
                    NULLIF(TRIM(info::json->>'iccidSlot2'), ''),
                    NULLIF(TRIM(info::json->>'iccid_2'), '')
                ) AS iccid2,

                COALESCE(
                    NULLIF(TRIM(info::json->>'phone2'), ''),
                    NULLIF(TRIM(info::json->>'line2'), ''),
                    NULLIF(TRIM(info::json->>'number2'), ''),
                    NULLIF(TRIM(info::json->>'phone_2'), '')
                ) AS linha2,

                (info::json->>'model') AS model,
                (info::json->>'serial') AS serial

            FROM devices
            WHERE "number" ILIKE %s
               OR info::json->>'phone2' ILIKE %s
               OR info::json->>'line2' ILIKE %s
            ORDER BY "number"
        """, (f"%{sufixo}%", f"%{sufixo}%", f"%{sufixo}%"))

        rows = cur.fetchall()
        cur.close()
        conn.close()

        wb = load_workbook("modelo.xlsx")
        ws = wb.active

        # Limpa as colunas G até L
        for r in range(2, ws.max_row + 10):
            for c in "GHIJKL":
                ws[f"{c}{r}"].value = None

        # Preenche exatamente como você pediu
        for idx, row in enumerate(rows, start=2):
            # Dados do CHIP 1
            ws[f"G{idx}"] = row[1] or ""   # IMEI 01
            ws[f"H{idx}"] = row[2] or ""   # ICCID 01
            ws[f"I{idx}"] = row[3] or row[0] or ""   # Linha 01 (usa phone se existir, senão o number)

            # Dados do CHIP 2
            ws[f"J{idx}"] = row[4] or ""   # IMEI 02
            ws[f"K{idx}"] = row[5] or ""   # ICCID 02
            ws[f"L{idx}"] = row[6] or ""   # Linha 02

            # Mantém o resto da planilha (B, C, D, etc.) como estava
            model = normalize(row[7])
            ws[f"B{idx}"] = model
            ws[f"C{idx}"] = col_c(model)
            ws[f"D{idx}"] = row[0]  # número principal do device

            # Serial (se quiser colocar em alguma coluna)
            # ws[f"??{idx}"] = row[8] or ""

        # Destaque vermelho em IMEIs duplicados
        red = PatternFill(start_color="FF0000", fill_type="solid")
        last = len(rows) + 1

        ws.conditional_formatting.add(f"G2:G{last}",
            FormulaRule(formula=[f'AND(G2<>"",COUNTIF($G$2:$G${last},G2)>1)'], fill=red))
        ws.conditional_formatting.add(f"J2:J{last}",
            FormulaRule(formula=[f'AND(J2<>"",COUNTIF($J$2:$J${last},J2)>1)'], fill=red))

        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)

        return send_file(buf, as_attachment=True,
                         download_name=f"dualchip_{sufixo}.xlsx",
                         mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    return """
    <h2>Busca Dual Chip</h2>
    <form method="post">
        Sufixo da linha: <input name="sufixo" size="30" required placeholder="ex: 9999">
        <button>Gerar Planilha</button>
    </form><br>
    <a href="/logout">Sair</a>
    """

@app.route("/logout")
def logout():
    session.clear()
    return redirect("/")

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)