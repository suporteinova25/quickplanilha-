import os
import io
import re
from flask import Flask, render_template_string, request, redirect, session, send_file
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font
from openpyxl.formatting.rule import FormulaRule
import pg8000
import secrets

app = Flask(__name__)
app.secret_key = secrets.token_hex(16)

# ==================== CONFIGURAÇÕES ====================
DB_HOST = "mdm.inovaptt.com.br"
DB_PORT = 5432
DB_NAME = "hmdm"
# =======================================================

def normalize(m):
    if not m:
        return ""
    m = m.strip().upper()
    if m.endswith("_O"):
        m = m[:-2]
    if m.startswith("TELOX_") or m.startswith("TELO_"):
        if "_" in m:
            m = m.split("_", 1)[1]
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
        Usuário:<br><input name="user" required autocomplete="off"><br><br>
        Senha:<br><input type="password" name="pwd" required><br><br>
        <button>Entrar</button>
    </form>
    {% if msg %}<br><small style="color:red">{{msg}}</small>{% endif %}
    """
    if request.method == "POST":
        user = request.form["user"]
        pwd = request.form["pwd"]
        try:
            with pg8000.connect(host=DB_HOST, port=DB_PORT, database=DB_NAME,
                                user=user, password=pwd) as conn:
                pass  # só testa conexão
            session["authenticated"] = True
            session["user"] = user
            session["pwd"] = pwd  # ainda não ideal, mas comum em apps internos pequenos
            return redirect("/busca")
        except Exception as e:
            return render_template_string(form_html, msg="Erro de login: " + str(e))
    return render_template_string(form_html)

@app.route("/busca", methods=["GET", "POST"])
def busca():
    if not session.get("authenticated"):
        return redirect("/")

    if request.method == "POST":
        sufixo = request.form["sufixo"].strip()
        if len(sufixo) < 3:
            return "Sufixo muito curto. Use pelo menos 3 dígitos."

        try:
            with pg8000.connect(host=DB_HOST, port=DB_PORT, database=DB_NAME,
                                user=session["user"], password=session["pwd"]) as conn:
                with conn.cursor() as cur:
                    cur.execute("""
                        SELECT
                            "number",
                            NULLIF(TRIM(info::json->>'imei'), '') AS imei1,
                            NULLIF(TRIM(info::json->>'iccid'), '') AS iccid1,
                            NULLIF(TRIM(info::json->>'phone'), '') AS linha1,
                            COALESCE(
                                NULLIF(TRIM(info::json->>'imei2'), ''),
                                NULLIF(TRIM(info::json->>'imeiSlot2'), '')
                            ) AS imei2,
                            COALESCE(
                                NULLIF(TRIM(info::json->>'iccid2'), ''),
                                NULLIF(TRIM(info::json->>'iccidSlot2'), '')
                            ) AS iccid2,
                            COALESCE(
                                NULLIF(TRIM(info::json->>'phone2'), ''),
                                NULLIF(TRIM(info::json->>'line2'), ''),
                                NULLIF(TRIM(info::json->>'number2'), '')
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

            if not rows:
                return "Nenhum dispositivo encontrado com esse sufixo."

            wb = load_workbook("modelo.xlsx")
            ws = wb.active

            # Limpa a partir da linha 2 nas colunas relevantes
            for row in ws["B2:M" + str(ws.max_row + 10)]:
                for cell in row:
                    cell.value = None

            roxo = Font(color="9C27B0")
            vermelho = Font(color="F44336")
            azul = Font(color="1976D2")
            vermelho_fundo = PatternFill(start_color="FF0000", fill_type="solid")

            for idx, row in enumerate(rows, start=2):
                ws[f"G{idx}"] = row[1] or ""
                ws[f"H{idx}"] = row[2] or ""
                ws[f"I{idx}"] = row[3] or row[0] or ""
                ws[f"J{idx}"] = row[4] or ""
                ws[f"K{idx}"] = row[5] or ""
                ws[f"L{idx}"] = row[6] or ""
                ws[f"M{idx}"] = row[8] or ""

                model = normalize(row[7])
                ws[f"B{idx}"] = model
                ws[f"C{idx}"] = col_c(model)
                ws[f"D{idx}"] = row[0]

                # Cores nos ICCIDs
                for col, iccid in [("H", row[2]), ("K", row[5])]:
                    iccid_str = str(iccid or "")
                    if iccid_str.startswith("895510"):
                        ws[f"{col}{idx}"].font = roxo
                    elif iccid_str.startswith("8955053"):
                        ws[f"{col}{idx}"].font = vermelho
                    else:
                        ws[f"{col}{idx}"].font = azul

            ultima_linha = len(rows) + 1
            ws.conditional_formatting.add(f"G2:G{ultima_linha}",
                FormulaRule(formula=[f'AND(G2<>"",COUNTIF($G$2:$G${ultima_linha},G2)>1)'],
                            fill=vermelho_fundo))
            ws.conditional_formatting.add(f"J2:J{ultima_linha}",
                FormulaRule(formula=[f'AND(J2<>"",COUNTIF($J$2:$J${ultima_linha},J2)>1)'],
                            fill=vermelho_fundo))

            buf = io.BytesIO()
            wb.save(buf)
            buf.seek(0)

            safe_sufixo = re.sub(r'[^\w\-]', '_', sufixo)
            return send_file(
                buf,
                as_attachment=True,
                download_name=f"dispositivos_{safe_sufixo}.xlsx",
                mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            return f"Erro ao conectar ou processar dados: {str(e)}<br><a href='/logout'>Sair</a>"

    return """
    <h2>Busca Dispositivos Dual-Chip</h2>
    <form method="post">
        Sufixo da linha: <input name="sufixo" size="35" required placeholder="ex: 9999 ou 1199999">
        <button style="padding:10px">Gerar Planilha Excel</button>
    </form>
    <br><a href="/logout">Sair</a>
    """

@app.route("/logout")
def logout():
    session.clear()
    return redirect("/")

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)  # debug=False em produção