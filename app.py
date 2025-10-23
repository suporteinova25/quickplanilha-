import os, io, secrets
from flask import Flask, render_template_string, request, redirect, session, send_file
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from openpyxl.formatting.rule import FormulaRule
import pg8000

app = Flask(__name__)
app.secret_key = secrets.token_hex(8)

# ---------- CONFIGURAÇÕES FIXAS ----------
DB_HOST = "mdm.inovaptt.com.br"   # ALTERE AQUI
DB_PORT = 5432                    # ALTERE AQUI se for outra
DB_NAME = "hmdm"                  # ALTERE AQUI se for outra
# ---------------------------------------

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

@app.route("/", methods=["GET","POST"])
def login():
    form_html = """
    <h2>Login</h2>
    <form method="post">
        Usuário:<br><input name="user" required><br>
        Senha:<br><input type="password" name="pwd" required><br><br>
        <button>Entrar</button>
    </form>
    {% if msg %}<br><small style="color:red">{{msg}}</small>{% endif %}
    """
    if request.method == "POST":
        user = request.form["user"]
        pwd  = request.form["pwd"]
        try:
            conn = pg8000.connect(
                host=DB_HOST, port=DB_PORT,
                database=DB_NAME, user=user, password=pwd
            )
            conn.close()
            session["user"] = user
            session["pwd"]  = pwd
            return redirect("/busca")
        except Exception as e:
            return render_template_string(form_html, msg=str(e))
    return render_template_string(form_html)

@app.route("/busca", methods=["GET","POST"])
def busca():
    if "user" not in session:
        return redirect("/")

    if request.method == "POST":
        sufixo = request.form["sufixo"]
        try:
            conn = pg8000.connect(
                host=DB_HOST, port=DB_PORT,
                database=DB_NAME,
                user=session["user"], password=session["pwd"]
            )
            cur = conn.cursor()
            cur.execute("""
                SELECT "number",
                       ("info"->>'model')::text,
                       ("info"->>'imei')::text,
                       ("info"->>'iccid')::text,
                       ("info"->>'phone')::text,
                       ("info"->>'serial')::text
                FROM devices
                WHERE "number" ILIKE %s
                ORDER BY "number"
            """, (f"%{sufixo}%",))
            rows = cur.fetchall()
            cur.close()
            conn.close()

            # ⚠️ Certifique-se de que modelo.xlsx está no mesmo diretório do app.py
            wb = load_workbook("modelo.xlsx")
            ws = wb.active

            for r in range(2, ws.max_row + 1):
                for c in ("B", "C", "D", "G", "H", "I", "J"):
                    ws[f"{c}{r}"].value = None

            for ridx, row in enumerate(rows, 2):
                model = normalize(row[1])
                ws[f"B{ridx}"] = model
                ws[f"C{ridx}"] = col_c(model)
                ws[f"D{ridx}"] = row[0]
                ws[f"G{ridx}"] = row[2] or ""
                ws[f"H{ridx}"] = row[3] or ""
                ws[f"I{ridx}"] = row[4] or ""
                ws[f"J{ridx}"] = row[5] or ""

            red = PatternFill(start_color="FF0000", fill_type="solid")
            lr = len(rows) + 1 if rows else 1
            rule = FormulaRule(formula=[f'AND(G2<>"",COUNTIF($G$2:$G${lr},G2)>1)'], fill=red)
            ws.conditional_formatting.add(f"G2:G{lr}", rule)

            buf = io.BytesIO()
            wb.save(buf)
            buf.seek(0)
            return send_file(
                buf,
                as_attachment=True,
                download_name=f"planilha_{sufixo}.xlsx",
                mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            # 🔍 Mostra o erro exato no navegador e logs do Render
            return f"<pre>ERRO: {e}</pre>"

    return """
    <h2>Buscar dispositivos</h2>
    <form method="post">
        Sufixo/linha: <input name="sufixo" required>
        <input type="submit" value="Gerar Excel">
    </form>
    <br><a href="/logout">Sair</a>
    """

@app.route("/logout")
def logout():
    session.clear()
    return redirect("/")

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
