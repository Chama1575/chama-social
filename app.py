
from flask import Flask, render_template, request, redirect, url_for, flash, send_file
import sqlite3
import io
import xlsxwriter
from pathlib import Path

app = Flask(__name__)
app.secret_key = "chama-social-v3"

DB_PATH = Path(__file__).parent / "chama_social.db"

def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    conn = get_db()
    conn.executescript("""
    CREATE TABLE IF NOT EXISTS eventos (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        nome TEXT NOT NULL,
        data TEXT,
        local TEXT,
        permitir_cpf_repetido INTEGER NOT NULL DEFAULT 0,
        ativo INTEGER NOT NULL DEFAULT 1
    );

    CREATE TABLE IF NOT EXISTS campos (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        evento_id INTEGER NOT NULL,
        titulo TEXT NOT NULL,
        tipo TEXT NOT NULL,
        obrigatorio INTEGER NOT NULL DEFAULT 0,
        opcoes TEXT,
        ordem INTEGER NOT NULL DEFAULT 0,
        marcador_cpf INTEGER NOT NULL DEFAULT 0,
        FOREIGN KEY(evento_id) REFERENCES eventos(id)
    );

    CREATE TABLE IF NOT EXISTS atendimentos (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        evento_id INTEGER NOT NULL,
        nome TEXT NOT NULL,
        vagas INTEGER NOT NULL DEFAULT 0,
        ativo INTEGER NOT NULL DEFAULT 1,
        FOREIGN KEY(evento_id) REFERENCES eventos(id)
    );

    CREATE TABLE IF NOT EXISTS inscricoes (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        evento_id INTEGER NOT NULL,
        atendimento_id INTEGER,
        criado_em DATETIME DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY(evento_id) REFERENCES eventos(id),
        FOREIGN KEY(atendimento_id) REFERENCES atendimentos(id)
    );


    CREATE TABLE IF NOT EXISTS campos_voluntarios (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        evento_id INTEGER NOT NULL,
        titulo TEXT NOT NULL,
        tipo TEXT NOT NULL,
        obrigatorio INTEGER NOT NULL DEFAULT 0,
        opcoes TEXT,
        ordem INTEGER NOT NULL DEFAULT 0,
        marcador_cpf INTEGER NOT NULL DEFAULT 0,
        FOREIGN KEY(evento_id) REFERENCES eventos(id)
    );

    CREATE TABLE IF NOT EXISTS voluntarios (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        evento_id INTEGER NOT NULL,
        criado_em DATETIME DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY(evento_id) REFERENCES eventos(id)
    );

    CREATE TABLE IF NOT EXISTS respostas_voluntarios (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        voluntario_id INTEGER NOT NULL,
        campo_id INTEGER NOT NULL,
        valor TEXT,
        FOREIGN KEY(voluntario_id) REFERENCES voluntarios(id),
        FOREIGN KEY(campo_id) REFERENCES campos_voluntarios(id)
    );

    CREATE TABLE IF NOT EXISTS respostas (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        inscricao_id INTEGER NOT NULL,
        campo_id INTEGER NOT NULL,
        valor TEXT,
        FOREIGN KEY(inscricao_id) REFERENCES inscricoes(id),
        FOREIGN KEY(campo_id) REFERENCES campos(id)
    );
    """)
    # Migração leve para banco antigo sem a coluna atendimento_id
    cols = [r[1] for r in conn.execute("PRAGMA table_info(inscricoes)").fetchall()]
    if "atendimento_id" not in cols:
        conn.execute("ALTER TABLE inscricoes ADD COLUMN atendimento_id INTEGER")
    conn.commit()
    conn.close()

@app.route("/")
def index():
    conn = get_db()
    eventos = conn.execute("""
        SELECT e.*,
               (SELECT COUNT(*) FROM campos c WHERE c.evento_id = e.id) AS total_campos,
               (SELECT COUNT(*) FROM inscricoes i WHERE i.evento_id = e.id) AS total_inscricoes,
               (SELECT COUNT(*) FROM atendimentos a WHERE a.evento_id = e.id) AS total_atendimentos,
               (SELECT COUNT(*) FROM voluntarios v WHERE v.evento_id = e.id) AS total_voluntarios
        FROM eventos e
        ORDER BY e.id DESC
    """).fetchall()
    conn.close()
    return render_template("index.html", eventos=eventos)

@app.route("/evento/novo", methods=["GET", "POST"])
def novo_evento():
    if request.method == "POST":
        nome = request.form["nome"].strip()
        data = request.form.get("data", "").strip()
        local = request.form.get("local", "").strip()
        permitir = 1 if request.form.get("permitir_cpf_repetido") == "on" else 0

        conn = get_db()
        cur = conn.execute("""
            INSERT INTO eventos (nome, data, local, permitir_cpf_repetido)
            VALUES (?, ?, ?, ?)
        """, (nome, data, local, permitir))
        evento_id = cur.lastrowid
        conn.commit()
        conn.close()
        return redirect(url_for("construtor", evento_id=evento_id))

    return render_template("novo_evento.html")

@app.route("/evento/<int:evento_id>/construtor", methods=["GET", "POST"])
def construtor(evento_id):
    conn = get_db()
    evento = conn.execute("SELECT * FROM eventos WHERE id = ?", (evento_id,)).fetchone()
    if not evento:
        conn.close()
        return "Evento não encontrado.", 404

    if request.method == "POST":
        acao = request.form.get("acao", "campo")

        if acao == "atendimento":
            nome = request.form["nome_atendimento"].strip()
            vagas = int(request.form.get("vagas", 0) or 0)
            if nome and vagas >= 0:
                conn.execute("""
                    INSERT INTO atendimentos (evento_id, nome, vagas)
                    VALUES (?, ?, ?)
                """, (evento_id, nome, vagas))
                conn.commit()
                flash("Atendimento adicionado.")
            return redirect(url_for("construtor", evento_id=evento_id))

        titulo = request.form["titulo"].strip()
        tipo = request.form["tipo"]
        obrigatorio = 1 if request.form.get("obrigatorio") == "on" else 0
        marcador_cpf = 1 if request.form.get("marcador_cpf") == "on" else 0
        opcoes = request.form.get("opcoes", "").strip()

        if marcador_cpf:
            conn.execute("UPDATE campos SET marcador_cpf = 0 WHERE evento_id = ?", (evento_id,))

        ordem = conn.execute(
            "SELECT COALESCE(MAX(ordem),0)+1 FROM campos WHERE evento_id = ?",
            (evento_id,)
        ).fetchone()[0]

        conn.execute("""
            INSERT INTO campos (evento_id, titulo, tipo, obrigatorio, opcoes, ordem, marcador_cpf)
            VALUES (?, ?, ?, ?, ?, ?, ?)
        """, (evento_id, titulo, tipo, obrigatorio, opcoes, ordem, marcador_cpf))
        conn.commit()
        flash("Campo adicionado.")
        return redirect(url_for("construtor", evento_id=evento_id))

    campos = conn.execute("""
        SELECT * FROM campos
        WHERE evento_id = ?
        ORDER BY ordem, id
    """, (evento_id,)).fetchall()

    atendimentos = conn.execute("""
        SELECT a.*,
               (SELECT COUNT(*) FROM inscricoes i WHERE i.atendimento_id = a.id) AS ocupadas
        FROM atendimentos a
        WHERE a.evento_id = ?
        ORDER BY a.id
    """, (evento_id,)).fetchall()

    conn.close()
    return render_template(
        "construtor.html",
        evento=evento,
        campos=campos,
        atendimentos=atendimentos
    )

@app.route("/atendimento/<int:atendimento_id>/excluir", methods=["POST"])
def excluir_atendimento(atendimento_id):
    conn = get_db()
    atendimento = conn.execute(
        "SELECT evento_id FROM atendimentos WHERE id = ?", (atendimento_id,)
    ).fetchone()
    if atendimento:
        uso = conn.execute(
            "SELECT COUNT(*) FROM inscricoes WHERE atendimento_id = ?", (atendimento_id,)
        ).fetchone()[0]
        evento_id = atendimento["evento_id"]
        if uso > 0:
            flash("Não é possível excluir um atendimento que já possui inscritos.")
        else:
            conn.execute("DELETE FROM atendimentos WHERE id = ?", (atendimento_id,))
            conn.commit()
            flash("Atendimento excluído.")
        conn.close()
        return redirect(url_for("construtor", evento_id=evento_id))
    conn.close()
    return redirect(url_for("index"))

@app.route("/atendimento/<int:atendimento_id>/alternar", methods=["POST"])
def alternar_atendimento(atendimento_id):
    conn = get_db()
    atendimento = conn.execute(
        "SELECT evento_id, ativo FROM atendimentos WHERE id = ?", (atendimento_id,)
    ).fetchone()
    if atendimento:
        conn.execute(
            "UPDATE atendimentos SET ativo = ? WHERE id = ?",
            (0 if atendimento["ativo"] else 1, atendimento_id)
        )
        conn.commit()
        evento_id = atendimento["evento_id"]
        conn.close()
        return redirect(url_for("construtor", evento_id=evento_id))
    conn.close()
    return redirect(url_for("index"))

@app.route("/campo/<int:campo_id>/excluir", methods=["POST"])
def excluir_campo(campo_id):
    conn = get_db()
    campo = conn.execute("SELECT evento_id FROM campos WHERE id = ?", (campo_id,)).fetchone()
    if campo:
        evento_id = campo["evento_id"]
        conn.execute("DELETE FROM campos WHERE id = ?", (campo_id,))
        conn.commit()
        conn.close()
        return redirect(url_for("construtor", evento_id=evento_id))
    conn.close()
    return redirect(url_for("index"))

@app.route("/campo/<int:campo_id>/subir", methods=["POST"])
def subir_campo(campo_id):
    conn = get_db()
    campo = conn.execute("SELECT * FROM campos WHERE id = ?", (campo_id,)).fetchone()
    if campo:
        anterior = conn.execute("""
            SELECT * FROM campos
            WHERE evento_id = ? AND ordem < ?
            ORDER BY ordem DESC LIMIT 1
        """, (campo["evento_id"], campo["ordem"])).fetchone()
        if anterior:
            conn.execute("UPDATE campos SET ordem = ? WHERE id = ?", (anterior["ordem"], campo["id"]))
            conn.execute("UPDATE campos SET ordem = ? WHERE id = ?", (campo["ordem"], anterior["id"]))
            conn.commit()
        evento_id = campo["evento_id"]
        conn.close()
        return redirect(url_for("construtor", evento_id=evento_id))
    conn.close()
    return redirect(url_for("index"))

@app.route("/campo/<int:campo_id>/descer", methods=["POST"])
def descer_campo(campo_id):
    conn = get_db()
    campo = conn.execute("SELECT * FROM campos WHERE id = ?", (campo_id,)).fetchone()
    if campo:
        proximo = conn.execute("""
            SELECT * FROM campos
            WHERE evento_id = ? AND ordem > ?
            ORDER BY ordem ASC LIMIT 1
        """, (campo["evento_id"], campo["ordem"])).fetchone()
        if proximo:
            conn.execute("UPDATE campos SET ordem = ? WHERE id = ?", (proximo["ordem"], campo["id"]))
            conn.execute("UPDATE campos SET ordem = ? WHERE id = ?", (campo["ordem"], proximo["id"]))
            conn.commit()
        evento_id = campo["evento_id"]
        conn.close()
        return redirect(url_for("construtor", evento_id=evento_id))
    conn.close()
    return redirect(url_for("index"))

@app.route("/cadastro/<int:evento_id>", methods=["GET", "POST"])
def cadastro_publico(evento_id):
    conn = get_db()
    evento = conn.execute("SELECT * FROM eventos WHERE id = ?", (evento_id,)).fetchone()
    if not evento:
        conn.close()
        return "Evento não encontrado.", 404
    if not evento["ativo"]:
        conn.close()
        return "Este evento não está aceitando inscrições.", 403

    campos = conn.execute("""
        SELECT * FROM campos WHERE evento_id = ? ORDER BY ordem, id
    """, (evento_id,)).fetchall()

    atendimentos = conn.execute("""
        SELECT a.*,
               (SELECT COUNT(*) FROM inscricoes i WHERE i.atendimento_id = a.id) AS ocupadas
        FROM atendimentos a
        WHERE a.evento_id = ? AND a.ativo = 1
        ORDER BY a.nome
    """, (evento_id,)).fetchall()

    if request.method == "POST":
        atendimento_id = request.form.get("atendimento_id", "").strip()
        if not atendimento_id:
            conn.close()
            flash("Selecione um atendimento.")
            return redirect(url_for("cadastro_publico", evento_id=evento_id))

        atendimento = conn.execute("""
            SELECT a.*,
                   (SELECT COUNT(*) FROM inscricoes i WHERE i.atendimento_id = a.id) AS ocupadas
            FROM atendimentos a
            WHERE a.id = ? AND a.evento_id = ? AND a.ativo = 1
        """, (atendimento_id, evento_id)).fetchone()

        if not atendimento:
            conn.close()
            flash("Atendimento inválido ou indisponível.")
            return redirect(url_for("cadastro_publico", evento_id=evento_id))

        if atendimento["ocupadas"] >= atendimento["vagas"]:
            conn.close()
            flash("As vagas para este atendimento acabaram. Escolha outra opção.")
            return redirect(url_for("cadastro_publico", evento_id=evento_id))

        campo_cpf = next((c for c in campos if c["marcador_cpf"]), None)
        if campo_cpf and not evento["permitir_cpf_repetido"]:
            cpf = request.form.get(f"campo_{campo_cpf['id']}", "").strip()
            if cpf:
                repetido = conn.execute("""
                    SELECT r.id
                    FROM respostas r
                    JOIN inscricoes i ON i.id = r.inscricao_id
                    WHERE i.evento_id = ? AND r.campo_id = ? AND r.valor = ?
                    LIMIT 1
                """, (evento_id, campo_cpf["id"], cpf)).fetchone()
                if repetido:
                    conn.close()
                    flash("Este CPF já foi inscrito neste evento.")
                    return redirect(url_for("cadastro_publico", evento_id=evento_id))

        cur = conn.execute(
            "INSERT INTO inscricoes (evento_id, atendimento_id) VALUES (?, ?)",
            (evento_id, atendimento_id)
        )
        inscricao_id = cur.lastrowid

        for campo in campos:
            valor = request.form.get(f"campo_{campo['id']}", "").strip()
            conn.execute("""
                INSERT INTO respostas (inscricao_id, campo_id, valor)
                VALUES (?, ?, ?)
            """, (inscricao_id, campo["id"], valor))

        conn.commit()
        conn.close()
        return render_template("sucesso.html", evento=evento, atendimento=atendimento)

    conn.close()
    return render_template(
        "cadastro.html",
        evento=evento,
        campos=campos,
        atendimentos=atendimentos
    )

@app.route("/evento/<int:evento_id>/inscritos")
def inscritos(evento_id):
    conn = get_db()
    evento = conn.execute("SELECT * FROM eventos WHERE id = ?", (evento_id,)).fetchone()
    if not evento:
        conn.close()
        return "Evento não encontrado.", 404

    campos = conn.execute("""
        SELECT * FROM campos WHERE evento_id = ? ORDER BY ordem, id
    """, (evento_id,)).fetchall()

    inscricoes = conn.execute("""
        SELECT i.*, a.nome AS atendimento_nome
        FROM inscricoes i
        LEFT JOIN atendimentos a ON a.id = i.atendimento_id
        WHERE i.evento_id = ?
        ORDER BY i.id DESC
    """, (evento_id,)).fetchall()

    linhas = []
    for ins in inscricoes:
        respostas = conn.execute("""
            SELECT campo_id, valor FROM respostas WHERE inscricao_id = ?
        """, (ins["id"],)).fetchall()
        mapa = {r["campo_id"]: r["valor"] for r in respostas}
        linhas.append({
            "id": ins["id"],
            "criado_em": ins["criado_em"],
            "atendimento_nome": ins["atendimento_nome"],
            "respostas": mapa
        })

    atendimentos = conn.execute("""
        SELECT a.*,
               (SELECT COUNT(*) FROM inscricoes i WHERE i.atendimento_id = a.id) AS ocupadas
        FROM atendimentos a
        WHERE a.evento_id = ?
        ORDER BY a.nome
    """, (evento_id,)).fetchall()

    conn.close()
    return render_template(
        "inscritos.html",
        evento=evento,
        campos=campos,
        linhas=linhas,
        atendimentos=atendimentos
    )



@app.route("/evento/<int:evento_id>/voluntarios/construtor", methods=["GET", "POST"])
def construtor_voluntarios(evento_id):
    conn = get_db()
    evento = conn.execute("SELECT * FROM eventos WHERE id = ?", (evento_id,)).fetchone()
    if not evento:
        conn.close()
        return "Evento não encontrado.", 404

    if request.method == "POST":
        titulo = request.form["titulo"].strip()
        tipo = request.form["tipo"]
        obrigatorio = 1 if request.form.get("obrigatorio") == "on" else 0
        marcador_cpf = 1 if request.form.get("marcador_cpf") == "on" else 0
        opcoes = request.form.get("opcoes", "").strip()

        if marcador_cpf:
            conn.execute(
                "UPDATE campos_voluntarios SET marcador_cpf = 0 WHERE evento_id = ?",
                (evento_id,)
            )

        ordem = conn.execute(
            "SELECT COALESCE(MAX(ordem),0)+1 FROM campos_voluntarios WHERE evento_id = ?",
            (evento_id,)
        ).fetchone()[0]

        conn.execute("""
            INSERT INTO campos_voluntarios
            (evento_id, titulo, tipo, obrigatorio, opcoes, ordem, marcador_cpf)
            VALUES (?, ?, ?, ?, ?, ?, ?)
        """, (evento_id, titulo, tipo, obrigatorio, opcoes, ordem, marcador_cpf))
        conn.commit()
        conn.close()
        flash("Campo de voluntário adicionado.")
        return redirect(url_for("construtor_voluntarios", evento_id=evento_id))

    campos = conn.execute("""
        SELECT * FROM campos_voluntarios
        WHERE evento_id = ?
        ORDER BY ordem, id
    """, (evento_id,)).fetchall()
    conn.close()
    return render_template("construtor_voluntarios.html", evento=evento, campos=campos)

@app.route("/campo-voluntario/<int:campo_id>/excluir", methods=["POST"])
def excluir_campo_voluntario(campo_id):
    conn = get_db()
    campo = conn.execute(
        "SELECT evento_id FROM campos_voluntarios WHERE id = ?", (campo_id,)
    ).fetchone()
    if campo:
        evento_id = campo["evento_id"]
        conn.execute("DELETE FROM campos_voluntarios WHERE id = ?", (campo_id,))
        conn.commit()
        conn.close()
        return redirect(url_for("construtor_voluntarios", evento_id=evento_id))
    conn.close()
    return redirect(url_for("index"))

@app.route("/campo-voluntario/<int:campo_id>/subir", methods=["POST"])
def subir_campo_voluntario(campo_id):
    conn = get_db()
    campo = conn.execute(
        "SELECT * FROM campos_voluntarios WHERE id = ?", (campo_id,)
    ).fetchone()
    if campo:
        anterior = conn.execute("""
            SELECT * FROM campos_voluntarios
            WHERE evento_id = ? AND ordem < ?
            ORDER BY ordem DESC LIMIT 1
        """, (campo["evento_id"], campo["ordem"])).fetchone()
        if anterior:
            conn.execute(
                "UPDATE campos_voluntarios SET ordem = ? WHERE id = ?",
                (anterior["ordem"], campo["id"])
            )
            conn.execute(
                "UPDATE campos_voluntarios SET ordem = ? WHERE id = ?",
                (campo["ordem"], anterior["id"])
            )
            conn.commit()
        evento_id = campo["evento_id"]
        conn.close()
        return redirect(url_for("construtor_voluntarios", evento_id=evento_id))
    conn.close()
    return redirect(url_for("index"))

@app.route("/campo-voluntario/<int:campo_id>/descer", methods=["POST"])
def descer_campo_voluntario(campo_id):
    conn = get_db()
    campo = conn.execute(
        "SELECT * FROM campos_voluntarios WHERE id = ?", (campo_id,)
    ).fetchone()
    if campo:
        proximo = conn.execute("""
            SELECT * FROM campos_voluntarios
            WHERE evento_id = ? AND ordem > ?
            ORDER BY ordem ASC LIMIT 1
        """, (campo["evento_id"], campo["ordem"])).fetchone()
        if proximo:
            conn.execute(
                "UPDATE campos_voluntarios SET ordem = ? WHERE id = ?",
                (proximo["ordem"], campo["id"])
            )
            conn.execute(
                "UPDATE campos_voluntarios SET ordem = ? WHERE id = ?",
                (campo["ordem"], proximo["id"])
            )
            conn.commit()
        evento_id = campo["evento_id"]
        conn.close()
        return redirect(url_for("construtor_voluntarios", evento_id=evento_id))
    conn.close()
    return redirect(url_for("index"))

@app.route("/voluntarios/<int:evento_id>", methods=["GET", "POST"])
def cadastro_voluntario_publico(evento_id):
    conn = get_db()
    evento = conn.execute("SELECT * FROM eventos WHERE id = ?", (evento_id,)).fetchone()
    if not evento:
        conn.close()
        return "Evento não encontrado.", 404
    if not evento["ativo"]:
        conn.close()
        return "Este evento não está aceitando cadastros.", 403

    campos = conn.execute("""
        SELECT * FROM campos_voluntarios
        WHERE evento_id = ?
        ORDER BY ordem, id
    """, (evento_id,)).fetchall()

    if request.method == "POST":
        campo_cpf = next((c for c in campos if c["marcador_cpf"]), None)
        if campo_cpf and not evento["permitir_cpf_repetido"]:
            cpf = request.form.get(f"campo_{campo_cpf['id']}", "").strip()
            if cpf:
                repetido = conn.execute("""
                    SELECT rv.id
                    FROM respostas_voluntarios rv
                    JOIN voluntarios v ON v.id = rv.voluntario_id
                    WHERE v.evento_id = ? AND rv.campo_id = ? AND rv.valor = ?
                    LIMIT 1
                """, (evento_id, campo_cpf["id"], cpf)).fetchone()
                if repetido:
                    conn.close()
                    flash("Este CPF já foi cadastrado como voluntário neste evento.")
                    return redirect(url_for("cadastro_voluntario_publico", evento_id=evento_id))

        cur = conn.execute(
            "INSERT INTO voluntarios (evento_id) VALUES (?)",
            (evento_id,)
        )
        voluntario_id = cur.lastrowid

        for campo in campos:
            valor = request.form.get(f"campo_{campo['id']}", "").strip()
            conn.execute("""
                INSERT INTO respostas_voluntarios
                (voluntario_id, campo_id, valor)
                VALUES (?, ?, ?)
            """, (voluntario_id, campo["id"], valor))

        conn.commit()
        conn.close()
        return render_template("sucesso_voluntario.html", evento=evento)

    conn.close()
    return render_template(
        "cadastro_voluntario.html",
        evento=evento,
        campos=campos
    )

@app.route("/evento/<int:evento_id>/ver-voluntarios")
def ver_voluntarios(evento_id):
    conn = get_db()
    evento = conn.execute("SELECT * FROM eventos WHERE id = ?", (evento_id,)).fetchone()
    if not evento:
        conn.close()
        return "Evento não encontrado.", 404

    campos = conn.execute("""
        SELECT * FROM campos_voluntarios
        WHERE evento_id = ?
        ORDER BY ordem, id
    """, (evento_id,)).fetchall()

    voluntarios = conn.execute("""
        SELECT * FROM voluntarios
        WHERE evento_id = ?
        ORDER BY id DESC
    """, (evento_id,)).fetchall()

    linhas = []
    for vol in voluntarios:
        respostas = conn.execute("""
            SELECT campo_id, valor
            FROM respostas_voluntarios
            WHERE voluntario_id = ?
        """, (vol["id"],)).fetchall()
        mapa = {r["campo_id"]: r["valor"] for r in respostas}
        linhas.append({
            "id": vol["id"],
            "criado_em": vol["criado_em"],
            "respostas": mapa
        })

    conn.close()
    return render_template(
        "voluntarios.html",
        evento=evento,
        campos=campos,
        linhas=linhas
    )

@app.route("/evento/<int:evento_id>/exportar-voluntarios")
def exportar_voluntarios(evento_id):
    conn = get_db()
    evento = conn.execute("SELECT * FROM eventos WHERE id = ?", (evento_id,)).fetchone()
    if not evento:
        conn.close()
        return "Evento não encontrado.", 404

    campos = conn.execute("""
        SELECT * FROM campos_voluntarios
        WHERE evento_id = ?
        ORDER BY ordem, id
    """, (evento_id,)).fetchall()

    voluntarios = conn.execute("""
        SELECT * FROM voluntarios
        WHERE evento_id = ?
        ORDER BY id
    """, (evento_id,)).fetchall()

    output = io.BytesIO()
    workbook = xlsxwriter.Workbook(output, {"in_memory": True})
    ws = workbook.add_worksheet("Voluntários")

    fmt_titulo = workbook.add_format({
        "bold": True, "font_size": 16, "align": "center",
        "valign": "vcenter", "bg_color": "#1F4D2E", "font_color": "#FFFFFF"
    })
    fmt_header = workbook.add_format({
        "bold": True, "bg_color": "#2F6B43", "font_color": "#FFFFFF",
        "border": 1, "align": "center", "text_wrap": True
    })
    fmt_cell = workbook.add_format({"border": 1, "valign": "top", "text_wrap": True})

    headers = ["ID"] + [c["titulo"] for c in campos] + ["Data/Hora do cadastro"]
    total_cols = len(headers)

    ws.merge_range(0, 0, 0, total_cols - 1, f"Voluntários — {evento['nome']}", fmt_titulo)

    for col, header in enumerate(headers):
        ws.write(2, col, header, fmt_header)

    row = 3
    for vol in voluntarios:
        respostas = conn.execute("""
            SELECT campo_id, valor
            FROM respostas_voluntarios
            WHERE voluntario_id = ?
        """, (vol["id"],)).fetchall()
        mapa = {r["campo_id"]: r["valor"] for r in respostas}
        valores = [vol["id"]] + [mapa.get(c["id"], "") for c in campos] + [vol["criado_em"] or ""]
        for col, valor in enumerate(valores):
            ws.write(row, col, valor, fmt_cell)
        row += 1

    ws.freeze_panes(3, 0)
    ws.autofilter(2, 0, max(2, row - 1), total_cols - 1)
    ws.set_column(0, 0, 8)
    if campos:
        ws.set_column(1, len(campos), 22)
    ws.set_column(total_cols - 1, total_cols - 1, 22)

    workbook.close()
    conn.close()
    output.seek(0)

    nome_seguro = "".join(
        c if c.isalnum() or c in (" ", "-", "_") else ""
        for c in evento["nome"]
    ).strip().replace(" ", "_") or f"evento_{evento_id}"

    return send_file(
        output,
        as_attachment=True,
        download_name=f"voluntarios_{nome_seguro}.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


@app.route("/evento/<int:evento_id>/exportar-excel")
def exportar_excel(evento_id):
    conn = get_db()

    evento = conn.execute(
        "SELECT * FROM eventos WHERE id = ?", (evento_id,)
    ).fetchone()
    if not evento:
        conn.close()
        return "Evento não encontrado.", 404

    campos = conn.execute("""
        SELECT * FROM campos
        WHERE evento_id = ?
        ORDER BY ordem, id
    """, (evento_id,)).fetchall()

    inscricoes = conn.execute("""
        SELECT i.*, a.nome AS atendimento_nome
        FROM inscricoes i
        LEFT JOIN atendimentos a ON a.id = i.atendimento_id
        WHERE i.evento_id = ?
        ORDER BY i.id
    """, (evento_id,)).fetchall()

    atendimentos = conn.execute("""
        SELECT a.*,
               (SELECT COUNT(*) FROM inscricoes i WHERE i.atendimento_id = a.id) AS ocupadas
        FROM atendimentos a
        WHERE a.evento_id = ?
        ORDER BY a.nome
    """, (evento_id,)).fetchall()

    output = io.BytesIO()
    workbook = xlsxwriter.Workbook(output, {"in_memory": True})

    # Formatações
    fmt_titulo = workbook.add_format({
        "bold": True, "font_size": 16, "align": "center",
        "valign": "vcenter", "bg_color": "#1F4D2E", "font_color": "#FFFFFF"
    })
    fmt_subtitulo = workbook.add_format({
        "italic": True, "align": "center", "font_color": "#555555"
    })
    fmt_header = workbook.add_format({
        "bold": True, "bg_color": "#2F6B43", "font_color": "#FFFFFF",
        "border": 1, "align": "center", "valign": "vcenter", "text_wrap": True
    })
    fmt_cell = workbook.add_format({
        "border": 1, "valign": "top", "text_wrap": True
    })
    fmt_center = workbook.add_format({
        "border": 1, "align": "center", "valign": "vcenter"
    })
    fmt_resumo_header = workbook.add_format({
        "bold": True, "bg_color": "#2F6B43", "font_color": "#FFFFFF",
        "border": 1, "align": "center"
    })
    fmt_resumo = workbook.add_format({"border": 1, "align": "center"})

    # Aba Inscritos
    ws = workbook.add_worksheet("Inscritos")

    headers = ["ID", "Atendimento"] + [c["titulo"] for c in campos] + ["Data/Hora da inscrição"]
    total_cols = len(headers)

    ws.merge_range(0, 0, 0, total_cols - 1, evento["nome"], fmt_titulo)

    info = []
    if evento["data"]:
        info.append(f"Data: {evento['data']}")
    if evento["local"]:
        info.append(f"Local: {evento['local']}")
    info_text = " | ".join(info) if info else "Relatório de inscritos"
    ws.merge_range(1, 0, 1, total_cols - 1, info_text, fmt_subtitulo)

    for col, header in enumerate(headers):
        ws.write(3, col, header, fmt_header)

    row = 4
    for ins in inscricoes:
        respostas = conn.execute("""
            SELECT campo_id, valor
            FROM respostas
            WHERE inscricao_id = ?
        """, (ins["id"],)).fetchall()
        mapa = {r["campo_id"]: r["valor"] for r in respostas}

        valores = [
            ins["id"],
            ins["atendimento_nome"] or ""
        ]
        valores += [mapa.get(c["id"], "") for c in campos]
        valores += [ins["criado_em"] or ""]

        for col, valor in enumerate(valores):
            if col == 0:
                ws.write(row, col, valor, fmt_center)
            else:
                ws.write(row, col, valor, fmt_cell)
        row += 1

    ws.freeze_panes(4, 0)
    ws.autofilter(3, 0, max(3, row - 1), total_cols - 1)
    ws.set_row(0, 28)
    ws.set_row(3, 32)
    ws.set_column(0, 0, 8)
    ws.set_column(1, 1, 22)
    if campos:
        ws.set_column(2, 1 + len(campos), 22)
    ws.set_column(total_cols - 1, total_cols - 1, 22)

    # Aba Resumo de vagas
    resumo = workbook.add_worksheet("Resumo de Vagas")
    resumo_headers = ["Atendimento", "Total de vagas", "Ocupadas", "Restantes"]
    for col, header in enumerate(resumo_headers):
        resumo.write(0, col, header, fmt_resumo_header)

    r = 1
    for a in atendimentos:
        restantes = max(0, a["vagas"] - a["ocupadas"])
        resumo.write(r, 0, a["nome"], fmt_cell)
        resumo.write(r, 1, a["vagas"], fmt_resumo)
        resumo.write(r, 2, a["ocupadas"], fmt_resumo)
        resumo.write(r, 3, restantes, fmt_resumo)
        r += 1

    resumo.set_column(0, 0, 28)
    resumo.set_column(1, 3, 16)
    resumo.freeze_panes(1, 0)

    workbook.close()
    conn.close()

    output.seek(0)

    nome_seguro = "".join(
        c if c.isalnum() or c in (" ", "-", "_") else ""
        for c in evento["nome"]
    ).strip().replace(" ", "_")
    if not nome_seguro:
        nome_seguro = f"evento_{evento_id}"

    return send_file(
        output,
        as_attachment=True,
        download_name=f"inscritos_{nome_seguro}.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


@app.route("/evento/<int:evento_id>/alternar", methods=["POST"])
def alternar_evento(evento_id):
    conn = get_db()
    evento = conn.execute("SELECT ativo FROM eventos WHERE id = ?", (evento_id,)).fetchone()
    if evento:
        conn.execute(
            "UPDATE eventos SET ativo = ? WHERE id = ?",
            (0 if evento["ativo"] else 1, evento_id)
        )
        conn.commit()
    conn.close()
    return redirect(url_for("index"))

if __name__ == "__main__":
    init_db()
    app.run(debug=True, host="127.0.0.1", port=5000)
