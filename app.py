from flask import Flask, render_template, request, redirect, url_for, flash, send_file, session, abort
import os, io, sqlite3, xlsxwriter
from pathlib import Path
from functools import wraps
from werkzeug.security import generate_password_hash, check_password_hash

try:
    import psycopg2
    from psycopg2.extras import DictCursor
except ImportError:
    psycopg2 = None
    DictCursor = None

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "chama-social-desenvolvimento-local")
DATABASE_URL = os.environ.get("DATABASE_URL", "").strip()
DB_PATH = Path(__file__).parent / "chama_social.db"


class CursorWrapper:
    def __init__(self, cursor, lastrowid=None):
        self.cursor = cursor
        self.lastrowid = lastrowid

    def fetchone(self):
        return self.cursor.fetchone()

    def fetchall(self):
        return self.cursor.fetchall()


class PostgresConnection:
    def __init__(self):
        if psycopg2 is None:
            raise RuntimeError("psycopg2 não está instalado.")
        self.conn = psycopg2.connect(DATABASE_URL, sslmode="require", cursor_factory=DictCursor)

    def execute(self, sql, params=()):
        sql = sql.replace("?", "%s")
        cur = self.conn.cursor()
        lastrowid = None
        if sql.lstrip().upper().startswith("INSERT INTO"):
            sql_exec = sql.rstrip().rstrip(";")
            if " RETURNING " not in sql_exec.upper():
                sql_exec += " RETURNING id"
            cur.execute(sql_exec, params)
            row = cur.fetchone()
            if row is not None:
                lastrowid = row["id"]
        else:
            cur.execute(sql, params)
        return CursorWrapper(cur, lastrowid)

    def commit(self):
        self.conn.commit()

    def rollback(self):
        self.conn.rollback()

    def close(self):
        self.conn.close()


def using_postgres():
    return bool(DATABASE_URL)


def get_db():
    if using_postgres():
        return PostgresConnection()
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    conn = get_db()
    try:
        if using_postgres():
            statements = [
                """CREATE TABLE IF NOT EXISTS usuarios (
                    id SERIAL PRIMARY KEY,
                    nome TEXT NOT NULL,
                    login TEXT NOT NULL UNIQUE,
                    senha_hash TEXT NOT NULL,
                    perfil TEXT NOT NULL CHECK(perfil IN ('admin','operador','visualizador')),
                    ativo INTEGER NOT NULL DEFAULT 1,
                    trocar_senha INTEGER NOT NULL DEFAULT 1,
                    criado_em TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )""",
                """CREATE TABLE IF NOT EXISTS eventos (
                    id SERIAL PRIMARY KEY,
                    nome TEXT NOT NULL,
                    data TEXT,
                    local TEXT,
                    permitir_cpf_repetido INTEGER NOT NULL DEFAULT 0,
                    ativo INTEGER NOT NULL DEFAULT 1,
                    criado_por INTEGER REFERENCES usuarios(id)
                )""",
                """CREATE TABLE IF NOT EXISTS campos (
                    id SERIAL PRIMARY KEY,
                    evento_id INTEGER NOT NULL REFERENCES eventos(id),
                    titulo TEXT NOT NULL,
                    tipo TEXT NOT NULL,
                    obrigatorio INTEGER NOT NULL DEFAULT 0,
                    opcoes TEXT,
                    ordem INTEGER NOT NULL DEFAULT 0,
                    marcador_cpf INTEGER NOT NULL DEFAULT 0
                )""",
                """CREATE TABLE IF NOT EXISTS atendimentos (
                    id SERIAL PRIMARY KEY,
                    evento_id INTEGER NOT NULL REFERENCES eventos(id),
                    nome TEXT NOT NULL,
                    vagas INTEGER NOT NULL DEFAULT 0,
                    ativo INTEGER NOT NULL DEFAULT 1
                )""",
                """CREATE TABLE IF NOT EXISTS inscricoes (
                    id SERIAL PRIMARY KEY,
                    evento_id INTEGER NOT NULL REFERENCES eventos(id),
                    atendimento_id INTEGER REFERENCES atendimentos(id),
                    criado_em TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )""",
                """CREATE TABLE IF NOT EXISTS respostas (
                    id SERIAL PRIMARY KEY,
                    inscricao_id INTEGER NOT NULL REFERENCES inscricoes(id),
                    campo_id INTEGER NOT NULL REFERENCES campos(id),
                    valor TEXT
                )""",
                """CREATE TABLE IF NOT EXISTS campos_voluntarios (
                    id SERIAL PRIMARY KEY,
                    evento_id INTEGER NOT NULL REFERENCES eventos(id),
                    titulo TEXT NOT NULL,
                    tipo TEXT NOT NULL,
                    obrigatorio INTEGER NOT NULL DEFAULT 0,
                    opcoes TEXT,
                    ordem INTEGER NOT NULL DEFAULT 0
                )""",
                """CREATE TABLE IF NOT EXISTS funcoes_voluntarios (
                    id SERIAL PRIMARY KEY,
                    evento_id INTEGER NOT NULL REFERENCES eventos(id),
                    nome TEXT NOT NULL,
                    vagas INTEGER NOT NULL DEFAULT 0,
                    ativo INTEGER NOT NULL DEFAULT 1
                )""",
                """CREATE TABLE IF NOT EXISTS voluntarios (
                    id SERIAL PRIMARY KEY,
                    evento_id INTEGER NOT NULL REFERENCES eventos(id),
                    funcao_id INTEGER REFERENCES funcoes_voluntarios(id),
                    criado_em TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )""",
                """CREATE TABLE IF NOT EXISTS respostas_voluntarios (
                    id SERIAL PRIMARY KEY,
                    voluntario_id INTEGER NOT NULL REFERENCES voluntarios(id),
                    campo_id INTEGER NOT NULL REFERENCES campos_voluntarios(id),
                    valor TEXT
                )""",
                """CREATE TABLE IF NOT EXISTS solicitacoes_exclusao (
                    id SERIAL PRIMARY KEY,
                    usuario_id INTEGER NOT NULL,
                    tipo TEXT NOT NULL,
                    alvo_id INTEGER NOT NULL,
                    descricao TEXT NOT NULL,
                    status TEXT NOT NULL DEFAULT 'pendente',
                    criado_em TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )""",
            ]
            for statement in statements:
                conn.execute(statement)
        else:
            conn.executescript('''
            CREATE TABLE IF NOT EXISTS usuarios (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                nome TEXT NOT NULL,
                login TEXT NOT NULL UNIQUE,
                senha_hash TEXT NOT NULL,
                perfil TEXT NOT NULL CHECK(perfil IN ('admin','operador','visualizador')),
                ativo INTEGER NOT NULL DEFAULT 1,
                trocar_senha INTEGER NOT NULL DEFAULT 1,
                criado_em DATETIME DEFAULT CURRENT_TIMESTAMP
            );
            CREATE TABLE IF NOT EXISTS eventos (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                nome TEXT NOT NULL,
                data TEXT,
                local TEXT,
                permitir_cpf_repetido INTEGER NOT NULL DEFAULT 0,
                ativo INTEGER NOT NULL DEFAULT 1,
                criado_por INTEGER,
                FOREIGN KEY(criado_por) REFERENCES usuarios(id)
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
            CREATE TABLE IF NOT EXISTS respostas (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                inscricao_id INTEGER NOT NULL,
                campo_id INTEGER NOT NULL,
                valor TEXT,
                FOREIGN KEY(inscricao_id) REFERENCES inscricoes(id),
                FOREIGN KEY(campo_id) REFERENCES campos(id)
            );
            CREATE TABLE IF NOT EXISTS campos_voluntarios (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                evento_id INTEGER NOT NULL,
                titulo TEXT NOT NULL,
                tipo TEXT NOT NULL,
                obrigatorio INTEGER NOT NULL DEFAULT 0,
                opcoes TEXT,
                ordem INTEGER NOT NULL DEFAULT 0,
                FOREIGN KEY(evento_id) REFERENCES eventos(id)
            );
            CREATE TABLE IF NOT EXISTS funcoes_voluntarios (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                evento_id INTEGER NOT NULL,
                nome TEXT NOT NULL,
                vagas INTEGER NOT NULL DEFAULT 0,
                ativo INTEGER NOT NULL DEFAULT 1,
                FOREIGN KEY(evento_id) REFERENCES eventos(id)
            );
            CREATE TABLE IF NOT EXISTS voluntarios (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                evento_id INTEGER NOT NULL,
                funcao_id INTEGER,
                criado_em DATETIME DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY(evento_id) REFERENCES eventos(id),
                FOREIGN KEY(funcao_id) REFERENCES funcoes_voluntarios(id)
            );
            CREATE TABLE IF NOT EXISTS respostas_voluntarios (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                voluntario_id INTEGER NOT NULL,
                campo_id INTEGER NOT NULL,
                valor TEXT,
                FOREIGN KEY(voluntario_id) REFERENCES voluntarios(id),
                FOREIGN KEY(campo_id) REFERENCES campos_voluntarios(id)
            );
            CREATE TABLE IF NOT EXISTS solicitacoes_exclusao (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                usuario_id INTEGER NOT NULL,
                tipo TEXT NOT NULL,
                alvo_id INTEGER NOT NULL,
                descricao TEXT NOT NULL,
                status TEXT NOT NULL DEFAULT 'pendente',
                criado_em DATETIME DEFAULT CURRENT_TIMESTAMP
            );
            ''')

        # Migração para bancos criados em versões anteriores.
        if using_postgres():
            conn.execute("ALTER TABLE voluntarios ADD COLUMN IF NOT EXISTS funcao_id INTEGER REFERENCES funcoes_voluntarios(id)")
        else:
            cols = conn.execute("PRAGMA table_info(voluntarios)").fetchall()
            if not any(c[1] == 'funcao_id' for c in cols):
                conn.execute("ALTER TABLE voluntarios ADD COLUMN funcao_id INTEGER")

        existe = conn.execute("SELECT id FROM usuarios LIMIT 1").fetchone()
        if not existe:
            conn.execute(
                "INSERT INTO usuarios (nome, login, senha_hash, perfil, trocar_senha) VALUES (?,?,?,?,0)",
                ('Administrador', 'Amorim', 'pbkdf2:sha256:600000$chamasocial-amorim-2026$e401e9cd651ae801f5b7d946e98d99e5790c49142c5675ef877de1ffbd32bcf4', 'admin')
            )
        conn.commit()
    except Exception:
        if hasattr(conn, "rollback"):
            conn.rollback()
        raise
    finally:
        conn.close()


def usuario_atual():
    uid = session.get('usuario_id')
    if not uid: return None
    conn = get_db(); u = conn.execute('SELECT * FROM usuarios WHERE id=? AND ativo=1',(uid,)).fetchone(); conn.close()
    return u


def login_required(fn):
    @wraps(fn)
    def inner(*args, **kwargs):
        u = usuario_atual()
        if not u:
            return redirect(url_for('login'))
        if u['trocar_senha'] and request.endpoint not in ('trocar_senha','logout','static'):
            return redirect(url_for('trocar_senha'))
        return fn(*args, **kwargs)
    return inner


def perfis(*permitidos):
    def deco(fn):
        @wraps(fn)
        @login_required
        def inner(*args, **kwargs):
            u=usuario_atual()
            if u['perfil'] not in permitidos:
                flash('Seu perfil não tem permissão para esta ação.','erro')
                return redirect(url_for('index'))
            return fn(*args, **kwargs)
        return inner
    return deco

@app.context_processor
def inject_user():
    return {'usuario': usuario_atual()}

@app.route('/login', methods=['GET','POST'])
def login():
    if request.method=='POST':
        login_nome=request.form.get('login','').strip()
        senha=request.form.get('senha','')
        conn=get_db(); u=conn.execute('SELECT * FROM usuarios WHERE login=?',(login_nome,)).fetchone(); conn.close()
        if not u or not u['ativo'] or not check_password_hash(u['senha_hash'], senha):
            flash('Login ou senha inválidos.','erro')
        else:
            session.clear(); session['usuario_id']=u['id']
            return redirect(url_for('trocar_senha' if u['trocar_senha'] else 'index'))
    return render_template('login.html')

@app.route('/logout')
def logout():
    session.clear(); return redirect(url_for('login'))

@app.route('/trocar-senha', methods=['GET','POST'])
@login_required
def trocar_senha():
    u=usuario_atual()
    if request.method=='POST':
        atual=request.form.get('atual',''); nova=request.form.get('nova',''); confirmar=request.form.get('confirmar','')
        if not check_password_hash(u['senha_hash'], atual): flash('Senha atual incorreta.','erro')
        elif len(nova)<1: flash('Digite uma nova senha.','erro')
        elif nova!=confirmar: flash('As senhas não coincidem.','erro')
        else:
            conn=get_db(); conn.execute('UPDATE usuarios SET senha_hash=?, trocar_senha=0 WHERE id=?',(generate_password_hash(nova),u['id'])); conn.commit(); conn.close()
            flash('Senha alterada com sucesso.','ok'); return redirect(url_for('index'))
    return render_template('trocar_senha.html', obrigatoria=bool(u['trocar_senha']))

@app.route('/')
@login_required
def index():
    conn=get_db()
    eventos=conn.execute('''SELECT e.*, (SELECT COUNT(*) FROM inscricoes i WHERE i.evento_id=e.id) total_inscricoes,
    (SELECT COUNT(*) FROM voluntarios v WHERE v.evento_id=e.id) total_voluntarios,
    (SELECT COUNT(*) FROM atendimentos a WHERE a.evento_id=e.id) total_atendimentos FROM eventos e ORDER BY e.id DESC''').fetchall()
    conn.close(); return render_template('index.html', eventos=eventos)

@app.route('/usuarios')
@perfis('admin')
def usuarios():
    conn=get_db(); us=conn.execute('SELECT * FROM usuarios ORDER BY nome').fetchall(); conn.close()
    return render_template('usuarios.html', usuarios=us)

@app.route('/usuarios/novo', methods=['GET','POST'])
@perfis('admin')
def novo_usuario():
    if request.method=='POST':
        nome=request.form['nome'].strip(); login_nome=request.form['login'].strip(); perfil=request.form['perfil']; senha=request.form.get('senha','')
        if not senha:
            flash('Defina uma senha para o novo usuário.','erro')
            return render_template('novo_usuario.html')
        conn=None
        try:
            conn=get_db(); conn.execute('INSERT INTO usuarios (nome,login,senha_hash,perfil,trocar_senha) VALUES (?,?,?,?,1)',(nome,login_nome,generate_password_hash(senha),perfil)); conn.commit(); conn.close(); conn=None
            flash('Usuário criado. No primeiro acesso ele deverá criar uma nova senha.','ok'); return redirect(url_for('usuarios'))
        except Exception as exc:
            if conn:
                if hasattr(conn,'rollback'): conn.rollback()
                conn.close()
            # Em ambos os bancos, login duplicado é tratado com a mesma mensagem.
            if (using_postgres() and psycopg2 and isinstance(exc, psycopg2.IntegrityError)) or isinstance(exc, sqlite3.IntegrityError):
                flash('Esse login já existe.','erro')
            else:
                raise
    return render_template('novo_usuario.html')

@app.route('/usuarios/<int:uid>/resetar', methods=['POST'])
@perfis('admin')
def resetar_usuario(uid):
    nova=request.form.get('nova_senha','')
    if not nova:
        flash('Digite a nova senha antes de redefinir.','erro')
        return redirect(url_for('usuarios'))
    conn=get_db(); conn.execute('UPDATE usuarios SET senha_hash=?, trocar_senha=1 WHERE id=?',(generate_password_hash(nova),uid)); conn.commit(); conn.close()
    flash('Senha redefinida. O usuário terá de criar uma nova senha no próximo acesso.','ok'); return redirect(url_for('usuarios'))

@app.route('/usuarios/<int:uid>/alternar', methods=['POST'])
@perfis('admin')
def alternar_usuario(uid):
    if uid==session.get('usuario_id'):
        flash('Você não pode desativar o próprio usuário.','erro'); return redirect(url_for('usuarios'))
    conn=get_db(); u=conn.execute('SELECT ativo FROM usuarios WHERE id=?',(uid,)).fetchone()
    if u: conn.execute('UPDATE usuarios SET ativo=? WHERE id=?',(0 if u['ativo'] else 1,uid)); conn.commit()
    conn.close(); return redirect(url_for('usuarios'))

@app.route('/evento/novo', methods=['GET','POST'])
@perfis('admin','operador')
def novo_evento():
    if request.method=='POST':
        nome=request.form['nome'].strip(); data=request.form.get('data',''); local=request.form.get('local',''); permitir=1 if request.form.get('permitir_cpf_repetido') else 0
        conn=get_db(); cur=conn.execute('INSERT INTO eventos (nome,data,local,permitir_cpf_repetido,criado_por) VALUES (?,?,?,?,?)',(nome,data,local,permitir,session['usuario_id'])); eid=cur.lastrowid; conn.commit(); conn.close()
        flash('Evento criado. Agora configure Atendimentos e Voluntários separadamente.','ok')
        return redirect(url_for('index'))
    return render_template('novo_evento.html')

@app.route('/evento/<int:evento_id>/construtor', methods=['GET','POST'])
@perfis('admin','operador')
def construtor(evento_id):
    # Rota antiga mantida apenas para compatibilidade.
    # A configuração foi separada em dois módulos independentes.
    flash('A configuração agora está separada em Atendimentos e Voluntários.','ok')
    return redirect(url_for('index'))

@app.route('/evento/<int:evento_id>/cadastro', methods=['GET','POST'])
def cadastro_publico(evento_id):
    conn=get_db(); evento=conn.execute('SELECT * FROM eventos WHERE id=?',(evento_id,)).fetchone()
    if not evento or not evento['ativo']: conn.close(); return render_template('publico_mensagem.html',mensagem='Este cadastro não está disponível no momento.')
    campos=conn.execute('SELECT * FROM campos WHERE evento_id=? ORDER BY ordem,id',(evento_id,)).fetchall(); at=conn.execute('SELECT a.*, (SELECT COUNT(*) FROM inscricoes i WHERE i.atendimento_id=a.id) ocupadas FROM atendimentos a WHERE evento_id=? AND ativo=1 ORDER BY nome',(evento_id,)).fetchall()
    if request.method=='POST':
        aid=int(request.form.get('atendimento_id',0)); atendimento=conn.execute('SELECT a.*, (SELECT COUNT(*) FROM inscricoes i WHERE i.atendimento_id=a.id) ocupadas FROM atendimentos a WHERE id=? AND evento_id=? AND ativo=1',(aid,evento_id)).fetchone()
        if not atendimento or atendimento['ocupadas']>=atendimento['vagas']:
            conn.close(); return render_template('publico_mensagem.html',mensagem='As vagas para esse atendimento acabaram.')
        cpf_campo=next((c for c in campos if c['marcador_cpf']),None)
        if cpf_campo and not evento['permitir_cpf_repetido']:
            cpf=request.form.get(f"campo_{cpf_campo['id']}",'').strip()
            if cpf and conn.execute('''SELECT 1 FROM respostas r JOIN inscricoes i ON i.id=r.inscricao_id WHERE i.evento_id=? AND r.campo_id=? AND r.valor=? LIMIT 1''',(evento_id,cpf_campo['id'],cpf)).fetchone():
                conn.close(); return render_template('publico_mensagem.html',mensagem='Este CPF já foi inscrito neste evento.')
        cur=conn.execute('INSERT INTO inscricoes (evento_id,atendimento_id) VALUES (?,?)',(evento_id,aid)); iid=cur.lastrowid
        for c in campos: conn.execute('INSERT INTO respostas (inscricao_id,campo_id,valor) VALUES (?,?,?)',(iid,c['id'],request.form.get(f"campo_{c['id']}",'')))
        conn.commit(); conn.close(); return render_template('sucesso.html',titulo='Inscrição realizada!',texto='Seu cadastro foi recebido com sucesso.')
    conn.close(); return render_template('cadastro_publico.html',evento=evento,campos=campos,atendimentos=at)

@app.route('/evento/<int:evento_id>/voluntario', methods=['GET','POST'])
def cadastro_voluntario(evento_id):
    conn=get_db(); evento=conn.execute('SELECT * FROM eventos WHERE id=?',(evento_id,)).fetchone(); campos=conn.execute('SELECT * FROM campos_voluntarios WHERE evento_id=? ORDER BY ordem,id',(evento_id,)).fetchall()
    funcoes=conn.execute('SELECT f.*, (SELECT COUNT(*) FROM voluntarios v WHERE v.funcao_id=f.id) ocupadas FROM funcoes_voluntarios f WHERE evento_id=? AND ativo=1 ORDER BY nome',(evento_id,)).fetchall()
    if not evento or not evento['ativo']: conn.close(); return render_template('publico_mensagem.html',mensagem='Este cadastro não está disponível no momento.')
    if request.method=='POST':
        fid=int(request.form.get('funcao_id',0) or 0)
        funcao=conn.execute('SELECT f.*, (SELECT COUNT(*) FROM voluntarios v WHERE v.funcao_id=f.id) ocupadas FROM funcoes_voluntarios f WHERE id=? AND evento_id=? AND ativo=1',(fid,evento_id)).fetchone()
        if not funcao or funcao['ocupadas']>=funcao['vagas']:
            conn.close(); return render_template('publico_mensagem.html',mensagem='As vagas para essa função de voluntário acabaram.')
        cur=conn.execute('INSERT INTO voluntarios (evento_id,funcao_id) VALUES (?,?)',(evento_id,fid)); vid=cur.lastrowid
        for c in campos: conn.execute('INSERT INTO respostas_voluntarios (voluntario_id,campo_id,valor) VALUES (?,?,?)',(vid,c['id'],request.form.get(f"campo_{c['id']}",'')))
        conn.commit(); conn.close(); return render_template('sucesso.html',titulo='Cadastro realizado!',texto='Cadastro de voluntário recebido com sucesso.')
    conn.close(); return render_template('cadastro_voluntario.html',evento=evento,campos=campos,funcoes=funcoes)

@app.route('/evento/<int:evento_id>/inscritos')
@login_required
def inscritos(evento_id):
    conn=get_db(); evento=conn.execute('SELECT * FROM eventos WHERE id=?',(evento_id,)).fetchone(); campos=conn.execute('SELECT * FROM campos WHERE evento_id=? ORDER BY ordem,id',(evento_id,)).fetchall(); ins=conn.execute('SELECT i.*,a.nome atendimento_nome FROM inscricoes i LEFT JOIN atendimentos a ON a.id=i.atendimento_id WHERE i.evento_id=? ORDER BY i.id DESC',(evento_id,)).fetchall(); linhas=[]
    for i in ins:
        mapa={r['campo_id']:r['valor'] for r in conn.execute('SELECT campo_id,valor FROM respostas WHERE inscricao_id=?',(i['id'],)).fetchall()}; linhas.append((i,mapa))
    conn.close(); return render_template('inscritos.html',evento=evento,campos=campos,linhas=linhas)

@app.route('/evento/<int:evento_id>/voluntarios')
@login_required
def voluntarios(evento_id):
    conn=get_db(); evento=conn.execute('SELECT * FROM eventos WHERE id=?',(evento_id,)).fetchone(); campos=conn.execute('SELECT * FROM campos_voluntarios WHERE evento_id=? ORDER BY ordem,id',(evento_id,)).fetchall(); vs=conn.execute('SELECT v.*,f.nome funcao_nome FROM voluntarios v LEFT JOIN funcoes_voluntarios f ON f.id=v.funcao_id WHERE v.evento_id=? ORDER BY v.id DESC',(evento_id,)).fetchall(); linhas=[]
    for v in vs:
        mapa={r['campo_id']:r['valor'] for r in conn.execute('SELECT campo_id,valor FROM respostas_voluntarios WHERE voluntario_id=?',(v['id'],)).fetchall()}; linhas.append((v,mapa))
    conn.close(); return render_template('voluntarios.html',evento=evento,campos=campos,linhas=linhas)

@app.route('/evento/<int:evento_id>/exportar')
@login_required
def exportar(evento_id):
    conn=get_db(); evento=conn.execute('SELECT * FROM eventos WHERE id=?',(evento_id,)).fetchone(); campos=conn.execute('SELECT * FROM campos WHERE evento_id=? ORDER BY ordem,id',(evento_id,)).fetchall(); ins=conn.execute('SELECT i.*,a.nome atendimento_nome FROM inscricoes i LEFT JOIN atendimentos a ON a.id=i.atendimento_id WHERE i.evento_id=? ORDER BY i.id',(evento_id,)).fetchall(); out=io.BytesIO(); wb=xlsxwriter.Workbook(out,{'in_memory':True}); ws=wb.add_worksheet('Inscritos'); headers=['ID','Atendimento']+[c['titulo'] for c in campos]+['Data/Hora']
    for col,h in enumerate(headers): ws.write(0,col,h)
    for row,i in enumerate(ins,1):
        mapa={r['campo_id']:r['valor'] for r in conn.execute('SELECT campo_id,valor FROM respostas WHERE inscricao_id=?',(i['id'],)).fetchall()}; vals=[i['id'],i['atendimento_nome'] or '']+[mapa.get(c['id'],'') for c in campos]+[str(i['criado_em'] or '')]
        for col,v in enumerate(vals): ws.write(row,col,v)
    wb.close(); conn.close(); out.seek(0); return send_file(out,as_attachment=True,download_name=f"inscritos_evento_{evento_id}.xlsx",mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

@app.route('/evento/<int:evento_id>/alternar', methods=['POST'])
@perfis('admin','operador')
def alternar_evento(evento_id):
    conn=get_db(); e=conn.execute('SELECT ativo FROM eventos WHERE id=?',(evento_id,)).fetchone()
    if e: conn.execute('UPDATE eventos SET ativo=? WHERE id=?',(0 if e['ativo'] else 1,evento_id)); conn.commit()
    conn.close(); return redirect(url_for('index'))


# ===== MÓDULOS SEPARADOS + EDIÇÃO/EXCLUSÃO COM APROVAÇÃO =====
def _registro(conn, tabela, rid):
    return conn.execute(f"SELECT * FROM {tabela} WHERE id=?", (rid,)).fetchone()

def _solicitacao_pendente(conn, tipo, alvo_id):
    return conn.execute("SELECT id FROM solicitacoes_exclusao WHERE tipo=? AND alvo_id=? AND status='pendente' LIMIT 1", (tipo, alvo_id)).fetchone()

def _pedir_exclusao(tipo, alvo_id, descricao):
    conn=get_db()
    if not _solicitacao_pendente(conn,tipo,alvo_id):
        conn.execute('INSERT INTO solicitacoes_exclusao (usuario_id,tipo,alvo_id,descricao) VALUES (?,?,?,?)',(session['usuario_id'],tipo,alvo_id,descricao))
        conn.commit(); flash('Solicitação de exclusão enviada ao administrador.','ok')
    else:
        flash('Já existe uma solicitação de exclusão pendente para este item.','erro')
    conn.close()

def _excluir_alvo(conn, tipo, alvo_id):
    if tipo=='evento':
        for r in conn.execute('SELECT id FROM inscricoes WHERE evento_id=?',(alvo_id,)).fetchall(): conn.execute('DELETE FROM respostas WHERE inscricao_id=?',(r['id'],))
        for r in conn.execute('SELECT id FROM voluntarios WHERE evento_id=?',(alvo_id,)).fetchall(): conn.execute('DELETE FROM respostas_voluntarios WHERE voluntario_id=?',(r['id'],))
        conn.execute('DELETE FROM inscricoes WHERE evento_id=?',(alvo_id,)); conn.execute('DELETE FROM voluntarios WHERE evento_id=?',(alvo_id,))
        conn.execute('DELETE FROM campos WHERE evento_id=?',(alvo_id,)); conn.execute('DELETE FROM atendimentos WHERE evento_id=?',(alvo_id,))
        conn.execute('DELETE FROM campos_voluntarios WHERE evento_id=?',(alvo_id,)); conn.execute('DELETE FROM funcoes_voluntarios WHERE evento_id=?',(alvo_id,))
        conn.execute('DELETE FROM eventos WHERE id=?',(alvo_id,))
    elif tipo=='atendimento': conn.execute('UPDATE inscricoes SET atendimento_id=NULL WHERE atendimento_id=?',(alvo_id,)); conn.execute('DELETE FROM atendimentos WHERE id=?',(alvo_id,))
    elif tipo=='campo': conn.execute('DELETE FROM respostas WHERE campo_id=?',(alvo_id,)); conn.execute('DELETE FROM campos WHERE id=?',(alvo_id,))
    elif tipo=='funcao_voluntario': conn.execute('UPDATE voluntarios SET funcao_id=NULL WHERE funcao_id=?',(alvo_id,)); conn.execute('DELETE FROM funcoes_voluntarios WHERE id=?',(alvo_id,))
    elif tipo=='campo_voluntario': conn.execute('DELETE FROM respostas_voluntarios WHERE campo_id=?',(alvo_id,)); conn.execute('DELETE FROM campos_voluntarios WHERE id=?',(alvo_id,))
    elif tipo=='inscricao': conn.execute('DELETE FROM respostas WHERE inscricao_id=?',(alvo_id,)); conn.execute('DELETE FROM inscricoes WHERE id=?',(alvo_id,))
    elif tipo=='voluntario': conn.execute('DELETE FROM respostas_voluntarios WHERE voluntario_id=?',(alvo_id,)); conn.execute('DELETE FROM voluntarios WHERE id=?',(alvo_id,))
    else: raise ValueError('Tipo de exclusão inválido')

@app.route('/evento/<int:evento_id>/atendimentos', methods=['GET','POST'])
@perfis('admin','operador')
def configurar_atendimentos(evento_id):
    conn=get_db(); evento=_registro(conn,'eventos',evento_id)
    if not evento: conn.close(); abort(404)
    if request.method=='POST':
        acao=request.form.get('acao')
        if acao=='atendimento':
            conn.execute('INSERT INTO atendimentos (evento_id,nome,vagas) VALUES (?,?,?)',(evento_id,request.form['nome_atendimento'].strip(),int(request.form.get('vagas',0) or 0)))
        elif acao=='campo':
            marcador=1 if request.form.get('marcador_cpf') else 0
            if marcador: conn.execute('UPDATE campos SET marcador_cpf=0 WHERE evento_id=?',(evento_id,))
            ordem=conn.execute('SELECT COALESCE(MAX(ordem),0)+1 FROM campos WHERE evento_id=?',(evento_id,)).fetchone()[0]
            conn.execute('INSERT INTO campos (evento_id,titulo,tipo,obrigatorio,opcoes,ordem,marcador_cpf) VALUES (?,?,?,?,?,?,?)',(evento_id,request.form['titulo'].strip(),request.form['tipo'],1 if request.form.get('obrigatorio') else 0,request.form.get('opcoes',''),ordem,marcador))
        conn.commit(); conn.close(); flash('Salvo com sucesso.','ok'); return redirect(url_for('configurar_atendimentos',evento_id=evento_id))
    campos=conn.execute('SELECT * FROM campos WHERE evento_id=? ORDER BY ordem,id',(evento_id,)).fetchall()
    atendimentos=conn.execute('SELECT a.*, (SELECT COUNT(*) FROM inscricoes i WHERE i.atendimento_id=a.id) ocupadas FROM atendimentos a WHERE evento_id=? ORDER BY id',(evento_id,)).fetchall()
    conn.close(); return render_template('config_atendimentos.html',evento=evento,campos=campos,atendimentos=atendimentos)

@app.route('/evento/<int:evento_id>/configurar-voluntarios', methods=['GET','POST'])
@perfis('admin','operador')
def configurar_voluntarios(evento_id):
    conn=get_db(); evento=_registro(conn,'eventos',evento_id)
    if not evento: conn.close(); abort(404)
    if request.method=='POST':
        acao=request.form.get('acao')
        if acao=='funcao_voluntario': conn.execute('INSERT INTO funcoes_voluntarios (evento_id,nome,vagas) VALUES (?,?,?)',(evento_id,request.form['nome_funcao'].strip(),int(request.form.get('vagas_funcao',0) or 0)))
        elif acao=='campo_voluntario':
            ordem=conn.execute('SELECT COALESCE(MAX(ordem),0)+1 FROM campos_voluntarios WHERE evento_id=?',(evento_id,)).fetchone()[0]
            conn.execute('INSERT INTO campos_voluntarios (evento_id,titulo,tipo,obrigatorio,opcoes,ordem) VALUES (?,?,?,?,?,?)',(evento_id,request.form['titulo'].strip(),request.form['tipo'],1 if request.form.get('obrigatorio') else 0,request.form.get('opcoes',''),ordem))
        conn.commit(); conn.close(); flash('Salvo com sucesso.','ok'); return redirect(url_for('configurar_voluntarios',evento_id=evento_id))
    campos=conn.execute('SELECT * FROM campos_voluntarios WHERE evento_id=? ORDER BY ordem,id',(evento_id,)).fetchall()
    funcoes=conn.execute('SELECT f.*, (SELECT COUNT(*) FROM voluntarios v WHERE v.funcao_id=f.id) ocupadas FROM funcoes_voluntarios f WHERE evento_id=? ORDER BY id',(evento_id,)).fetchall()
    conn.close(); return render_template('config_voluntarios.html',evento=evento,campos=campos,funcoes=funcoes)

@app.route('/evento/<int:evento_id>/editar', methods=['GET','POST'])
@perfis('admin','operador')
def editar_evento(evento_id):
    conn=get_db(); item=_registro(conn,'eventos',evento_id)
    if not item: conn.close(); abort(404)
    if request.method=='POST':
        conn.execute('UPDATE eventos SET nome=?, data=?, local=?, permitir_cpf_repetido=? WHERE id=?',(request.form['nome'].strip(),request.form.get('data',''),request.form.get('local',''),1 if request.form.get('permitir_cpf_repetido') else 0,evento_id))
        conn.commit(); conn.close(); flash('Evento atualizado.','ok'); return redirect(url_for('index'))
    conn.close(); return render_template('editar_evento.html',item=item)

@app.route('/atendimento/<int:rid>/editar', methods=['GET','POST'])
@perfis('admin','operador')
def editar_atendimento(rid):
    conn=get_db(); item=_registro(conn,'atendimentos',rid)
    if not item: conn.close(); abort(404)
    if request.method=='POST':
        conn.execute('UPDATE atendimentos SET nome=?, vagas=? WHERE id=?',(request.form['nome'].strip(),int(request.form.get('vagas',0) or 0),rid)); conn.commit(); eid=item['evento_id']; conn.close(); flash('Atendimento atualizado.','ok'); return redirect(url_for('configurar_atendimentos',evento_id=eid))
    conn.close(); return render_template('editar_item.html',titulo='Editar atendimento',item=item)

@app.route('/funcao-voluntario/<int:rid>/editar', methods=['GET','POST'])
@perfis('admin','operador')
def editar_funcao_voluntario(rid):
    conn=get_db(); item=_registro(conn,'funcoes_voluntarios',rid)
    if not item: conn.close(); abort(404)
    if request.method=='POST':
        conn.execute('UPDATE funcoes_voluntarios SET nome=?, vagas=? WHERE id=?',(request.form['nome'].strip(),int(request.form.get('vagas',0) or 0),rid)); conn.commit(); eid=item['evento_id']; conn.close(); flash('Função atualizada.','ok'); return redirect(url_for('configurar_voluntarios',evento_id=eid))
    conn.close(); return render_template('editar_item.html',titulo='Editar função de voluntário',item=item)

@app.route('/campo/<int:rid>/editar', methods=['GET','POST'])
@perfis('admin','operador')
def editar_campo(rid):
    conn=get_db(); item=_registro(conn,'campos',rid)
    if not item: conn.close(); abort(404)
    if request.method=='POST':
        marcador=1 if request.form.get('marcador_cpf') else 0
        if marcador: conn.execute('UPDATE campos SET marcador_cpf=0 WHERE evento_id=?',(item['evento_id'],))
        conn.execute('UPDATE campos SET titulo=?,tipo=?,obrigatorio=?,opcoes=?,marcador_cpf=? WHERE id=?',(request.form['titulo'].strip(),request.form['tipo'],1 if request.form.get('obrigatorio') else 0,request.form.get('opcoes',''),marcador,rid)); conn.commit(); eid=item['evento_id']; conn.close(); flash('Campo atualizado.','ok'); return redirect(url_for('configurar_atendimentos',evento_id=eid))
    conn.close(); return render_template('editar_campo.html',titulo='Editar campo do inscrito',item=item,voluntario=False)

@app.route('/campo-voluntario/<int:rid>/editar', methods=['GET','POST'])
@perfis('admin','operador')
def editar_campo_voluntario(rid):
    conn=get_db(); item=_registro(conn,'campos_voluntarios',rid)
    if not item: conn.close(); abort(404)
    if request.method=='POST':
        conn.execute('UPDATE campos_voluntarios SET titulo=?,tipo=?,obrigatorio=?,opcoes=? WHERE id=?',(request.form['titulo'].strip(),request.form['tipo'],1 if request.form.get('obrigatorio') else 0,request.form.get('opcoes',''),rid)); conn.commit(); eid=item['evento_id']; conn.close(); flash('Campo atualizado.','ok'); return redirect(url_for('configurar_voluntarios',evento_id=eid))
    conn.close(); return render_template('editar_campo.html',titulo='Editar campo do voluntário',item=item,voluntario=True)

def _rota_excluir(tipo, rid, descricao, destino, evento_id=None):
    if usuario_atual()['perfil']=='admin':
        conn=get_db(); _excluir_alvo(conn,tipo,rid); conn.commit(); conn.close(); flash('Exclusão realizada.','ok')
    else: _pedir_exclusao(tipo,rid,descricao)
    kwargs={'evento_id':evento_id} if evento_id is not None else {}
    return redirect(url_for(destino,**kwargs))

@app.route('/evento/<int:rid>/excluir', methods=['POST'])
@perfis('admin','operador')
def excluir_evento(rid):
    conn=get_db(); x=_registro(conn,'eventos',rid); conn.close()
    if not x: abort(404)
    return _rota_excluir('evento',rid,f"Evento: {x['nome']}",'index')

@app.route('/atendimento/<int:rid>/excluir', methods=['POST'])
@perfis('admin','operador')
def excluir_atendimento(rid):
    conn=get_db(); x=_registro(conn,'atendimentos',rid); conn.close()
    if not x: abort(404)
    return _rota_excluir('atendimento',rid,f"Atendimento: {x['nome']}",'configurar_atendimentos',x['evento_id'])

@app.route('/campo/<int:rid>/excluir', methods=['POST'])
@perfis('admin','operador')
def excluir_campo(rid):
    conn=get_db(); x=_registro(conn,'campos',rid); conn.close()
    if not x: abort(404)
    return _rota_excluir('campo',rid,f"Campo de atendimento: {x['titulo']}",'configurar_atendimentos',x['evento_id'])

@app.route('/funcao-voluntario/<int:rid>/excluir', methods=['POST'])
@perfis('admin','operador')
def excluir_funcao_voluntario(rid):
    conn=get_db(); x=_registro(conn,'funcoes_voluntarios',rid); conn.close()
    if not x: abort(404)
    return _rota_excluir('funcao_voluntario',rid,f"Função de voluntário: {x['nome']}",'configurar_voluntarios',x['evento_id'])

@app.route('/campo-voluntario/<int:rid>/excluir', methods=['POST'])
@perfis('admin','operador')
def excluir_campo_voluntario(rid):
    conn=get_db(); x=_registro(conn,'campos_voluntarios',rid); conn.close()
    if not x: abort(404)
    return _rota_excluir('campo_voluntario',rid,f"Campo de voluntário: {x['titulo']}",'configurar_voluntarios',x['evento_id'])

@app.route('/inscricao/<int:rid>/excluir', methods=['POST'])
@perfis('admin','operador')
def excluir_inscricao(rid):
    conn=get_db(); x=_registro(conn,'inscricoes',rid); conn.close()
    if not x: abort(404)
    return _rota_excluir('inscricao',rid,f"Inscrição #{rid}",'inscritos',x['evento_id'])

@app.route('/voluntario/<int:rid>/excluir', methods=['POST'])
@perfis('admin','operador')
def excluir_voluntario(rid):
    conn=get_db(); x=_registro(conn,'voluntarios',rid); conn.close()
    if not x: abort(404)
    return _rota_excluir('voluntario',rid,f"Voluntário #{rid}",'voluntarios',x['evento_id'])

@app.route('/solicitacoes-exclusao')
@perfis('admin')
def solicitacoes_exclusao():
    conn=get_db(); itens=conn.execute("SELECT s.*,u.nome usuario_nome FROM solicitacoes_exclusao s LEFT JOIN usuarios u ON u.id=s.usuario_id WHERE s.status='pendente' ORDER BY s.id DESC").fetchall(); conn.close()
    return render_template('solicitacoes_exclusao.html',itens=itens)

@app.route('/solicitacoes-exclusao/<int:sid>/aprovar', methods=['POST'])
@perfis('admin')
def aprovar_exclusao(sid):
    conn=get_db(); s=conn.execute("SELECT * FROM solicitacoes_exclusao WHERE id=? AND status='pendente'",(sid,)).fetchone()
    if s:
        _excluir_alvo(conn,s['tipo'],s['alvo_id']); conn.execute("UPDATE solicitacoes_exclusao SET status='aprovada' WHERE id=?",(sid,)); conn.commit(); flash('Exclusão aprovada e realizada.','ok')
    conn.close(); return redirect(url_for('solicitacoes_exclusao'))

@app.route('/solicitacoes-exclusao/<int:sid>/recusar', methods=['POST'])
@perfis('admin')
def recusar_exclusao(sid):
    conn=get_db(); conn.execute("UPDATE solicitacoes_exclusao SET status='recusada' WHERE id=? AND status='pendente'",(sid,)); conn.commit(); conn.close(); flash('Solicitação recusada.','ok'); return redirect(url_for('solicitacoes_exclusao'))


# Inicializa tanto no computador quanto no servidor online.
init_db()

if __name__=='__main__':
    app.run(host='127.0.0.1',port=int(os.environ.get('PORT','5000')),debug=False)
