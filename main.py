from fastapi import FastAPI, Depends, HTTPException, Query, UploadFile, File
from fastapi.responses import HTMLResponse, StreamingResponse
from pydantic import BaseModel, Field
from sqlalchemy import create_engine, Column, Integer, String, Float, desc, func
from sqlalchemy.orm import sessionmaker, declarative_base, Session
from datetime import datetime
import io
import csv
import re
import os

# OpenPyXL para exportação de planilhas Excel formatadas
try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False

# ==============================================================================
# 1. BANCO DE DADOS (Neon PostgreSQL em produção ou SQLite local)
# ==============================================================================
if os.path.exists(".env"):
    try:
        with open(".env", "r", encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if line and not line.startswith("#") and "=" in line:
                    k, v = line.split("=", 1)
                    os.environ.setdefault(k.strip(), v.strip())
    except Exception:
        pass

SQLALCHEMY_DATABASE_URL = os.getenv("DATABASE_URL", "sqlite:///./hospital.db")
if SQLALCHEMY_DATABASE_URL.startswith("postgres://"):
    SQLALCHEMY_DATABASE_URL = SQLALCHEMY_DATABASE_URL.replace("postgres://", "postgresql://", 1)

# Se for pooler do Neon, remover channel_binding pois PgBouncer do Neon não suporta SCRAM channel binding
if "-pooler" in SQLALCHEMY_DATABASE_URL and "channel_binding=" in SQLALCHEMY_DATABASE_URL:
    SQLALCHEMY_DATABASE_URL = re.sub(r'[?&]channel_binding=[^&]*', '', SQLALCHEMY_DATABASE_URL)
    if "?" not in SQLALCHEMY_DATABASE_URL and "&" in SQLALCHEMY_DATABASE_URL:
        SQLALCHEMY_DATABASE_URL = SQLALCHEMY_DATABASE_URL.replace("&", "?", 1)

try:
    if "sqlite" in SQLALCHEMY_DATABASE_URL:
        engine = create_engine(SQLALCHEMY_DATABASE_URL, connect_args={"check_same_thread": False})
    else:
        engine = create_engine(
            SQLALCHEMY_DATABASE_URL,
            pool_pre_ping=True,
            pool_recycle=300,
            connect_args={"connect_timeout": 5}
        )
    with engine.connect() as conn:
        pass
    print("✅ Conectado ao banco de dados com sucesso!")
except Exception as e:
    print(f"⚠️ Atenção: Falha ao conectar ao banco remoto. Usando SQLite local (hospital.db). Detalhe: {e}")
    SQLALCHEMY_DATABASE_URL = "sqlite:///./hospital.db"
    engine = create_engine(SQLALCHEMY_DATABASE_URL, connect_args={"check_same_thread": False})

SessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=engine)
Base = declarative_base()

# Tabela de Manutenção (BD_Manutencao)
class Manutencao(Base):
    __tablename__ = "manutencoes"
    
    id = Column(Integer, primary_key=True, index=True)
    codigo_id = Column(String, index=True) # PRED-0001, CLIM-0002, etc.
    data = Column(String) # DD/MM/YYYY
    mes_ano = Column(String, index=True) # MM/YYYY
    horario = Column(String, default="")
    turno = Column(String, default="Matutino") # Matutino, Vespertino, Noturno
    categoria = Column(String, index=True) # Manutenção Predial, Climatização, etc.
    tipo_manutencao = Column(String, default="Corretiva", index=True) # Preventiva, Corretiva
    item = Column(String, index=True) # Sub-categoria / Item
    setor = Column(String, index=True) # Setor / Local
    descricao = Column(String)
    quantidade = Column(Float, default=1.0) # Peças / Volume / Insumos
    numero_chamado = Column(String, default="")
    responsavel_entrega = Column(String, default="") # Rosenilde Costa
    responsavel_recebimento = Column(String, default="")
    status = Column(String, default="Pendente", index=True) # Concluído, Pendente, Solicitado, Ok
    data_conclusao = Column(String, default="")
    observacoes = Column(String, default="")
    retrabalho = Column(Integer, default=0) # 1 = Sim, 0 = Não

# Tabela de Rouparia (BD_Rouparia)
class Rouparia(Base):
    __tablename__ = "rouparias"
    
    id = Column(Integer, primary_key=True, index=True)
    codigo_id = Column(String, index=True) # R-9, R-10, etc.
    data = Column(String) # DD/MM/YYYY
    mes_ano = Column(String, index=True) # MM/YYYY
    horario = Column(String, default="")
    turno = Column(String, default="Matutino")
    tipo_registro = Column(String, index=True) # Recebidos Limpo, Estoque, Recolhimento Sujo, Entrega aos Setores, Devolução
    item_enxoval = Column(String, index=True)
    setor = Column(String, index=True) # Destino / Origem
    qtd_pecas = Column(Float, default=0.0)
    qtd_hamper = Column(Float, default=0.0)
    peso_kg = Column(Float, default=0.0)
    responsavel_unidade = Column(String, default="") # Nome e matrícula
    responsavel_terceirizada = Column(String, default="LAVARE")
    inconformidade = Column(String, default="")
    pontual = Column(Integer, default=1) # 1 = Dentro do horário (07h às 19h)

# Tabela de Configurações e Cadastros Base (Itens, Insumos, Setores, etc.)
class ConfiguracaoItem(Base):
    __tablename__ = "configuracoes_itens"
    
    id = Column(Integer, primary_key=True, index=True)
    tipo = Column(String, index=True) # 'categoria_manutencao', 'item_manutencao', 'item_enxoval', 'setor'
    nome = Column(String, index=True)
    detalhe = Column(String, default="") # ex: 'Insumo', 'Equipamento', 'Crítico'
    ativo = Column(Integer, default=1)

# Tabela de Parâmetros Gerais
class Parametro(Base):
    __tablename__ = "parametros"
    
    chave = Column(String, primary_key=True)
    valor = Column(String)
    descricao = Column(String, default="")

# Tabela de Usuários do Sistema (Login e Controle de Acesso)
import hashlib

ITENS_OFICIAIS_ENXOVAL = [
    "Lençol Paciente",
    "Lençol de Maca",
    "Travessa",
    "Calça Privativa",
    "Camisa Privativa",
    "Roupa Privativa (Kit)",
    "Capote",
    "Campo Simples P (40x40)",
    "Campo Simples P (50x50)",
    "Campo Simples M (80x80)",
    "Campo Simples G (1,20x1,20)",
    "Campo Duplo P (50x50)",
    "Campo Duplo M (80x80)",
    "Campo Duplo G (1,20x1,20)",
    "Campo Fenestrado P",
    "Campo Fenestrado M (80x80)",
    "Campo Fenestrado G (1,20x1,20)",
    "Campo Fenestrado G (1,60x1,60)",
    "Saco Hamper",
    "Pano de Biombo"
]

def gerar_hash_senha(senha: str) -> str:
    return hashlib.sha256(senha.encode('utf-8')).hexdigest()

class Usuario(Base):
    __tablename__ = "usuarios"
    
    id = Column(Integer, primary_key=True, index=True)
    username = Column(String, unique=True, index=True)
    nome = Column(String)
    senha_hash = Column(String)
    perfil = Column(String, default="admin") # admin, operador
    ativo = Column(Integer, default=1)

Base.metadata.create_all(bind=engine)

def get_db():
    db = SessionLocal()
    try:
        yield db
    finally:
        db.close()

# ==============================================================================
# 2. ESQUEMAS PYDANTIC (Validação)
# ==============================================================================
class ManutencaoCreate(BaseModel):
    data: str = ""
    horario: str = ""
    turno: str = "Matutino"
    categoria: str
    tipo_manutencao: str = "Corretiva"
    item: str
    setor: str
    descricao: str
    quantidade: float = 1.0
    numero_chamado: str = ""
    responsavel_entrega: str = ""
    responsavel_recebimento: str = ""
    status: str = "Pendente"
    data_conclusao: str = ""
    observacoes: str = ""
    retrabalho: int = 0

class RoupariaCreate(BaseModel):
    data: str = ""
    horario: str = ""
    turno: str = "Matutino"
    tipo_registro: str = "Estoque"
    item_enxoval: str
    setor: str
    qtd_pecas: float = 0.0
    qtd_hamper: float = 0.0
    peso_kg: float = 0.0
    responsavel_unidade: str = ""
    responsavel_terceirizada: str = "LAVARE"
    inconformidade: str = ""
    pontual: int = 1

class ConfigItemCreate(BaseModel):
    tipo: str # categoria_manutencao, item_manutencao, item_enxoval, setor
    nome: str
    detalhe: str = ""

class ParametroUpdate(BaseModel):
    chave: str
    valor: str

class StatusUpdate(BaseModel):
    status: str

class LoginRequest(BaseModel):
    username: str
    password: str

class ManutencaoUpdate(BaseModel):
    data: str
    horario: str = ""
    turno: str = "Matutino"
    categoria: str
    tipo_manutencao: str = "Corretiva"
    item: str
    setor: str
    descricao: str
    quantidade: float = 1.0
    numero_chamado: str = ""
    responsavel_entrega: str = ""
    responsavel_recebimento: str = ""
    status: str = "Pendente"
    data_conclusao: str = ""
    observacoes: str = ""
    retrabalho: int = 0

class RoupariaUpdate(BaseModel):
    data: str
    horario: str = ""
    turno: str = "Matutino"
    tipo_registro: str = "Estoque"
    item_enxoval: str
    setor: str
    qtd_pecas: float = 0.0
    qtd_hamper: float = 0.0
    peso_kg: float = 0.0
    responsavel_unidade: str = ""
    responsavel_terceirizada: str = "LAVARE"
    inconformidade: str = ""

# ==============================================================================
# 3. ROTAS DA API
# ==============================================================================
app = FastAPI(title="SaaS Gestão de Manutenção & Rouparia - PS Anil", version="3.0")

@app.on_event("startup")
def inicializar_usuarios_padrao():
    db = SessionLocal()
    try:
        # Garantir admin com senha anil2026
        admin = db.query(Usuario).filter(func.lower(Usuario.username) == "admin").first()
        if not admin:
            admin = Usuario(
                username="admin",
                nome="Administrador PS Anil",
                senha_hash=gerar_hash_senha("anil2026"),
                perfil="admin",
                ativo=1
            )
            db.add(admin)
        else:
            admin.senha_hash = gerar_hash_senha("anil2026")
            admin.ativo = 1

        # Garantir plantao com senha psanil
        plantao = db.query(Usuario).filter(func.lower(Usuario.username) == "plantao").first()
        if not plantao:
            plantao = Usuario(
                username="plantao",
                nome="Plantão PS Anil",
                senha_hash=gerar_hash_senha("psanil"),
                perfil="operador",
                ativo=1
            )
            db.add(plantao)
        else:
            plantao.senha_hash = gerar_hash_senha("psanil")
            plantao.ativo = 1

        # Sincronizar itens oficiais de enxoval (manter estritamente os 20 itens oficiais)
        db.query(ConfiguracaoItem).filter(
            ConfiguracaoItem.tipo == "item_enxoval",
            ConfiguracaoItem.nome.notin_(ITENS_OFICIAIS_ENXOVAL)
        ).update({"ativo": 0}, synchronize_session=False)

        for it_nome in ITENS_OFICIAIS_ENXOVAL:
            it_exist = db.query(ConfiguracaoItem).filter(
                ConfiguracaoItem.tipo == "item_enxoval",
                ConfiguracaoItem.nome == it_nome
            ).first()
            if it_exist:
                it_exist.ativo = 1
            else:
                db.add(ConfiguracaoItem(tipo="item_enxoval", nome=it_nome, detalhe="Oficial", ativo=1))

        db.commit()
    except Exception as e:
        print(f"Erro ao inicializar usuários padrão: {e}")
    finally:
        db.close()

# --- AUTENTICAÇÃO / LOGIN ---
@app.post("/api/login")
def fazer_login(dados: LoginRequest, db: Session = Depends(get_db)):
    user_input = dados.username.strip().lower()
    senha_input = dados.password.strip()
    u = db.query(Usuario).filter(func.lower(Usuario.username) == user_input, Usuario.ativo == 1).first()
    if not u or u.senha_hash != gerar_hash_senha(senha_input):
        raise HTTPException(status_code=401, detail="Usuário ou senha incorretos.")
    return {
        "status": "sucesso",
        "usuario": {
            "id": u.id,
            "username": u.username,
            "nome": u.nome,
            "perfil": u.perfil
        }
    }

def normalizar_data(data_str: str) -> str:
    if not data_str:
        return datetime.now().strftime("%d/%m/%Y")
    data_str = data_str.strip()
    if "-" in data_str:
        partes = data_str.split("-")
        if len(partes) == 3:
            return f"{partes[2]}/{partes[1]}/{partes[0]}"
    return data_str

def extrair_mes_ano(data_str: str) -> str:
    if not data_str:
        return datetime.now().strftime("%m/%Y")
    data_str = data_str.strip()
    if "/" in data_str:
        p = data_str.split("/")
        if len(p) >= 3:
            return f"{p[1]}/{p[2]}"
    elif "-" in data_str:
        p = data_str.split("-")
        if len(p) >= 3:
            return f"{p[1]}/{p[0]}"
    return datetime.now().strftime("%m/%Y")

# --- MANUTENÇÃO ---
@app.post("/api/manutencao")
def criar_manutencao(dados: ManutencaoCreate, db: Session = Depends(get_db)):
    data_str = normalizar_data(dados.data)
    mes_ano_str = extrair_mes_ano(dados.data)
    horario_str = dados.horario if dados.horario else datetime.now().strftime("%H:%M")
    data_conc_str = normalizar_data(dados.data_conclusao) if dados.data_conclusao else ""
    
    # Gerar Código ID automático (ex: PRED-0021, CLIM-0022, etc.)
    prefixo = "MANUT"
    cat_lower = dados.categoria.lower()
    if "predial" in cat_lower:
        prefixo = "PRED"
    elif "climatiza" in cat_lower:
        prefixo = "CLIM"
    elif "praga" in cat_lower:
        prefixo = "PRAG"
    elif "elétr" in cat_lower or "eletr" in cat_lower:
        prefixo = "ELET"
    elif "hidro" in cat_lower:
        prefixo = "HIDR"
        
    proximo_num = db.query(Manutencao).count() + 1
    codigo_gerado = f"{prefixo}-{str(proximo_num).zfill(4)}"
    
    nova = Manutencao(
        codigo_id=codigo_gerado,
        data=data_str,
        mes_ano=mes_ano_str,
        horario=horario_str,
        turno=dados.turno,
        categoria=dados.categoria,
        tipo_manutencao=dados.tipo_manutencao,
        item=dados.item,
        setor=dados.setor,
        descricao=dados.descricao,
        quantidade=dados.quantidade,
        numero_chamado=dados.numero_chamado,
        responsavel_entrega=dados.responsavel_entrega,
        responsavel_recebimento=dados.responsavel_recebimento,
        status=dados.status,
        data_conclusao=dados.data_conclusao,
        observacoes=dados.observacoes,
        retrabalho=dados.retrabalho
    )
    db.add(nova)
    db.commit()
    db.refresh(nova)
    return {"status": "sucesso", "mensagem": f"Lançamento {codigo_gerado} gravado com sucesso!", "id": nova.id, "codigo": codigo_gerado}

@app.get("/api/manutencao")
def listar_manutencoes(
    status: str = Query(None),
    categoria: str = Query(None),
    tipo: str = Query(None),
    setor: str = Query(None),
    mes_ano: str = Query(None),
    turno: str = Query(None),
    retrabalho: str = Query(None),
    busca: str = Query(None),
    db: Session = Depends(get_db)
):
    q = db.query(Manutencao)
    if status and status != "Todos":
        q = q.filter(Manutencao.status == status)
    if categoria and categoria not in ["Todas", "Todos"]:
        q = q.filter(Manutencao.categoria == categoria)
    if tipo and tipo != "Todos":
        q = q.filter(Manutencao.tipo_manutencao == tipo)
    if setor and setor != "Todos":
        q = q.filter(Manutencao.setor == setor)
    if mes_ano and mes_ano != "Todos":
        q = q.filter(Manutencao.mes_ano == mes_ano)
    if turno and turno != "Todos":
        q = q.filter(Manutencao.turno == turno)
    if retrabalho and retrabalho != "Todos":
        if retrabalho in ["Sim", "1", "true", "True"]:
            q = q.filter(Manutencao.retrabalho == 1)
        elif retrabalho in ["Não", "Nao", "0", "false", "False"]:
            q = q.filter(Manutencao.retrabalho == 0)
    if busca:
        termo = f"%{busca}%"
        q = q.filter(
            (Manutencao.codigo_id.like(termo)) |
            (Manutencao.item.like(termo)) |
            (Manutencao.descricao.like(termo)) |
            (Manutencao.setor.like(termo)) |
            (Manutencao.responsavel_entrega.like(termo)) |
            (Manutencao.responsavel_recebimento.like(termo)) |
            (Manutencao.numero_chamado.like(termo))
        )
    registros = q.order_by(desc(Manutencao.id)).all()
    return registros

@app.get("/api/manutencao/{id}")
def obter_manutencao(id: int, db: Session = Depends(get_db)):
    reg = db.query(Manutencao).filter(Manutencao.id == id).first()
    if not reg:
        raise HTTPException(status_code=404, detail="Registro não encontrado")
    return reg

@app.put("/api/manutencao/{id}")
def atualizar_manutencao(id: int, dados: ManutencaoUpdate, db: Session = Depends(get_db)):
    reg = db.query(Manutencao).filter(Manutencao.id == id).first()
    if not reg:
        raise HTTPException(status_code=404, detail="Registro não encontrado")
    
    reg.data = normalizar_data(dados.data)
    reg.mes_ano = extrair_mes_ano(dados.data)
    reg.horario = dados.horario
    reg.turno = dados.turno
    reg.categoria = dados.categoria
    reg.tipo_manutencao = dados.tipo_manutencao
    reg.item = dados.item
    reg.setor = dados.setor
    reg.descricao = dados.descricao
    reg.quantidade = dados.quantidade
    reg.numero_chamado = dados.numero_chamado
    reg.responsavel_entrega = dados.responsavel_entrega
    reg.responsavel_recebimento = dados.responsavel_recebimento
    reg.status = dados.status
    reg.data_conclusao = normalizar_data(dados.data_conclusao) if dados.data_conclusao else ""
    
    # Acrescenta observação/nota de auditoria quando o item é editado
    obs_informada = (dados.observacoes or "").strip()
    agora_txt = datetime.now().strftime("%d/%m/%Y às %H:%M")
    if "[Editado em" not in obs_informada:
        reg.observacoes = f"{obs_informada}\n[Editado em {agora_txt}]".strip()
    else:
        reg.observacoes = obs_informada
        
    reg.retrabalho = dados.retrabalho
    db.commit()
    return {"status": "sucesso", "mensagem": f"Chamado {reg.codigo_id} atualizado com sucesso!", "observacoes": reg.observacoes}

@app.patch("/api/manutencao/{id}/status")
def atualizar_status_manutencao(id: int, dados: StatusUpdate, db: Session = Depends(get_db)):
    reg = db.query(Manutencao).filter(Manutencao.id == id).first()
    if not reg:
        raise HTTPException(status_code=404, detail="Registro não encontrado")
    reg.status = dados.status
    if dados.status in ["Concluído", "Ok"] and not reg.data_conclusao:
        reg.data_conclusao = datetime.now().strftime("%d/%m/%Y")
    elif dados.status in ["Pendente", "Solicitado"]:
        reg.data_conclusao = ""
    db.commit()
    return {"status": "sucesso", "mensagem": f"Status atualizado para {dados.status}"}

@app.delete("/api/manutencao/{id}")
def excluir_manutencao(id: int, db: Session = Depends(get_db)):
    reg = db.query(Manutencao).filter(Manutencao.id == id).first()
    if not reg:
        raise HTTPException(status_code=404, detail="Registro não encontrado")
    db.delete(reg)
    db.commit()
    return {"status": "sucesso", "mensagem": "Registro excluído com sucesso"}

# --- ROUPARIA ---
@app.post("/api/rouparia")
def criar_rouparia(dados: RoupariaCreate, db: Session = Depends(get_db)):
    data_str = normalizar_data(dados.data)
    mes_ano_str = extrair_mes_ano(dados.data)
    horario_str = dados.horario if dados.horario else datetime.now().strftime("%H:%M")
    
    proximo_num = db.query(Rouparia).count() + 1
    codigo_gerado = f"R-{proximo_num}"
    
    nova = Rouparia(
        codigo_id=codigo_gerado,
        data=data_str,
        mes_ano=mes_ano_str,
        horario=horario_str,
        turno=dados.turno,
        tipo_registro=dados.tipo_registro,
        item_enxoval=dados.item_enxoval,
        setor=dados.setor,
        qtd_pecas=dados.qtd_pecas,
        qtd_hamper=dados.qtd_hamper,
        peso_kg=dados.peso_kg,
        responsavel_unidade=dados.responsavel_unidade,
        responsavel_terceirizada=dados.responsavel_terceirizada,
        inconformidade=dados.inconformidade,
        pontual=dados.pontual
    )
    db.add(nova)
    db.commit()
    db.refresh(nova)
    return {"status": "sucesso", "mensagem": f"Movimentação {codigo_gerado} registrada com sucesso!", "id": nova.id, "codigo": codigo_gerado}

@app.get("/api/rouparia")
def listar_rouparia(
    tipo_registro: str = Query(None),
    item: str = Query(None),
    setor: str = Query(None),
    mes_ano: str = Query(None),
    turno: str = Query(None),
    busca: str = Query(None),
    db: Session = Depends(get_db)
):
    q = db.query(Rouparia)
    if tipo_registro and tipo_registro != "Todos":
        q = q.filter(Rouparia.tipo_registro == tipo_registro)
    if item and item != "Todos":
        q = q.filter(Rouparia.item_enxoval == item)
    if setor and setor != "Todos":
        q = q.filter(Rouparia.setor == setor)
    if mes_ano and mes_ano != "Todos":
        q = q.filter(Rouparia.mes_ano == mes_ano)
    if turno and turno != "Todos":
        q = q.filter(Rouparia.turno == turno)
    if busca:
        termo = f"%{busca}%"
        q = q.filter(
            (Rouparia.codigo_id.like(termo)) |
            (Rouparia.item_enxoval.like(termo)) |
            (Rouparia.setor.like(termo)) |
            (Rouparia.responsavel_unidade.like(termo)) |
            (Rouparia.responsavel_terceirizada.like(termo)) |
            (Rouparia.inconformidade.like(termo))
        )
    return q.order_by(desc(Rouparia.id)).all()

@app.get("/api/rouparia/{id}")
def obter_rouparia(id: int, db: Session = Depends(get_db)):
    reg = db.query(Rouparia).filter(Rouparia.id == id).first()
    if not reg:
        raise HTTPException(status_code=404, detail="Registro não encontrado")
    return reg

@app.put("/api/rouparia/{id}")
def atualizar_rouparia(id: int, dados: RoupariaUpdate, db: Session = Depends(get_db)):
    reg = db.query(Rouparia).filter(Rouparia.id == id).first()
    if not reg:
        raise HTTPException(status_code=404, detail="Registro não encontrado")
        
    reg.data = normalizar_data(dados.data)
    reg.mes_ano = extrair_mes_ano(dados.data)
    reg.horario = dados.horario
    reg.turno = dados.turno
    reg.tipo_registro = dados.tipo_registro
    reg.item_enxoval = dados.item_enxoval
    reg.setor = dados.setor
    reg.qtd_pecas = dados.qtd_pecas
    reg.qtd_hamper = dados.qtd_hamper
    reg.peso_kg = dados.peso_kg
    reg.responsavel_unidade = dados.responsavel_unidade
    reg.responsavel_terceirizada = dados.responsavel_terceirizada
    
    # Acrescenta observação/nota de auditoria quando o item é editado
    inconf_informada = (dados.inconformidade or "").strip()
    agora_txt = datetime.now().strftime("%d/%m/%Y às %H:%M")
    if "[Editado em" not in inconf_informada:
        reg.inconformidade = f"{inconf_informada}\n[Editado em {agora_txt}]".strip()
    else:
        reg.inconformidade = inconf_informada
        
    db.commit()
    return {"status": "sucesso", "mensagem": f"Movimentação {reg.codigo_id} atualizada com sucesso!"}

@app.delete("/api/rouparia/{id}")
def excluir_rouparia(id: int, db: Session = Depends(get_db)):
    reg = db.query(Rouparia).filter(Rouparia.id == id).first()
    if not reg:
        raise HTTPException(status_code=404, detail="Registro não encontrado")
    db.delete(reg)
    db.commit()
    return {"status": "sucesso", "mensagem": "Movimentação excluída com sucesso"}

# --- CONFIGURAÇÕES & CADASTROS ---
@app.get("/api/configuracoes")
def listar_configuracoes(tipo: str = Query(None), db: Session = Depends(get_db)):
    q = db.query(ConfiguracaoItem).filter(ConfiguracaoItem.ativo == 1)
    if tipo:
        q = q.filter(ConfiguracaoItem.tipo == tipo)
    if tipo == "item_enxoval":
        itens = q.all()
        ordem_map = {nome: i for i, nome in enumerate(ITENS_OFICIAIS_ENXOVAL)}
        itens = sorted([i for i in itens if i.nome in ordem_map], key=lambda x: ordem_map.get(x.nome, 999))
        return itens
    itens = q.order_by(ConfiguracaoItem.nome.asc()).all()
    return itens

@app.post("/api/configuracoes")
def adicionar_configuracao(dados: ConfigItemCreate, db: Session = Depends(get_db)):
    nome_limpo = dados.nome.strip()
    if not nome_limpo:
        raise HTTPException(status_code=400, detail="O nome não pode ser vazio")
    
    existe = db.query(ConfiguracaoItem).filter(
        ConfiguracaoItem.tipo == dados.tipo,
        ConfiguracaoItem.nome == nome_limpo
    ).first()
    
    if existe:
        if existe.ativo == 0:
            existe.ativo = 1
            existe.detalhe = dados.detalhe
            db.commit()
            return {"status": "sucesso", "mensagem": "Item reativado com sucesso!"}
        raise HTTPException(status_code=400, detail="Este item já está cadastrado")
        
    novo = ConfiguracaoItem(
        tipo=dados.tipo,
        nome=nome_limpo,
        detalhe=dados.detalhe,
        ativo=1
    )
    db.add(novo)
    db.commit()
    return {"status": "sucesso", "mensagem": f"{nome_limpo} adicionado com sucesso!"}

@app.delete("/api/configuracoes/{id}")
def remover_configuracao(id: int, db: Session = Depends(get_db)):
    reg = db.query(ConfiguracaoItem).filter(ConfiguracaoItem.id == id).first()
    if not reg:
        raise HTTPException(status_code=404, detail="Item não encontrado")
    # Soft delete
    reg.ativo = 0
    db.commit()
    return {"status": "sucesso", "mensagem": "Item desativado com sucesso"}

# --- PARÂMETROS ---
@app.get("/api/parametros")
def obter_parametros(db: Session = Depends(get_db)):
    params = db.query(Parametro).all()
    return {p.chave: p.valor for p in params}

@app.post("/api/parametros")
def salvar_parametro(dados: ParametroUpdate, db: Session = Depends(get_db)):
    p = db.query(Parametro).filter(Parametro.chave == dados.chave).first()
    if p:
        p.valor = dados.valor
    else:
        p = Parametro(chave=dados.chave, valor=dados.valor)
        db.add(p)
    db.commit()
    return {"status": "sucesso", "mensagem": "Parâmetro atualizado com sucesso"}

# ==============================================================================
# 4. PAINEL CONSOLIDADO DE INDICADORES (FÓRMULAS EXATAS - IMAGEM 6)
# ==============================================================================
@app.get("/api/consolidado")
def obter_consolidado(mes_ano: str = Query("Todos"), db: Session = Depends(get_db)):
    # 1. Filtros por Mês/Ano
    q_manut = db.query(Manutencao)
    q_roupa = db.query(Rouparia)
    
    if mes_ano != "Todos":
        q_manut = q_manut.filter(Manutencao.mes_ano == mes_ano)
        q_roupa = q_roupa.filter(Rouparia.mes_ano == mes_ano)
        
    manutencoes = q_manut.all()
    rouparias = q_roupa.all()
    
    # --------------------------------------------------------------------------
    # FÓRMULAS: 1. INDICADORES DE MANUTENÇÃO
    # --------------------------------------------------------------------------
    # a) Preventivas Agendadas / Total
    preventivas_total = sum(1 for m in manutencoes if m.tipo_manutencao == "Preventiva")
    
    # b) Preventivas Realizadas (Status Concluído ou Ok)
    preventivas_realizadas = sum(1 for m in manutencoes if m.tipo_manutencao == "Preventiva" and m.status in ["Concluído", "Ok"])
    
    # c) % Aderência Preventiva = (Preventivas Realizadas / Preventivas Agendadas) * 100
    aderencia_preventiva = round((preventivas_realizadas / preventivas_total * 100), 2) if preventivas_total > 0 else 0.0
    
    # d) Corretivas Solicitadas
    corretivas_solicitadas = sum(1 for m in manutencoes if m.tipo_manutencao == "Corretiva")
    
    # e) Corretivas Concluídas (Status Concluído ou Ok)
    corretivas_concluidas = sum(1 for m in manutencoes if m.tipo_manutencao == "Corretiva" and m.status in ["Concluído", "Ok"])
    
    # f) % Resolução de Corretivas = (Corretivas Concluídas / Corretivas Solicitadas) * 100
    resolucao_corretivas = round((corretivas_concluidas / corretivas_solicitadas * 100), 2) if corretivas_solicitadas > 0 else 0.0
    
    # g) Total Peças / Insumos Aplicados = SOMA(Quantidade)
    total_pecas_insumos = sum(m.quantidade or 0.0 for m in manutencoes)
    
    # h) Taxa de Retrabalho = (Retrabalhos / Total de Concluídas ou Realizadas) * 100
    total_concluidas_geral = sum(1 for m in manutencoes if m.status in ["Concluído", "Ok"])
    total_retrabalho = sum(1 for m in manutencoes if m.retrabalho == 1)
    
    # Se houver manutenções, calcular percentual; se houver concluídas, usa concluídas
    divisor_retrabalho = total_concluidas_geral if total_concluidas_geral > 0 else len(manutencoes)
    taxa_retrabalho = round((total_retrabalho / divisor_retrabalho * 100), 2) if divisor_retrabalho > 0 else 0.0
    
    # --------------------------------------------------------------------------
    # FÓRMULAS: 2. INDICADORES DE ROUPARIA
    # --------------------------------------------------------------------------
    # a) Total de Peças Entregues aos Setores
    pecas_entregues = sum(r.qtd_pecas or 0.0 for r in rouparias if r.tipo_registro == "Entrega aos Setores")
    
    # b) Total de Enxoval Sujo Recolhido (KG)
    enxoval_sujo_kg = sum(r.peso_kg or 0.0 for r in rouparias if r.tipo_registro == "Recolhimento Sujo")
    
    # c) Total de Enxoval Limpo Recebido (KG)
    enxoval_limpo_kg = sum(r.peso_kg or 0.0 for r in rouparias if (r.tipo_registro or "").strip().lower() in ["estoque", "recebimento limpo", "recebidos limpo", "recebido limpo", "recebido lavanderia", "recebimento"])
    
    # d) Devoluções por Não Conformidade (KG / Peças)
    devolucoes = sum(r.qtd_pecas or 0.0 for r in rouparias if r.tipo_registro == "Devolução" or (r.inconformidade and r.inconformidade.strip()))
    
    # e) Nível de Atendimento de Rouparia (%) = (Entregues / (Entregues + Devoluções)) * 100
    total_movimentado = pecas_entregues + devolucoes
    nivel_atendimento = round((pecas_entregues / total_movimentado * 100), 2) if total_movimentado > 0 else 100.0
    
    # f) Pontualidade na Entrega (07 as 19h)
    entregas_total = sum(1 for r in rouparias if r.tipo_registro == "Entrega aos Setores")
    entregas_pontuais = sum(1 for r in rouparias if r.tipo_registro == "Entrega aos Setores" and r.pontual == 1)
    pontualidade = round((entregas_pontuais / entregas_total * 100), 2) if entregas_total > 0 else 100.0
    
    # g) Custo por quilo (R$/KG) e Custo Total Lavanderia (R$)
    param_custo = db.query(Parametro).filter(Parametro.chave == "custo_kg_lavanderia").first()
    custo_kg = float(param_custo.valor) if param_custo and param_custo.valor else 0.0
    custo_total_lavanderia = round(enxoval_sujo_kg * custo_kg, 2)
    
    # Obter lista de Meses/Anos únicos disponíveis para o filtro
    meses_manut = db.query(Manutencao.mes_ano).distinct().all()
    meses_roupa = db.query(Rouparia.mes_ano).distinct().all()
    meses_set = set([m[0] for m in meses_manut if m[0]] + [r[0] for r in meses_roupa if r[0]])
    meses_disponiveis = sorted(list(meses_set), reverse=True)
    
    return {
        "mes_ano_selecionado": mes_ano,
        "meses_disponiveis": meses_disponiveis,
        "manutencao": {
            "preventivas_realizadas": preventivas_realizadas,
            "preventivas_agendadas": preventivas_total,
            "aderencia_preventiva": f"{aderencia_preventiva:.2f}%",
            "aderencia_preventiva_num": aderencia_preventiva,
            "corretivas_solicitadas": corretivas_solicitadas,
            "corretivas_concluidas": corretivas_concluidas,
            "resolucao_corretivas": f"{resolucao_corretivas:.2f}%",
            "resolucao_corretivas_num": resolucao_corretivas,
            "total_pecas_insumos": int(total_pecas_insumos),
            "taxa_retrabalho": f"{taxa_retrabalho:.2f}%",
            "taxa_retrabalho_num": taxa_retrabalho,
            "total_registros": len(manutencoes)
        },
        "rouparia": {
            "pecas_entregues": int(pecas_entregues),
            "enxoval_sujo_kg": round(enxoval_sujo_kg, 2),
            "enxoval_limpo_kg": round(enxoval_limpo_kg, 2),
            "devolucoes_inconformidade": int(devolucoes),
            "nivel_atendimento": f"{nivel_atendimento:.2f}%",
            "pontualidade": f"{pontualidade:.2f}%",
            "custo_kg": f"R$ {custo_kg:.2f}".replace(".", ","),
            "custo_total_lavanderia": f"R$ {custo_total_lavanderia:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."),
            "total_registros": len(rouparias)
        }
    }

# ==============================================================================
# 5. RELATÓRIOS DETALHADOS (IMAGEM 4 - STATUS, CATEGORIAS, SETORES E MATRIZ)
# ==============================================================================
@app.get("/api/relatorios/manutencao")
def obter_relatorio_manutencao(mes_ano: str = Query("Todos"), db: Session = Depends(get_db)):
    q = db.query(Manutencao)
    if mes_ano != "Todos":
        q = q.filter(Manutencao.mes_ano == mes_ano)
    registros = q.all()
    
    # 1. Status / Situação
    status_contagem = {}
    for r in registros:
        st = r.status if r.status else "Pendente"
        status_contagem[st] = status_contagem.get(st, 0) + 1
        
    # 2. Categorias
    categoria_contagem = {}
    for r in registros:
        cat = r.categoria if r.categoria else "Outros"
        categoria_contagem[cat] = categoria_contagem.get(cat, 0) + 1
        
    # 3. Setor / Local (Ranking ordenado)
    setor_contagem = {}
    for r in registros:
        st = r.setor if r.setor else "Não informado"
        setor_contagem[st] = setor_contagem.get(st, 0) + 1
    ranking_setores = sorted(setor_contagem.items(), key=lambda x: x[1], reverse=True)
    
    # 4. Matriz Cruzada (Setor x Item) - Como na Imagem 4!
    matriz_setor_item = {} # {setor: {item: qtd}}
    itens_unicos = set()
    for r in registros:
        st = r.setor if r.setor else "Não informado"
        it = r.item if r.item else "Diversos"
        itens_unicos.add(it)
        if st not in matriz_setor_item:
            matriz_setor_item[st] = {}
        matriz_setor_item[st][it] = matriz_setor_item[st].get(it, 0) + 1
        
    return {
        "total_geral": len(registros),
        "status": status_contagem,
        "categorias": categoria_contagem,
        "setores_ranking": [{"setor": s, "quantidade": q} for s, q in ranking_setores],
        "itens_unicos": sorted(list(itens_unicos)),
        "matriz": matriz_setor_item
    }

# ==============================================================================
# 5.1 DOCUMENTOS OFICIAIS DE ROUPARIA (RELATÓRIO DIÁRIO, ENXOVAL ENTREGUE, PESO SUJO)
# ==============================================================================
@app.get("/api/relatorios/rouparia/documentos")
def obter_documentos_rouparia(data: str = Query("01/08/2026"), db: Session = Depends(get_db)):
    # 1. Obter itens de enxoval oficiais (estritamente os 20 itens oficiais da planilha física)
    itens_nomes = ITENS_OFICIAIS_ENXOVAL

    # 2. Datas disponíveis na base de Rouparia
    datas_brutas = db.query(Rouparia.data).distinct().all()
    datas_disponiveis = [r[0] for r in datas_brutas if r[0] and "/" in r[0]]
    def parse_data_br(d_str):
        try:
            pts = d_str.split("/")
            return (int(pts[2]), int(pts[1]), int(pts[0]))
        except:
            return (0, 0, 0)
    datas_disponiveis = sorted(list(set(datas_disponiveis)), key=parse_data_br, reverse=True)
    if not datas_disponiveis:
        datas_disponiveis = ["01/08/2026"]

    data_consulta = data if (data and data in datas_disponiveis) else datas_disponiveis[0]

    # 3. Registros do dia selecionado
    registros_dia = db.query(Rouparia).filter(Rouparia.data == data_consulta).all()

    # 4. Cálculo do Estoque Global (Saldo acumulado: recebidos/estoque menos entregas)
    todos_registros = db.query(Rouparia).all()
    estoque_map = {it: 0 for it in itens_nomes}
    for r in todos_registros:
        it = (r.item_enxoval or "").strip()
        if not it:
            continue
        it_key = next((x for x in itens_nomes if x.lower() == it.lower()), it)
        if it_key not in estoque_map:
            estoque_map[it_key] = 0
        tipo = (r.tipo_registro or "").strip().lower()
        if tipo in ["recebido lavanderia", "estoque", "recebimento", "recebidos limpo", "recebimento limpo", "recebido limpo"]:
            estoque_map[it_key] += int(r.qtd_pecas or 0)
        elif tipo in ["entrega aos setores", "entrega setor"]:
            estoque_map[it_key] -= int(r.qtd_pecas or 0)

    # 5. Documento 1: Relatório Diário
    recebidos_dia = {it: 0 for it in itens_nomes}
    for r in registros_dia:
        tipo = (r.tipo_registro or "").strip().lower()
        if tipo in ["recebido lavanderia", "recebimento", "recebidos limpo", "recebimento limpo", "recebido limpo"]:
            it = (r.item_enxoval or "").strip()
            it_key = next((x for x in itens_nomes if x.lower() == it.lower()), it)
            recebidos_dia[it_key] = recebidos_dia.get(it_key, 0) + int(r.qtd_pecas or 0)

    # Intercorrências por Setor no dia
    intercorrencias_setor = {}
    for r in registros_dia:
        if r.inconformidade and r.inconformidade.strip():
            st = r.setor or "Geral"
            if st not in intercorrencias_setor:
                intercorrencias_setor[st] = []
            intercorrencias_setor[st].append({
                "item": r.item_enxoval or "-",
                "inconformidade": r.inconformidade,
                "qtd": r.qtd_pecas or 0
            })

    # Distribuição Total de Roupas por Setor no dia
    distribuicao_dia = []
    for r in registros_dia:
        tipo = (r.tipo_registro or "").strip().lower()
        if tipo in ["entrega aos setores", "entrega setor"]:
            distribuicao_dia.append({
                "setor": r.setor or "-",
                "item": r.item_enxoval or "-",
                "qtd": int(r.qtd_pecas or 0)
            })

    # Recolhimento Interno
    recolhimento_interno_dia = []
    for r in registros_dia:
        tipo = (r.tipo_registro or "").strip().lower()
        if tipo == "recolhimento interno":
            recolhimento_interno_dia.append({
                "setor": r.setor or "-",
                "peso_kg": float(r.peso_kg or 0.0)
            })

    devolucao_inconf_kg = sum(float(r.peso_kg or 0.0) for r in registros_dia if (r.tipo_registro or "").strip().lower() == "devolução inconformidade")
    recolhimento_sujo_kg = sum(float(r.peso_kg or 0.0) for r in registros_dia if (r.tipo_registro or "").strip().lower() == "recolhimento sujo")
    ocorrencias_turno = [r.inconformidade for r in registros_dia if r.inconformidade and r.inconformidade.strip()]

    # 6. Documento 2: Planilha de Enxoval Entregue nas Áreas (Doc 01)
    entregas_areas = {}
    for r in registros_dia:
        tipo = (r.tipo_registro or "").strip().lower()
        if tipo in ["entrega aos setores", "entrega setor"]:
            st = r.setor or "Setor"
            if st not in entregas_areas:
                entregas_areas[st] = {
                    "setor": st,
                    "horario": r.horario or "",
                    "responsavel_recebimento": r.responsavel_unidade or "",
                    "itens": {it: 0 for it in itens_nomes}
                }
            it = (r.item_enxoval or "").strip()
            it_key = next((x for x in itens_nomes if x.lower() == it.lower()), it)
            entregas_areas[st]["itens"][it_key] = entregas_areas[st]["itens"].get(it_key, 0) + int(r.qtd_pecas or 0)

    # 7. Documento 3: Planilha de Peso do Enxoval Sujo (Doc 02)
    q_sujo = db.query(Rouparia).filter(Rouparia.tipo_registro.ilike("%Recolhimento Sujo%"))
    regs_sujo = q_sujo.filter(Rouparia.data == data_consulta).order_by(Rouparia.id.asc()).all()
    if not regs_sujo and "/" in data_consulta:
        p = data_consulta.split("/")
        mes_ano_ref = f"{p[1]}/{p[2]}"
        regs_sujo = q_sujo.filter(Rouparia.mes_ano == mes_ano_ref).order_by(Rouparia.id.asc()).all()
    if not regs_sujo:
        regs_sujo = q_sujo.order_by(Rouparia.id.asc()).limit(30).all()

    lista_sujo = []
    tot_h = 0.0
    tot_p = 0.0
    for idx, r in enumerate(regs_sujo, 1):
        qh = float(r.qtd_hamper or 0.0)
        pk = float(r.peso_kg or 0.0)
        tot_h += qh
        tot_p += pk
        lista_sujo.append({
            "ord": idx,
            "data": r.data,
            "qtd_hamper": qh,
            "peso_kg": round(pk, 2),
            "resp_entrega": r.responsavel_unidade or "SEMUS / PS ANIL",
            "resp_recebimento": r.responsavel_terceirizada or "LAVARE"
        })

    return {
        "data_selecionada": data_consulta,
        "datas_disponiveis": datas_disponiveis,
        "itens_nomes": itens_nomes,
        "relatorio_diario": {
            "recebidos": recebidos_dia,
            "estoque": estoque_map,
            "intercorrencias": intercorrencias_setor,
            "distribuicao": distribuicao_dia,
            "recolhimento_interno": recolhimento_interno_dia,
            "devolucao_inconf_kg": round(devolucao_inconf_kg, 2),
            "recolhimento_sujo_kg": round(recolhimento_sujo_kg, 2),
            "ocorrencias": ocorrencias_turno
        },
        "enxoval_entregue": list(entregas_areas.values()),
        "peso_sujo": {
            "registros": lista_sujo,
            "total_hampers": round(tot_h, 2),
            "total_peso_kg": round(tot_p, 2)
        }
    }

# ==============================================================================
# 6. EXPORTAÇÕES EXCEL E CSV (MANUTENÇÃO, ROUPARIA E CONSOLIDADO)
# ==============================================================================
@app.get("/api/exportar/manutencao/excel")
def exportar_manutencao_excel(db: Session = Depends(get_db)):
    registros = db.query(Manutencao).order_by(Manutencao.id.asc()).all()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "BD_Manutencao"
    
    # Título
    ws.merge_cells("A1:Q1")
    t = ws["A1"]
    t.value = "BANCO DE DADOS DE MANUTENÇÃO - PS ANIL"
    t.font = Font(name="Segoe UI", size=13, bold=True, color="FFFFFF")
    t.fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
    t.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 30
    
    headers = [
        "ID", "DATA", "MÊS/ANO", "CATEGORIA", "TIPO MANUTENÇÃO",
        "SUB-CATEGORIA / ITEM", "SETOR / LOCAL", "DESCRIÇÃO / ESPECIFICAÇÃO",
        "QUANTIDADE / VOLUME", "HORÁRIO", "TURNO", "Nº PEDIDO / CHAMADO",
        "RESPONSÁVEL ENTREGA / VISTORIA", "RESPONSÁVEL RECEBIMENTO / ACOMPANHAMENTO",
        "STATUS / SITUAÇÃO", "DATA CONCLUSÃO / RETORNO", "RETRABALHO"
    ]
    
    ws.append(headers)
    ws.row_dimensions[2].height = 24
    
    h_font = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
    h_fill = PatternFill(start_color="2563EB", end_color="2563EB", fill_type="solid")
    for col_idx in range(1, len(headers) + 1):
        c = ws.cell(row=2, column=col_idx)
        c.font = h_font
        c.fill = h_fill
        c.alignment = Alignment(horizontal="center", vertical="center")
        
    for r in registros:
        ws.append([
            r.codigo_id or f"MANUT-{r.id}",
            r.data or "",
            r.mes_ano or "",
            r.categoria or "",
            r.tipo_manutencao or "",
            r.item or "",
            r.setor or "",
            r.descricao or "",
            r.quantidade or 1,
            r.horario or "",
            r.turno or "",
            r.numero_chamado or "",
            r.responsavel_entrega or "",
            r.responsavel_recebimento or "",
            r.status or "",
            r.data_conclusao or "",
            "Sim" if r.retrabalho == 1 else "Não"
        ])
        
    # Auto ajuste largura
    for col in ws.columns:
        col_letter = get_column_letter(col[0].column)
        max_l = max(len(str(cell.value or '')) for cell in col)
        ws.column_dimensions[col_letter].width = max(max_l + 3, 11)
        
    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    nome = f"BD_Manutencao_PS_ANIL_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    return StreamingResponse(
        buffer,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename={nome}"}
    )

@app.get("/api/exportar/consolidado/excel")
def exportar_consolidado_excel(mes_ano: str = Query("Todos"), db: Session = Depends(get_db)):
    data_cons = obter_consolidado(mes_ano, db)
    m = data_cons["manutencao"]
    r = data_cons["rouparia"]
    
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Consolidado Geral"
    
    # Cabeçalho Principal (Igual à Imagem 6)
    ws.merge_cells("A1:D1")
    t = ws["A1"]
    t.value = "PAINEL CONSOLIDADO DE INDICADORES - PS ANIL"
    t.font = Font(name="Segoe UI", size=13, bold=True, color="FFFFFF")
    t.fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
    t.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 32
    
    # Linha Mês/Ano
    ws["C2"].value = "Mês/Ano:"
    ws["C2"].font = Font(name="Segoe UI", size=10, bold=True)
    ws["D2"].value = mes_ano
    ws["D2"].font = Font(name="Segoe UI", size=10, bold=True)
    ws["D2"].alignment = Alignment(horizontal="center")
    
    # 1. Indicadores de Manutenção
    ws["A3"].value = "1. INDICADORES DE MANUTENÇÃO"
    ws["A3"].font = Font(name="Segoe UI", size=11, bold=True, color="1E3A8A")
    ws["A3"].fill = PatternFill(start_color="DBEAFE", end_color="DBEAFE", fill_type="solid")
    ws.merge_cells("A3:B3")
    
    linhas_manut = [
        ("Preventivas Realizadas", m["preventivas_realizadas"]),
        ("Preventivas Agendadas / Total", m["preventivas_agendadas"]),
        ("% Aderência Preventiva", m["aderencia_preventiva"]),
        ("Corretivas Solicitadas", m["corretivas_solicitadas"]),
        ("Corretivas Concluídas", m["corretivas_concluidas"]),
        ("% Resolução de Corretivas", m["resolucao_corretivas"]),
        ("Total Peças / Insumos Aplicados", m["total_pecas_insumos"]),
        ("Taxa de Retrabalho", m["taxa_retrabalho"])
    ]
    
    row_idx = 4
    for desc_item, val in linhas_manut:
        ws.cell(row=row_idx, column=1, value=desc_item).font = Font(name="Segoe UI", size=10)
        c_val = ws.cell(row=row_idx, column=2, value=val)
        c_val.font = Font(name="Segoe UI", size=10, bold=True)
        c_val.alignment = Alignment(horizontal="right")
        row_idx += 1
        
    # 2. Indicadores de Rouparia
    ws.cell(row=row_idx, column=1, value="2. INDICADORES DE ROUPARIA").font = Font(name="Segoe UI", size=11, bold=True, color="1E3A8A")
    ws.cell(row=row_idx, column=1).fill = PatternFill(start_color="DBEAFE", end_color="DBEAFE", fill_type="solid")
    ws.merge_cells(f"A{row_idx}:B{row_idx}")
    row_idx += 1
    
    linhas_roupa = [
        ("Total de Peças Entregues aos Setores", r["pecas_entregues"]),
        ("Total de Enxoval Sujo Recolhido (KG)", r["enxoval_sujo_kg"]),
        ("Total de Enxoval Limpo Recebido (KG)", r["enxoval_limpo_kg"]),
        ("Devoluções por Não Conformidade (KG / Peças)", r["devolucoes_inconformidade"]),
        ("Nível de Atendimento de Rouparia (%)", r["nivel_atendimento"]),
        ("Pontualidade na Entrega (07 as 19h)", r["pontualidade"]),
        ("Custo Total Lavanderia (R$)", r["custo_total_lavanderia"]),
        ("Custo por quilo (R$/KG)", r["custo_kg"])
    ]
    
    for desc_item, val in linhas_roupa:
        ws.cell(row=row_idx, column=1, value=desc_item).font = Font(name="Segoe UI", size=10)
        c_val = ws.cell(row=row_idx, column=2, value=val)
        c_val.font = Font(name="Segoe UI", size=10, bold=True)
        c_val.alignment = Alignment(horizontal="right")
        row_idx += 1
        
    ws.column_dimensions["A"].width = 45
    ws.column_dimensions["B"].width = 20
    ws.column_dimensions["C"].width = 15
    ws.column_dimensions["D"].width = 20
    
    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    nome = f"Consolidado_Indicadores_PS_ANIL_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    return StreamingResponse(
        buffer,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename={nome}"}
    )

# ==============================================================================
# 6.1 IMPORTAÇÃO DE PLANILHA EXCEL COMPLETA (Upload de Arquivo .xlsx)
# ==============================================================================
@app.post("/api/importar/excel")
async def importar_planilha_excel(arquivo: UploadFile = File(...), db: Session = Depends(get_db)):
    if not OPENPYXL_AVAILABLE:
        raise HTTPException(status_code=500, detail="OpenPyXL não está disponível no servidor.")
    
    conteudo = await arquivo.read()
    try:
        wb = openpyxl.load_workbook(io.BytesIO(conteudo), data_only=True)
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Erro ao abrir arquivo Excel: {str(e)}")
        
    manut_count = 0
    roupa_count = 0
    
    for sheet_name in wb.sheetnames:
        s_lower = sheet_name.lower().strip()
        ws = wb[sheet_name]
        
        # Aba de Manutenção
        if "manuten" in s_lower:
            header_row = 1
            for r in range(1, 10):
                val_a = str(ws.cell(row=r, column=1).value or "").upper()
                if "ID" in val_a:
                    header_row = r
                    break
                    
            for row in range(header_row + 1, ws.max_row + 1):
                val_id = ws.cell(row=row, column=1).value
                if not val_id:
                    continue
                codigo_str = str(val_id).strip()
                if not codigo_str or codigo_str.upper() in ["TOTAL GERAL", "TOTAL"]:
                    continue
                
                raw_data = ws.cell(row=row, column=2).value
                data_str = ""
                if isinstance(raw_data, datetime):
                    data_str = raw_data.strftime("%d/%m/%Y")
                elif raw_data:
                    data_str = str(raw_data).strip()
                    
                # Ignorar linhas de totais ou sem data válida
                if not data_str or codigo_str.upper().startswith("TOTAL"):
                    continue
                    
                mes_ano_raw = ws.cell(row=row, column=3).value
                mes_ano_str = ""
                if isinstance(mes_ano_raw, datetime):
                    mes_ano_str = mes_ano_raw.strftime("%m/%Y")
                elif mes_ano_raw:
                    mes_ano_str = str(mes_ano_raw).strip()
                elif "/" in data_str:
                    try:
                        p = data_str.split("/")
                        mes_ano_str = f"{p[1]}/{p[2]}"
                    except:
                        pass
                        
                cat = str(ws.cell(row=row, column=4).value or "").strip()
                tipo_m = str(ws.cell(row=row, column=5).value or "Corretiva").strip()
                it = str(ws.cell(row=row, column=6).value or "").strip()
                st = str(ws.cell(row=row, column=7).value or "").strip()
                desc = str(ws.cell(row=row, column=8).value or "").strip()
                
                try:
                    qtd = float(ws.cell(row=row, column=9).value or 1.0)
                except:
                    qtd = 1.0
                    
                hor = str(ws.cell(row=row, column=10).value or "").strip()
                tur = str(ws.cell(row=row, column=11).value or "Matutino").strip()
                num_cham = str(ws.cell(row=row, column=12).value or "").strip()
                resp_ent = str(ws.cell(row=row, column=13).value or "").strip()
                resp_rec = str(ws.cell(row=row, column=14).value or "").strip()
                status_s = str(ws.cell(row=row, column=15).value or "Pendente").strip()
                
                raw_conc = ws.cell(row=row, column=16).value
                data_conc = ""
                if isinstance(raw_conc, datetime):
                    data_conc = raw_conc.strftime("%d/%m/%Y")
                elif raw_conc:
                    data_conc = str(raw_conc).strip()
                    
                obs_val = str(ws.cell(row=row, column=17).value or "").strip()
                if obs_val.lower() == "none":
                    obs_val = ""

                ret_raw = ws.cell(row=row, column=18).value
                ret_val = str(ret_raw or "").strip().lower()
                ret_int = 1 if (ret_raw is True or ret_val in ["sim", "1", "true", "s", "x"]) else 0
                
                existente = db.query(Manutencao).filter(Manutencao.codigo_id == codigo_str).first()
                if existente:
                    existente.data = data_str or existente.data
                    existente.mes_ano = mes_ano_str or existente.mes_ano
                    existente.categoria = cat or existente.categoria
                    existente.tipo_manutencao = tipo_m or existente.tipo_manutencao
                    existente.item = it or existente.item
                    existente.setor = st or existente.setor
                    existente.descricao = desc or existente.descricao
                    existente.quantidade = qtd
                    existente.horario = hor or existente.horario
                    existente.turno = tur or existente.turno
                    existente.numero_chamado = num_cham or existente.numero_chamado
                    existente.responsavel_entrega = resp_ent or existente.responsavel_entrega
                    existente.responsavel_recebimento = resp_rec or existente.responsavel_recebimento
                    existente.status = status_s or existente.status
                    existente.data_conclusao = data_conc or existente.data_conclusao
                    existente.observacoes = obs_val or existente.observacoes
                    existente.retrabalho = ret_int
                else:
                    novo_m = Manutencao(
                        codigo_id=codigo_str,
                        data=data_str,
                        mes_ano=mes_ano_str,
                        horario=hor,
                        turno=tur,
                        categoria=cat,
                        tipo_manutencao=tipo_m,
                        item=it,
                        setor=st,
                        descricao=desc,
                        quantidade=qtd,
                        numero_chamado=num_cham,
                        responsavel_entrega=resp_ent,
                        responsavel_recebimento=resp_rec,
                        status=status_s,
                        data_conclusao=data_conc,
                        observacoes=obs_val,
                        retrabalho=ret_int
                    )
                    db.add(novo_m)
                manut_count += 1
                
        # Aba de Rouparia
        elif "roupa" in s_lower:
            header_row = 1
            for r in range(1, 10):
                val_a = str(ws.cell(row=r, column=1).value or "").upper()
                if "ID" in val_a:
                    header_row = r
                    break
                    
            for row in range(header_row + 1, ws.max_row + 1):
                val_id = ws.cell(row=row, column=1).value
                if not val_id:
                    continue
                codigo_str = str(val_id).strip()
                if not codigo_str or codigo_str.upper() in ["TOTAL GERAL", "TOTAL"]:
                    continue
                    
                raw_data = ws.cell(row=row, column=2).value
                data_str = ""
                if isinstance(raw_data, datetime):
                    data_str = raw_data.strftime("%d/%m/%Y")
                elif raw_data:
                    data_str = str(raw_data).strip()
                    
                # Ignorar linhas de resumos/totais do rodapé da planilha de rouparia
                if not (re.match(r'^R-\d+', codigo_str, re.IGNORECASE) and data_str):
                    continue
                    
                mes_ano_str = ""
                if "/" in data_str:
                    try:
                        p = data_str.split("/")
                        mes_ano_str = f"{p[1]}/{p[2]}"
                    except:
                        pass
                        
                hor = str(ws.cell(row=row, column=3).value or "").strip()
                tur = str(ws.cell(row=row, column=4).value or "Matutino").strip()
                tipo_r = str(ws.cell(row=row, column=5).value or "Estoque").strip()
                st = str(ws.cell(row=row, column=6).value or "Unidade").strip()
                item_enx = str(ws.cell(row=row, column=7).value or "").strip()
                
                try:
                    qtd_p = float(ws.cell(row=row, column=8).value or 0.0)
                except:
                    qtd_p = 0.0
                    
                try:
                    qtd_h = float(ws.cell(row=row, column=9).value or 0.0)
                except:
                    qtd_h = 0.0
                    
                try:
                    peso = float(ws.cell(row=row, column=10).value or 0.0)
                except:
                    peso = 0.0
                    
                resp_u = str(ws.cell(row=row, column=11).value or "").strip()
                resp_t = str(ws.cell(row=row, column=12).value or "LAVARE").strip()
                inconf = str(ws.cell(row=row, column=13).value or "").strip()
                
                existente = db.query(Rouparia).filter(Rouparia.codigo_id == codigo_str).first()
                if existente:
                    existente.data = data_str or existente.data
                    existente.mes_ano = mes_ano_str or existente.mes_ano
                    existente.horario = hor or existente.horario
                    existente.turno = tur or existente.turno
                    existente.tipo_registro = tipo_r or existente.tipo_registro
                    existente.setor = st or existente.setor
                    existente.item_enxoval = item_enx or existente.item_enxoval
                    existente.qtd_pecas = qtd_p
                    existente.qtd_hamper = qtd_h
                    existente.peso_kg = peso
                    existente.responsavel_unidade = resp_u or existente.responsavel_unidade
                    existente.responsavel_terceirizada = resp_t or existente.responsavel_terceirizada
                    existente.inconformidade = inconf or existente.inconformidade
                else:
                    novo_r = Rouparia(
                        codigo_id=codigo_str,
                        data=data_str,
                        mes_ano=mes_ano_str,
                        horario=hor,
                        turno=tur,
                        tipo_registro=tipo_r,
                        item_enxoval=item_enx,
                        setor=st,
                        qtd_pecas=qtd_p,
                        qtd_hamper=qtd_h,
                        peso_kg=peso,
                        responsavel_unidade=resp_u,
                        responsavel_terceirizada=resp_t,
                        inconformidade=inconf,
                        pontual=1
                    )
                    db.add(novo_r)
                roupa_count += 1
                
    db.commit()
    return {
        "status": "sucesso",
        "mensagem": f"Importação concluída! {manut_count} manutenções e {roupa_count} registros de rouparia processados.",
        "manutencoes_importadas": manut_count,
        "rouparias_importadas": roupa_count
    }

# ==============================================================================
# 7. FRONT-END INTERATIVO (SPA COMPLETA COM 5 ABAS E DESIGN OFICIAL)
# ==============================================================================
@app.get("/", response_class=HTMLResponse)
def painel_principal():
    return """
    <!DOCTYPE html>
    <html lang="pt-BR">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>SaaS Gestão Integrada - PS Anil</title>
        <script src="https://cdn.tailwindcss.com"></script>
        <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
        <link rel="preconnect" href="https://fonts.googleapis.com">
        <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
        <link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap" rel="stylesheet">
        <style>
            body { font-family: 'Plus Jakarta Sans', sans-serif; }
            .tab-btn.active {
                border-color: #2563eb;
                color: #1e3a8a;
                background-color: #eff6ff;
                font-weight: 700;
            }
        </style>
    </head>
    <body class="bg-slate-100 text-slate-800 min-h-screen flex flex-col font-sans">
        
        <!-- TELA DE LOGIN (ACESSO RESTRITO) -->
        <div id="telaLogin" class="min-h-screen flex items-center justify-center p-4 bg-gradient-to-br from-slate-900 via-blue-950 to-slate-900">
            <div class="max-w-md w-full bg-white rounded-3xl shadow-2xl border border-slate-200 overflow-hidden transform transition-all">
                <div class="bg-gradient-to-r from-blue-900 via-blue-800 to-indigo-900 text-white p-8 text-center relative overflow-hidden">
                    <div class="w-16 h-16 mx-auto mb-3 rounded-2xl bg-white/10 backdrop-blur-md flex items-center justify-center border border-white/20 shadow-inner">
                        <svg class="w-9 h-9 text-white" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                            <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M19 21V5a2 2 0 00-2-2H7a2 2 0 00-2 2v16m14 0h2m-2 0h-5m-9 0H3m2 0h5M9 7h1m-1 4h1m4-4h1m-1 4h1m-5 10v-5a1 1 0 011-1h2a1 1 0 011 1v5m-4 0h4"></path>
                        </svg>
                    </div>
                    <h2 class="text-2xl font-black tracking-tight">PS ANIL</h2>
                    <p class="text-xs text-blue-200 font-semibold uppercase tracking-wider mt-1">SEMUS - São Luís</p>
                    <p class="text-xs text-blue-300 mt-2">Sistema Integrado de Manutenção & Rouparia</p>
                </div>

                <form id="formLogin" onsubmit="executarLogin(event)" class="p-8 space-y-5">
                    <div>
                        <label class="block text-xs font-bold text-slate-700 uppercase mb-1.5">Usuário</label>
                        <div class="relative">
                            <span class="absolute inset-y-0 left-0 pl-3 flex items-center pointer-events-none text-slate-400">
                                <svg class="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M16 7a4 4 0 11-8 0 4 4 0 018 0zM12 14a7 7 0 00-7 7h14a7 7 0 00-7-7z"></path></svg>
                            </span>
                            <input type="text" id="loginUsuario" required placeholder="Digite seu usuário (ex: admin)..." autocapitalize="none" autocorrect="off" spellcheck="false" class="w-full pl-10 pr-4 py-2.5 text-sm bg-slate-50 border border-slate-300 rounded-xl focus:ring-2 focus:ring-blue-600 focus:bg-white focus:outline-none transition font-medium">
                        </div>
                    </div>

                    <div>
                        <label class="block text-xs font-bold text-slate-700 uppercase mb-1.5">Senha de Acesso</label>
                        <div class="relative">
                            <span class="absolute inset-y-0 left-0 pl-3 flex items-center pointer-events-none text-slate-400">
                                <svg class="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 15v2m-6 4h12a2 2 0 002-2v-6a2 2 0 00-2-2H6a2 2 0 00-2 2v6a2 2 0 002 2zm10-10V7a4 4 0 00-8 0v4h8z"></path></svg>
                            </span>
                            <input type="password" id="loginSenha" required placeholder="••••••••" class="w-full pl-10 pr-10 py-2.5 text-sm bg-slate-50 border border-slate-300 rounded-xl focus:ring-2 focus:ring-blue-600 focus:bg-white focus:outline-none transition">
                            <button type="button" onclick="alternarVerSenha()" class="absolute inset-y-0 right-0 pr-3 flex items-center text-slate-400 hover:text-slate-600 text-xs" title="Ver / Ocultar Senha">
                                👁️
                            </button>
                        </div>
                    </div>

                    <div id="loginErro" class="hidden p-3 rounded-xl bg-rose-50 border border-rose-200 text-xs font-semibold text-rose-700 flex items-center gap-2">
                        <span>⚠️</span> <span id="loginErroMsg">Usuário ou senha incorretos.</span>
                    </div>

                    <button type="submit" id="btnLogin" class="w-full py-3 bg-blue-900 hover:bg-blue-800 text-white font-black text-sm rounded-xl shadow-lg shadow-blue-900/30 transition flex items-center justify-center gap-2 active:scale-95">
                        <span>🔐</span> Entrar no Sistema
                    </button>

                    <div class="p-3.5 bg-blue-50 border border-blue-200 rounded-xl space-y-2 text-center">
                        <div>
                            <p class="text-[11px] text-blue-950 font-bold uppercase tracking-wider">Acesso Rápido com 1 Clique:</p>
                        </div>
                        <div class="flex gap-2">
                            <button type="button" onclick="preencherLogin('admin', 'anil2026')" class="flex-1 py-1.5 px-2 bg-blue-900 hover:bg-blue-800 text-white font-bold text-xs rounded-lg transition shadow-sm active:scale-95">
                                👤 Admin
                            </button>
                            <button type="button" onclick="preencherLogin('plantao', 'psanil')" class="flex-1 py-1.5 px-2 bg-indigo-600 hover:bg-indigo-700 text-white font-bold text-xs rounded-lg transition shadow-sm active:scale-95">
                                👨‍⚕️ Plantão
                            </button>
                        </div>
                        <p class="text-[10px] text-slate-500">Admin: admin / anil2026 • Plantão: plantao / psanil</p>
                    </div>
                </form>
            </div>
        </div>

        <!-- SISTEMA PRINCIPAL (Exibido após login) -->
        <div id="telaSistema" class="hidden min-h-screen flex flex-col">
            
            <!-- CABEÇALHO OFICIAL COM IDENTIDADE PS ANIL / SEMUS -->
            <header class="bg-white border-b border-slate-200 sticky top-0 z-40 shadow-sm">
                <div class="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-3 flex flex-wrap justify-between items-center gap-4">
                    <div class="flex items-center gap-3">
                        <div class="w-11 h-11 rounded-xl bg-gradient-to-tr from-blue-700 to-sky-500 flex items-center justify-center text-white shadow-md shadow-blue-500/20">
                            <svg class="w-6 h-6" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M19 21V5a2 2 0 00-2-2H7a2 2 0 00-2 2v16m14 0h2m-2 0h-5m-9 0H3m2 0h5M9 7h1m-1 4h1m4-4h1m-1 4h1m-5 10v-5a1 1 0 011-1h2a1 1 0 011 1v5m-4 0h4"></path>
                            </svg>
                        </div>
                        <div>
                            <div class="flex items-center gap-2">
                                <h1 class="text-xl font-extrabold text-slate-900 tracking-tight">PS ANIL</h1>
                                <span class="bg-blue-100 text-blue-900 text-[11px] font-bold px-2.5 py-0.5 rounded-full border border-blue-200">SISTEMA INTEGRADO</span>
                                <span class="text-xs text-slate-400 font-semibold hidden sm:inline">| SEMUS São Luís</span>
                            </div>
                            <p class="text-xs text-slate-500">Módulo de Manutenção Predial, Rouparia, Consolidado e Configurações</p>
                        </div>
                    </div>

                    <div class="flex items-center gap-2">
                        <input type="file" id="inputArquivoExcel" accept=".xlsx" class="hidden" onchange="enviarPlanilhaExcel(event)">
                        <button onclick="document.getElementById('inputArquivoExcel').click()" class="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-bold text-indigo-800 bg-indigo-50 hover:bg-indigo-100 border border-indigo-300 rounded-lg transition shadow-sm" title="Importar arquivo .xlsx da sua máquina para o banco">
                            <svg class="w-4 h-4 text-indigo-600" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-8l-4-4m0 0L8 8m4-4v12"></path></svg>
                            Importar Planilha (.xlsx)
                        </button>
                        <button onclick="baixarConsolidadoExcel()" class="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-bold text-emerald-800 bg-emerald-50 hover:bg-emerald-100 border border-emerald-300 rounded-lg transition shadow-sm">
                            <svg class="w-4 h-4 text-emerald-600" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 10v6m0 0l-3-3m3 3l3-3m2 8H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z"></path></svg>
                            Exportar Consolidado (.xlsx)
                        </button>
                        <a href="/api/exportar/manutencao/excel" target="_blank" class="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-bold text-blue-800 bg-blue-50 hover:bg-blue-100 border border-blue-300 rounded-lg transition shadow-sm">
                            <svg class="w-4 h-4 text-blue-600" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-4l-4 4m0 0l-4-4m4 4V4"></path></svg>
                            BD Manutenção (.xlsx)
                        </a>

                        <!-- Perfil e Logout -->
                        <div class="flex items-center gap-2 border-l border-slate-300 pl-3 ml-1">
                            <span class="text-xs font-bold text-slate-700 flex items-center gap-1.5" id="headerUsuarioBadge">
                                <span class="w-2 h-2 rounded-full bg-emerald-500 inline-block animate-pulse"></span>
                                <span id="nomeUsuarioLogado">Admin</span>
                            </span>
                            <button onclick="fazerLogout()" class="px-2.5 py-1 text-xs font-bold text-rose-700 hover:text-white hover:bg-rose-600 bg-rose-50 border border-rose-200 rounded-lg transition" title="Sair do sistema">
                                Sair
                            </button>
                        </div>
                    </div>
                </div>

                <!-- NAVEGAÇÃO ENTRE ABAS -->
                <div class="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 border-t border-slate-200 flex space-x-1 overflow-x-auto text-xs font-semibold">
                    <button onclick="mudarAba('lancamentos')" id="tabBtn-lancamentos" class="tab-btn active px-4 py-3 border-b-2 border-transparent text-slate-600 hover:text-blue-900 transition flex items-center gap-2">
                        <span>✍️</span> Painel de Lançamentos
                    </button>
                    <button onclick="mudarAba('consolidado')" id="tabBtn-consolidado" class="tab-btn px-4 py-3 border-b-2 border-transparent text-slate-600 hover:text-blue-900 transition flex items-center gap-2">
                        <span>📊</span> Consolidado de Indicadores
                    </button>
                    <button onclick="mudarAba('relatorios')" id="tabBtn-relatorios" class="tab-btn px-4 py-3 border-b-2 border-transparent text-slate-600 hover:text-blue-900 transition flex items-center gap-2">
                        <span>📈</span> Relatórios & Gráficos
                    </button>
                    <button onclick="mudarAba('historico')" id="tabBtn-historico" class="tab-btn px-4 py-3 border-b-2 border-transparent text-slate-600 hover:text-blue-900 transition flex items-center gap-2">
                        <span>🗄️</span> Base de Dados (Tabelas)
                    </button>
                    <button onclick="mudarAba('configuracoes')" id="tabBtn-configuracoes" class="tab-btn px-4 py-3 border-b-2 border-transparent text-slate-600 hover:text-blue-900 transition flex items-center gap-2">
                        <span>⚙️</span> Configurações & Cadastros
                    </button>
                </div>
            </header>

        <!-- CONTEÚDO PRINCIPAL -->
        <main class="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-6 flex-1 w-full space-y-6">

            <!-- =============================================================== -->
            <!-- ABA 1: PAINEL DE LANÇAMENTOS (IMAGENS 1 E 2)                    -->
            <!-- =============================================================== -->
            <section id="aba-lancamentos" class="space-y-6">
                <div class="bg-white rounded-2xl border border-slate-200 shadow-sm overflow-hidden max-w-4xl mx-auto">
                    <!-- Título do Formulário com Seletor de Módulo -->
                    <div class="bg-gradient-to-r from-blue-900 to-blue-800 text-white px-6 py-4 flex flex-wrap justify-between items-center gap-4">
                        <div>
                            <h2 class="text-lg font-bold">PAINEL CENTRAL DE LANÇAMENTOS - PS ANIL</h2>
                            <p class="text-xs text-blue-200">Selecione o módulo para carregar os campos específicos</p>
                        </div>
                        <div class="flex items-center gap-2 bg-blue-950/60 p-1.5 rounded-xl border border-blue-700/50">
                            <span class="text-xs font-bold text-blue-200 pl-2">Módulo / Destino:</span>
                            <select id="seletorModulo" onchange="alternarModuloLancamento()" class="text-xs font-bold text-slate-900 bg-white rounded-lg px-3 py-1.5 focus:outline-none border-0 shadow-sm cursor-pointer">
                                <option value="Manutenção">Manutenção</option>
                                <option value="Rouparia">Rouparia</option>
                            </select>
                        </div>
                    </div>

                    <!-- FORMULÁRIO MANUTENÇÃO (IDÊNTICO À IMAGEM 1) -->
                    <form id="formManutencao" onsubmit="salvarManutencao(event)" class="p-6 space-y-4">
                        <div class="grid grid-cols-1 sm:grid-cols-3 gap-4">
                            <div>
                                <label class="block text-xs font-bold text-slate-600 uppercase mb-1">Data *</label>
                                <input type="date" id="manutData" onclick="this.showPicker()" class="w-full text-xs border border-slate-300 rounded-lg p-2.5 focus:ring-2 focus:ring-blue-500 cursor-pointer bg-white" required>
                            </div>
                            <div>
                                <label class="block text-xs font-bold text-slate-600 uppercase mb-1">Horário</label>
                                <div class="flex gap-1.5">
                                    <input type="time" id="manutHorario" class="w-full text-xs border border-slate-300 rounded-lg p-2.5 focus:ring-2 focus:ring-blue-500">
                                    <button type="button" onclick="preencherHoraAtual('manutHorario')" class="px-3 py-2 bg-blue-100 hover:bg-blue-200 text-blue-900 border border-blue-300 font-bold text-xs rounded-lg transition whitespace-nowrap shadow-sm active:scale-95 flex items-center gap-1" title="Preencher com a hora atual">
                                        ⏰ Agora
                                    </button>
                                </div>
                            </div>
                            <div>
                                <label class="block text-xs font-bold text-slate-600 uppercase mb-1">Turno *</label>
                                <select id="manutTurno" class="w-full text-xs border border-slate-300 rounded-lg p-2.5 focus:ring-2 focus:ring-blue-500">
                                    <option value="Matutino">Matutino</option>
                                    <option value="Vespertino">Vespertino</option>
                                    <option value="Noturno">Noturno</option>
                                </select>
                            </div>
                        </div>

                        <div class="grid grid-cols-1 sm:grid-cols-2 gap-4">
                            <div>
                                <label class="block text-xs font-bold text-slate-600 uppercase mb-1">Categoria do Serviço *</label>
                                <select id="manutCategoria" class="w-full text-xs border border-slate-300 rounded-lg p-2.5 focus:ring-2 focus:ring-blue-500" required>
                                    <!-- Dinâmico -->
                                </select>
                            </div>
                            <div>
                                <label class="block text-xs font-bold text-slate-600 uppercase mb-1">Tipo de Manutenção *</label>
                                <select id="manutTipo" class="w-full text-xs border border-slate-300 rounded-lg p-2.5 focus:ring-2 focus:ring-blue-500 font-semibold" required>
                                    <option value="Corretiva">Corretiva</option>
                                    <option value="Preventiva">Preventiva</option>
                                </select>
                            </div>
                        </div>

                        <div class="grid grid-cols-1 sm:grid-cols-2 gap-4">
                            <div>
                                <label class="block text-xs font-bold text-slate-600 uppercase mb-1">Item / Equipamento / Insumo *</label>
                                <select id="manutItem" class="w-full text-xs border border-slate-300 rounded-lg p-2.5 focus:ring-2 focus:ring-blue-500" required>
                                    <!-- Dinâmico -->
                                </select>
                            </div>
                            <div>
                                <label class="block text-xs font-bold text-slate-600 uppercase mb-1">Setor / Local *</label>
                                <select id="manutSetor" class="w-full text-xs border border-slate-300 rounded-lg p-2.5 focus:ring-2 focus:ring-blue-500" required>
                                    <!-- Dinâmico -->
                                </select>
                            </div>
                        </div>

                        <div>
                            <label class="block text-xs font-bold text-slate-600 uppercase mb-1">Descrição / Especificação (se necessário) *</label>
                            <textarea id="manutDescricao" rows="2" placeholder="Ex: Conserto da porta de vidro da recepção, troca de bocal de tomada..." class="w-full text-xs border border-slate-300 rounded-lg p-2.5 focus:ring-2 focus:ring-blue-500" required></textarea>
                        </div>

                        <div class="grid grid-cols-1 sm:grid-cols-2 gap-4">
                            <div>
                                <label class="block text-xs font-bold text-slate-600 uppercase mb-1">Quantidade (Peças / Volume / Insumos) *</label>
                                <input type="number" id="manutQuantidade" step="1" value="1" min="1" class="w-full text-xs border border-slate-300 rounded-lg p-2.5 focus:ring-2 focus:ring-blue-500" required>
                            </div>
                            <div>
                                <label class="block text-xs font-bold text-slate-600 uppercase mb-1">Nº Chamado / Pedido</label>
                                <input type="text" id="manutNumeroChamado" placeholder="Ex: 44570" class="w-full text-xs border border-slate-300 rounded-lg p-2.5 focus:ring-2 focus:ring-blue-500">
                            </div>
                        </div>

                        <div class="grid grid-cols-1 sm:grid-cols-2 gap-4">
                            <div>
                                <label class="block text-xs font-bold text-slate-600 uppercase mb-1">Responsável Entrega / Vistoria</label>
                                <input type="text" id="manutRespEntrega" placeholder="Ex: Rosenilde Costa" class="w-full text-xs border border-slate-300 rounded-lg p-2.5 focus:ring-2 focus:ring-blue-500">
                            </div>
                            <div>
                                <label class="block text-xs font-bold text-slate-600 uppercase mb-1">Responsável Recebimento / Acompanhamento</label>
                                <input type="text" id="manutRespRecebimento" placeholder="Ex: Samira (Mastec), Téc. Anderson" class="w-full text-xs border border-slate-300 rounded-lg p-2.5 focus:ring-2 focus:ring-blue-500">
                            </div>
                        </div>

                        <div class="grid grid-cols-1 sm:grid-cols-2 gap-4">
                            <div>
                                <label class="block text-xs font-bold text-slate-600 uppercase mb-1">Status / Situação *</label>
                                <select id="manutStatus" class="w-full text-xs border border-slate-300 rounded-lg p-2.5 focus:ring-2 focus:ring-blue-500 font-bold" required>
                                    <option value="Pendente" class="text-amber-700">Pendente</option>
                                    <option value="Solicitado" class="text-blue-700">Solicitado</option>
                                    <option value="Ok" class="text-emerald-700">Ok</option>
                                    <option value="Concluído" class="text-emerald-800">Concluído</option>
                                </select>
                            </div>
                            <div>
                                <label class="block text-xs font-bold text-slate-600 uppercase mb-1">Data Conclusão / Retorno</label>
                                <input type="date" id="manutDataConclusao" onclick="this.showPicker()" class="w-full text-xs border border-slate-300 rounded-lg p-2.5 focus:ring-2 focus:ring-blue-500 cursor-pointer bg-white">
                            </div>
                        </div>

                        <div>
                            <label class="block text-xs font-bold text-slate-600 uppercase mb-1">Observações Gerais</label>
                            <textarea id="manutObservacoes" rows="2" class="w-full text-xs border border-slate-300 rounded-lg p-2.5 focus:ring-2 focus:ring-blue-500"></textarea>
                        </div>

                        <!-- Retrabalho Checkbox -->
                        <div class="bg-amber-50 border border-amber-200 rounded-xl p-3 flex items-center justify-between">
                            <div>
                                <span class="text-xs font-bold text-amber-900 block">Retrabalho (marcar se Sim)</span>
                                <span class="text-[11px] text-amber-700">Identifica se este serviço foi aberto para corrigir um serviço anterior</span>
                            </div>
                            <input type="checkbox" id="manutRetrabalho" class="w-5 h-5 text-blue-600 rounded border-slate-300 focus:ring-blue-500 cursor-pointer">
                        </div>

                        <div class="pt-2 flex justify-end">
                            <button type="submit" id="btnSalvarManutencao" class="px-6 py-3 bg-blue-900 hover:bg-blue-800 text-white font-extrabold text-sm rounded-xl shadow-md transition flex items-center gap-2">
                                <span>💾</span> GRAVAR REGISTRO
                            </button>
                        </div>
                    </form>

                    <!-- FORMULÁRIO ROUPARIA (IDÊNTICO À IMAGEM 2) -->
                    <form id="formRouparia" onsubmit="salvarRouparia(event)" class="p-6 space-y-4 hidden">
                        <div class="grid grid-cols-1 sm:grid-cols-3 gap-4">
                            <div>
                                <label class="block text-xs font-bold text-slate-600 uppercase mb-1">Data *</label>
                                <input type="date" id="roupaData" onclick="this.showPicker()" class="w-full text-xs border border-slate-300 rounded-lg p-2.5 focus:ring-2 focus:ring-blue-500 cursor-pointer bg-white" required>
                            </div>
                            <div>
                                <label class="block text-xs font-bold text-slate-600 uppercase mb-1">Horário</label>
                                <div class="flex gap-1.5">
                                    <input type="time" id="roupaHorario" class="w-full text-xs border border-slate-300 rounded-lg p-2.5 focus:ring-2 focus:ring-blue-500">
                                    <button type="button" onclick="preencherHoraAtual('roupaHorario')" class="px-3 py-2 bg-blue-100 hover:bg-blue-200 text-blue-900 border border-blue-300 font-bold text-xs rounded-lg transition whitespace-nowrap shadow-sm active:scale-95 flex items-center gap-1" title="Preencher com a hora atual">
                                        ⏰ Agora
                                    </button>
                                </div>
                            </div>
                            <div>
                                <label class="block text-xs font-bold text-slate-600 uppercase mb-1">Turno *</label>
                                <select id="roupaTurno" class="w-full text-xs border border-slate-300 rounded-lg p-2.5 focus:ring-2 focus:ring-blue-500">
                                    <option value="Matutino">Matutino</option>
                                    <option value="Vespertino">Vespertino</option>
                                    <option value="Noturno">Noturno</option>
                                </select>
                            </div>
                        </div>

                        <div class="grid grid-cols-1 sm:grid-cols-2 gap-4">
                            <div>
                                <label class="block text-xs font-bold text-slate-600 uppercase mb-1">Tipo de Movimentação *</label>
                                <select id="roupaTipo" class="w-full text-xs border border-slate-300 rounded-lg p-2.5 focus:ring-2 focus:ring-blue-500 font-semibold" required>
                                    <option value="Recebidos Limpo">Recebidos Limpo</option>
                                    <option value="Estoque">Estoque</option>
                                    <option value="Recolhimento Sujo">Recolhimento Sujo</option>
                                    <option value="Entrega aos Setores">Entrega aos Setores</option>
                                    <option value="Devolução">Devolução por Não Conformidade</option>
                                </select>
                            </div>
                            <div>
                                <label class="block text-xs font-bold text-slate-600 uppercase mb-1">Item Enxoval *</label>
                                <select id="roupaItem" class="w-full text-xs border border-slate-300 rounded-lg p-2.5 focus:ring-2 focus:ring-blue-500" required>
                                    <!-- Dinâmico -->
                                </select>
                            </div>
                        </div>

                        <div class="grid grid-cols-1 sm:grid-cols-3 gap-4">
                            <div>
                                <label class="block text-xs font-bold text-slate-600 uppercase mb-1">Setor / Local *</label>
                                <select id="roupaSetor" class="w-full text-xs border border-slate-300 rounded-lg p-2.5 focus:ring-2 focus:ring-blue-500" required>
                                    <!-- Dinâmico -->
                                </select>
                            </div>
                            <div>
                                <label class="block text-xs font-bold text-slate-600 uppercase mb-1">Quantidade (Peças / Volume)</label>
                                <input type="number" id="roupaQtdPecas" step="1" value="0" min="0" class="w-full text-xs border border-slate-300 rounded-lg p-2.5 focus:ring-2 focus:ring-blue-500">
                            </div>
                            <div>
                                <label class="block text-xs font-bold text-slate-600 uppercase mb-1">Peso (KG) / Hampers</label>
                                <input type="number" id="roupaPesoKg" step="0.01" value="0.00" min="0" class="w-full text-xs border border-slate-300 rounded-lg p-2.5 focus:ring-2 focus:ring-blue-500">
                            </div>
                        </div>

                        <div class="grid grid-cols-1 sm:grid-cols-2 gap-4">
                            <div>
                                <label class="block text-xs font-bold text-slate-600 uppercase mb-1">Responsável Unidade (Nome e Matrícula)</label>
                                <input type="text" id="roupaRespUnidade" placeholder="Ex: Rosenilde Costa, Edson" class="w-full text-xs border border-slate-300 rounded-lg p-2.5 focus:ring-2 focus:ring-blue-500">
                            </div>
                            <div>
                                <label class="block text-xs font-bold text-slate-600 uppercase mb-1">Responsável Empresa Terceirizada</label>
                                <input type="text" id="roupaRespTerceirizada" value="LAVARE" class="w-full text-xs border border-slate-300 rounded-lg p-2.5 focus:ring-2 focus:ring-blue-500">
                            </div>
                        </div>

                        <div>
                            <label class="block text-xs font-bold text-slate-600 uppercase mb-1">Inconformidade / Ocorrência (se houver)</label>
                            <textarea id="roupaInconformidade" rows="2" placeholder="Ex: Enxoval manchado, rasgado ou fora do padrão..." class="w-full text-xs border border-slate-300 rounded-lg p-2.5 focus:ring-2 focus:ring-blue-500"></textarea>
                        </div>

                        <div class="pt-2 flex justify-end">
                            <button type="submit" id="btnSalvarRouparia" class="px-6 py-3 bg-blue-900 hover:bg-blue-800 text-white font-extrabold text-sm rounded-xl shadow-md transition flex items-center gap-2">
                                <span>💾</span> GRAVAR REGISTRO
                            </button>
                        </div>
                    </form>
                </div>
            </section>

            <!-- =============================================================== -->
            <!-- ABA 2: PAINEL CONSOLIDADO DE INDICADORES (IMAGEM 6)             -->
            <!-- =============================================================== -->
            <section id="aba-consolidado" class="space-y-6 hidden">
                <!-- Barra de Filtro de Mês/Ano -->
                <div class="bg-white p-4 rounded-2xl border border-slate-200 shadow-sm flex flex-wrap justify-between items-center gap-4">
                    <div class="flex items-center gap-3">
                        <div class="w-9 h-9 rounded-lg bg-blue-100 text-blue-800 flex items-center justify-center font-bold">
                            📅
                        </div>
                        <div>
                            <h2 class="text-sm font-bold text-slate-900">Período de Análise</h2>
                            <p class="text-xs text-slate-500">Filtre os indicadores por competência mensal</p>
                        </div>
                    </div>

                    <div class="flex items-center gap-2">
                        <span class="text-xs font-bold text-slate-600">Mês / Ano:</span>
                        <select id="consolidadoFiltroMes" onchange="carregarConsolidado()" class="text-xs font-bold text-blue-900 bg-slate-50 border border-slate-300 rounded-lg px-3 py-2 focus:ring-2 focus:ring-blue-500">
                            <option value="Todos">Todos os Períodos</option>
                            <!-- Dinâmico -->
                        </select>
                        <button onclick="carregarConsolidado()" class="px-3 py-2 bg-slate-100 hover:bg-slate-200 text-slate-700 text-xs font-bold rounded-lg transition">
                            Atualizar
                        </button>
                    </div>
                </div>

                <!-- Tabela Consolidada (Estrutura idêntica à Imagem 6) -->
                <div class="bg-white rounded-2xl border border-slate-200 shadow-sm overflow-hidden">
                    <div class="bg-blue-900 text-white px-6 py-4 flex justify-between items-center">
                        <h3 class="text-base font-extrabold tracking-wide">PAINEL CONSOLIDADO DE INDICADORES - PS ANIL</h3>
                        <span id="consolidadoLabelMes" class="text-xs bg-blue-800 px-3 py-1 rounded-full font-bold">Mês/Ano: Todos</span>
                    </div>

                    <div class="p-6 grid grid-cols-1 md:grid-cols-2 gap-8">
                        <!-- 1. INDICADORES DE MANUTENÇÃO -->
                        <div class="space-y-4">
                            <div class="border-b-2 border-blue-600 pb-2 flex justify-between items-center">
                                <h4 class="text-sm font-extrabold text-blue-900 uppercase">1. Indicadores de Manutenção</h4>
                                <span class="text-[11px] font-semibold text-slate-400">Valores / Metas</span>
                            </div>

                            <div class="divide-y divide-slate-100 text-xs">
                                <div class="py-2.5 flex justify-between items-center">
                                    <span class="text-slate-600 font-medium">Preventivas Realizadas</span>
                                    <span id="kpiPrevRealizadas" class="font-extrabold text-slate-900 text-sm">0</span>
                                </div>
                                <div class="py-2.5 flex justify-between items-center">
                                    <span class="text-slate-600 font-medium">Preventivas Agendadas / Total</span>
                                    <span id="kpiPrevTotal" class="font-extrabold text-slate-900 text-sm">0</span>
                                </div>
                                <div class="py-2.5 flex justify-between items-center bg-blue-50/50 px-2 rounded-lg">
                                    <span class="text-blue-900 font-bold">% Aderência Preventiva</span>
                                    <span id="kpiAderenciaPrev" class="font-extrabold text-blue-700 text-base">0,00%</span>
                                </div>
                                <div class="py-2.5 flex justify-between items-center">
                                    <span class="text-slate-600 font-medium">Corretivas Solicitadas</span>
                                    <span id="kpiCorrSolicitadas" class="font-extrabold text-slate-900 text-sm">0</span>
                                </div>
                                <div class="py-2.5 flex justify-between items-center">
                                    <span class="text-slate-600 font-medium">Corretivas Concluídas</span>
                                    <span id="kpiCorrConcluidas" class="font-extrabold text-slate-900 text-sm">0</span>
                                </div>
                                <div class="py-2.5 flex justify-between items-center bg-emerald-50/50 px-2 rounded-lg">
                                    <span class="text-emerald-900 font-bold">% Resolução de Corretivas</span>
                                    <span id="kpiResolucaoCorr" class="font-extrabold text-emerald-700 text-base">0,00%</span>
                                </div>
                                <div class="py-2.5 flex justify-between items-center">
                                    <span class="text-slate-600 font-medium">Total Peças / Insumos Aplicados</span>
                                    <span id="kpiTotalInsumos" class="font-extrabold text-slate-900 text-sm">0</span>
                                </div>
                                <div class="py-2.5 flex justify-between items-center bg-amber-50/50 px-2 rounded-lg">
                                    <span class="text-amber-900 font-bold">Taxa de Retrabalho</span>
                                    <span id="kpiTaxaRetrabalho" class="font-extrabold text-amber-700 text-base">0,00%</span>
                                </div>
                            </div>
                        </div>

                        <!-- 2. INDICADORES DE ROUPARIA -->
                        <div class="space-y-4">
                            <div class="border-b-2 border-indigo-600 pb-2 flex justify-between items-center">
                                <h4 class="text-sm font-extrabold text-indigo-900 uppercase">2. Indicadores de Rouparia</h4>
                                <span class="text-[11px] font-semibold text-slate-400">Valores / Metas</span>
                            </div>

                            <div class="divide-y divide-slate-100 text-xs">
                                <div class="py-2.5 flex justify-between items-center">
                                    <span class="text-slate-600 font-medium">Total de Peças Entregues aos Setores</span>
                                    <span id="kpiPecasEntregues" class="font-extrabold text-slate-900 text-sm">0</span>
                                </div>
                                <div class="py-2.5 flex justify-between items-center">
                                    <span class="text-slate-600 font-medium">Total de Enxoval Sujo Recolhido (KG)</span>
                                    <span id="kpiSujoKg" class="font-extrabold text-slate-900 text-sm">0,00</span>
                                </div>
                                <div class="py-2.5 flex justify-between items-center">
                                    <span class="text-slate-600 font-medium">Total de Enxoval Limpo Recebido (KG)</span>
                                    <span id="kpiLimpoKg" class="font-extrabold text-slate-900 text-sm">0,00</span>
                                </div>
                                <div class="py-2.5 flex justify-between items-center">
                                    <span class="text-slate-600 font-medium">Devoluções por Não Conformidade (KG / Peças)</span>
                                    <span id="kpiDevolucoes" class="font-extrabold text-rose-700 text-sm">0</span>
                                </div>
                                <div class="py-2.5 flex justify-between items-center bg-indigo-50/50 px-2 rounded-lg">
                                    <span class="text-indigo-900 font-bold">Nível de Atendimento de Rouparia (%)</span>
                                    <span id="kpiNivelAtendimento" class="font-extrabold text-indigo-700 text-base">100,00%</span>
                                </div>
                                <div class="py-2.5 flex justify-between items-center bg-emerald-50/50 px-2 rounded-lg">
                                    <span class="text-emerald-900 font-bold">Pontualidade na Entrega (07 as 19h)</span>
                                    <span id="kpiPontualidade" class="font-extrabold text-emerald-700 text-base">100,00%</span>
                                </div>
                                <div class="py-2.5 flex justify-between items-center">
                                    <span class="text-slate-600 font-medium">Custo por quilo (R$/KG)</span>
                                    <span id="kpiCustoKg" class="font-extrabold text-slate-900 text-sm">R$ 0,00</span>
                                </div>
                                <div class="py-2.5 flex justify-between items-center bg-slate-100 px-2 rounded-lg">
                                    <span class="text-slate-900 font-bold">Custo Total Lavanderia (R$)</span>
                                    <span id="kpiCustoTotal" class="font-extrabold text-slate-900 text-base">R$ 0,00</span>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>
            </section>

            <!-- =============================================================== -->
            <!-- ABA 3: RELATÓRIOS & DOCUMENTOS OFICIAIS                          -->
            <!-- =============================================================== -->
            <section id="aba-relatorios" class="space-y-6 hidden">
                <!-- Barra Superior de Navegação das Sub-Abas e Ações -->
                <div class="bg-white p-4 rounded-2xl border border-slate-200 shadow-sm flex flex-wrap justify-between items-center gap-4">
                    <div class="flex flex-wrap items-center gap-2">
                        <button onclick="alternarSubRelatorio('manutencao')" id="btnSubRelManut" class="px-3.5 py-2 text-xs font-bold rounded-xl bg-blue-900 text-white shadow-sm transition flex items-center gap-1.5">
                            <span>📈</span> Indicadores de Manutenção
                        </button>
                        <button onclick="alternarSubRelatorio('rouparia_diario')" id="btnSubRelDiario" class="px-3.5 py-2 text-xs font-bold rounded-xl bg-slate-100 text-slate-700 hover:bg-slate-200 transition flex items-center gap-1.5">
                            <span>📋</span> Relatório Diário de Rouparia
                        </button>
                        <button onclick="alternarSubRelatorio('rouparia_enxoval')" id="btnSubRelEnxoval" class="px-3.5 py-2 text-xs font-bold rounded-xl bg-slate-100 text-slate-700 hover:bg-slate-200 transition flex items-center gap-1.5">
                            <span>📦</span> Enxoval Entregue nas Áreas (Doc 01)
                        </button>
                        <button onclick="alternarSubRelatorio('rouparia_peso')" id="btnSubRelPeso" class="px-3.5 py-2 text-xs font-bold rounded-xl bg-slate-100 text-slate-700 hover:bg-slate-200 transition flex items-center gap-1.5">
                            <span>⚖️</span> Peso do Enxoval Sujo (Doc 02)
                        </button>
                    </div>

                    <!-- Controles de Filtro e Impressão -->
                    <div class="flex items-center gap-3">
                        <!-- Filtro de Mês (Manutenção) -->
                        <div id="wrapperFiltroMesManut" class="flex items-center gap-1.5 text-xs">
                            <span class="text-slate-500 font-bold uppercase text-[10px]">Mês:</span>
                            <select id="relatorioFiltroMes" onchange="carregarRelatorios()" class="text-xs font-bold text-slate-800 bg-slate-50 border border-slate-300 rounded-lg px-3 py-1.5 focus:ring-2 focus:ring-blue-500 cursor-pointer">
                                <option value="Todos">Todos os Mêses</option>
                            </select>
                        </div>

                        <!-- Filtro de Data (Rouparia) -->
                        <div id="wrapperFiltroDataRouparia" class="hidden flex items-center gap-1.5 text-xs">
                            <span class="text-slate-500 font-bold uppercase text-[10px]">Data:</span>
                            <select id="relatorioFiltroDataRoupa" onchange="carregarDocumentosRouparia()" class="text-xs font-bold text-slate-800 bg-slate-50 border border-slate-300 rounded-lg px-3 py-1.5 focus:ring-2 focus:ring-blue-500 cursor-pointer">
                                <!-- Dinâmico -->
                            </select>
                        </div>

                        <!-- Botão de Impressão Oficial -->
                        <button onclick="imprimirRelatorioAtivo()" class="px-4 py-2 bg-emerald-700 hover:bg-emerald-800 text-white font-bold text-xs rounded-xl shadow-sm transition flex items-center gap-1.5 active:scale-95" title="Imprimir documento em formato oficial A4">
                            <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M17 17h2a2 2 0 002-2v-4a2 2 0 00-2-2H5a2 2 0 00-2 2v4a2 2 0 002 2h2m2 4h6a2 2 0 002-2v-4a2 2 0 00-2-2H9a2 2 0 00-2 2v4a2 2 0 002 2zm8-12V5a2 2 0 00-2-2H9a2 2 0 00-2 2v4h10z"></path></svg>
                            <span>Imprimir Relatório</span>
                        </button>
                    </div>
                </div>

                <!-- SUB-VISÃO 1: INDICADORES DE MANUTENÇÃO (GRÁFICOS, RANKING E MATRIZ) -->
                <div id="subview-rel-manut" class="space-y-6">
                    <!-- Cabeçalho do Relatório de Manutenção -->
                    <div class="bg-slate-50 p-3.5 rounded-xl border border-slate-200 flex justify-between items-center text-xs">
                        <div>
                            <span class="font-extrabold text-slate-800 text-sm block">Relatório de Manutenções e Insumos</span>
                            <span class="text-slate-500 text-[11px]">Distribuição por Status, Categoria, Setor e Matriz Cruzada de Itens</span>
                        </div>
                        <span class="px-2.5 py-1 bg-blue-100 text-blue-900 font-bold rounded-lg text-xs" id="badgeTotalChamadosRel">-</span>
                    </div>

                    <!-- LINHA 1: STATUS & CATEGORIAS (GRÁFICOS E TABELAS) -->
                    <div class="grid grid-cols-1 lg:grid-cols-2 gap-6">
                        <!-- Status / Situação -->
                        <div class="bg-white p-5 rounded-2xl border border-slate-200 shadow-sm space-y-4">
                            <h3 class="text-xs font-extrabold text-slate-700 uppercase tracking-wider border-b pb-2">Status / Situação das Manutenções</h3>
                            <div class="grid grid-cols-1 sm:grid-cols-2 gap-4 items-center">
                                <div class="h-56">
                                    <canvas id="chartStatus"></canvas>
                                </div>
                                <div class="overflow-hidden border border-slate-200 rounded-xl">
                                    <table class="w-full text-xs text-left">
                                        <thead class="bg-slate-50 text-slate-700 font-bold border-b">
                                            <tr>
                                                <th class="p-2.5">Status / Situação</th>
                                                <th class="p-2.5 text-right">Qtd</th>
                                            </tr>
                                        </thead>
                                        <tbody id="tabelaRelatorioStatus" class="divide-y divide-slate-100">
                                            <!-- Dinâmico -->
                                        </tbody>
                                    </table>
                                </div>
                            </div>
                        </div>

                        <!-- Manutenções por Categoria -->
                        <div class="bg-white p-5 rounded-2xl border border-slate-200 shadow-sm space-y-4">
                            <h3 class="text-xs font-extrabold text-slate-700 uppercase tracking-wider border-b pb-2">Manutenções por Categoria</h3>
                            <div class="grid grid-cols-1 sm:grid-cols-2 gap-4 items-center">
                                <div class="h-56">
                                    <canvas id="chartCategoria"></canvas>
                                </div>
                                <div class="overflow-y-auto max-h-56 border border-slate-200 rounded-xl">
                                    <table class="w-full text-xs text-left">
                                        <thead class="bg-slate-50 text-slate-700 font-bold border-b sticky top-0">
                                            <tr>
                                                <th class="p-2.5">Categoria</th>
                                                <th class="p-2.5 text-right">Qtd</th>
                                            </tr>
                                        </thead>
                                        <tbody id="tabelaRelatorioCategoria" class="divide-y divide-slate-100">
                                            <!-- Dinâmico -->
                                        </tbody>
                                    </table>
                                </div>
                            </div>
                        </div>
                    </div>

                    <!-- LINHA 2: RANKING POR SETOR / LOCAL -->
                    <div class="bg-white p-5 rounded-2xl border border-slate-200 shadow-sm space-y-4">
                        <h3 class="text-xs font-extrabold text-slate-700 uppercase tracking-wider border-b pb-2">Solicitações por Setor / Local (Ranking de Demandas)</h3>
                        <div class="overflow-x-auto">
                            <table class="w-full text-xs text-left">
                                <thead class="bg-slate-50 text-slate-700 font-bold border-b">
                                    <tr>
                                        <th class="p-2.5">Setor / Local</th>
                                        <th class="p-2.5 text-right">Qtd Solicitações</th>
                                        <th class="p-2.5 text-left">Representatividade Visual</th>
                                    </tr>
                                </thead>
                                <tbody id="tabelaRelatorioSetores" class="divide-y divide-slate-100">
                                    <!-- Dinâmico -->
                                </tbody>
                            </table>
                        </div>
                    </div>

                    <!-- LINHA 3: MATRIZ CRUZADA SETOR X SUB-CATEGORIA / ITEM (IDÊNTICO À IMAGEM 4) -->
                    <div class="bg-white p-5 rounded-2xl border border-slate-200 shadow-sm space-y-4">
                        <div class="flex justify-between items-center border-b pb-2">
                            <h3 class="text-xs font-extrabold text-slate-700 uppercase tracking-wider">Matriz Cruzada: Setor / Local × Sub-Categoria / Item</h3>
                            <span class="text-[11px] text-slate-500 font-medium">Contagem de peças e ocorrências por setor</span>
                        </div>
                        <div class="overflow-x-auto border border-slate-200 rounded-xl">
                            <table id="tabelaMatrizCruzada" class="w-full text-xs text-left border-collapse">
                                <!-- Gerado dinamicamente -->
                            </table>
                        </div>
                    </div>
                </div>

                <!-- SUB-VISÃO 2: RELATÓRIO DIÁRIO - ROUPARIA HOSPITALAR (IMAGEM 1) -->
                <div id="subview-rel-diario" class="space-y-6 hidden">
                    <div class="bg-white p-6 rounded-2xl border border-slate-200 shadow-sm space-y-5" id="printableRelatorioDiario">
                        <!-- Cabeçalho Oficial SEMUS / PS ANIL -->
                        <div class="border-b-2 border-slate-800 pb-4 flex flex-wrap justify-between items-center gap-4">
                            <div class="flex items-center gap-3">
                                <div class="w-12 h-12 bg-blue-900 rounded-xl flex items-center justify-center text-white text-xl font-black shadow-sm">
                                    🏥
                                </div>
                                <div>
                                    <span class="text-[11px] font-black uppercase text-blue-900 tracking-wider block">Prefeitura de São Luís • SEMUS</span>
                                    <h2 class="text-lg font-black text-slate-900 uppercase tracking-tight">Relatório Diário - Rouparia Hospitalar</h2>
                                    <p class="text-[11px] text-slate-500 font-bold">Secretaria Municipal de Saúde • Pronto Socorro do Anil (PS Anil)</p>
                                </div>
                            </div>
                            <div class="text-right">
                                <span class="inline-block px-3 py-1 bg-amber-100 text-amber-900 rounded-full text-xs font-black border border-amber-300">Documento Oficial</span>
                            </div>
                        </div>

                        <!-- Metadados: Data, Horário e Assinatura -->
                        <div class="grid grid-cols-1 sm:grid-cols-3 gap-3 p-3.5 bg-slate-50 rounded-xl border border-slate-200 text-xs">
                            <div>
                                <span class="block text-[10px] font-bold text-slate-500 uppercase">Data de Referência</span>
                                <span id="diarioDataRef" class="font-extrabold text-blue-900 text-sm">01/08/2026</span>
                            </div>
                            <div>
                                <span class="block text-[10px] font-bold text-slate-500 uppercase">Horário de Funcionamento</span>
                                <span class="font-bold text-slate-800">07:00 às 19:00 (Diurno)</span>
                            </div>
                            <div>
                                <span class="block text-[10px] font-bold text-slate-500 uppercase">Assinatura e Matrícula do Responsável</span>
                                <span class="font-semibold text-slate-700 italic">____________________________</span>
                            </div>
                        </div>

                        <!-- GRID LINHA 1: 1. Quantitativo Recebido (Laranja) & 7. Estoque (Verde) -->
                        <div class="grid grid-cols-1 lg:grid-cols-2 gap-5">
                            <!-- 1. Quantitativo de Roupas Recebidas da Lavanderia -->
                            <div class="border border-amber-300 rounded-xl overflow-hidden shadow-sm">
                                <div class="bg-amber-500 text-white font-extrabold px-3 py-2 text-xs flex justify-between items-center">
                                    <span>1. Quantitativo de Roupas Recebidas da Lavanderia</span>
                                    <span class="text-[10px] bg-amber-600 px-2 py-0.5 rounded-full" id="totRecebidosPecas">0 peças</span>
                                </div>
                                <div class="max-h-96 overflow-y-auto">
                                    <table class="w-full text-xs text-left">
                                        <thead class="bg-amber-50 text-amber-950 font-bold border-b border-amber-200">
                                            <tr>
                                                <th class="p-2 border-r border-amber-200">Tipo de Roupa</th>
                                                <th class="p-2 text-center w-36">Quantidade Recebida</th>
                                            </tr>
                                        </thead>
                                        <tbody id="tbodyDiarioRecebidos" class="divide-y divide-amber-100">
                                            <!-- Dinâmico -->
                                        </tbody>
                                    </table>
                                </div>
                            </div>

                            <!-- 7. Estoque -->
                            <div class="border border-emerald-400 rounded-xl overflow-hidden shadow-sm">
                                <div class="bg-emerald-600 text-white font-extrabold px-3 py-2 text-xs flex justify-between items-center">
                                    <span>7. Estoque</span>
                                    <span class="text-[10px] bg-emerald-700 px-2 py-0.5 rounded-full" id="totEstoquePecas">Saldo Geral</span>
                                </div>
                                <div class="max-h-96 overflow-y-auto">
                                    <table class="w-full text-xs text-left">
                                        <thead class="bg-emerald-50 text-emerald-950 font-bold border-b border-emerald-200">
                                            <tr>
                                                <th class="p-2 border-r border-emerald-200">Tipo de Roupa</th>
                                                <th class="p-2 text-center w-36">Quantidade</th>
                                            </tr>
                                        </thead>
                                        <tbody id="tbodyDiarioEstoque" class="divide-y divide-emerald-100">
                                            <!-- Dinâmico -->
                                        </tbody>
                                    </table>
                                </div>
                            </div>
                        </div>

                        <!-- GRID LINHA 2: 2. Intercorrências (Amarelo) & 3. Distribuição (Ciano) -->
                        <div class="grid grid-cols-1 lg:grid-cols-2 gap-5">
                            <!-- 2. Intercorrências Totais por Setor -->
                            <div class="border border-yellow-300 rounded-xl overflow-hidden shadow-sm">
                                <div class="bg-yellow-400 text-slate-900 font-extrabold px-3 py-2 text-xs">
                                    2. Intercorrências Totais por Setor
                                </div>
                                <div class="max-h-56 overflow-y-auto">
                                    <table class="w-full text-xs text-left">
                                        <thead class="bg-yellow-50 text-yellow-950 font-bold border-b border-yellow-200">
                                            <tr>
                                                <th class="p-2 border-r border-yellow-200">Setor</th>
                                                <th class="p-2">Item / Inconformidade</th>
                                                <th class="p-2 text-center w-16">Qtd</th>
                                            </tr>
                                        </thead>
                                        <tbody id="tbodyDiarioIntercorrencias" class="divide-y divide-yellow-100">
                                            <!-- Dinâmico -->
                                        </tbody>
                                    </table>
                                </div>
                            </div>

                            <!-- 3. Distribuição Total de Roupas por Setor -->
                            <div class="border border-cyan-400 rounded-xl overflow-hidden shadow-sm">
                                <div class="bg-cyan-600 text-white font-extrabold px-3 py-2 text-xs flex justify-between items-center">
                                    <span>3. Distribuição Total de Roupas por Setor</span>
                                    <span class="text-[10px] bg-cyan-700 px-2 py-0.5 rounded-full" id="totDistribuidasPecas">0 enviadas</span>
                                </div>
                                <div class="max-h-56 overflow-y-auto">
                                    <table class="w-full text-xs text-left">
                                        <thead class="bg-cyan-50 text-cyan-950 font-bold border-b border-cyan-200">
                                            <tr>
                                                <th class="p-2 border-r border-cyan-200">Setor</th>
                                                <th class="p-2 border-r border-cyan-200">Tipo de Roupa</th>
                                                <th class="p-2 text-center w-28">Qtd Enviada</th>
                                            </tr>
                                        </thead>
                                        <tbody id="tbodyDiarioDistribuicao" class="divide-y divide-cyan-100">
                                            <!-- Dinâmico -->
                                        </tbody>
                                    </table>
                                </div>
                            </div>
                        </div>

                        <!-- GRID LINHA 3: 4. Recolhimento Interno & 5/6. Devolução e Recolhimento Sujo -->
                        <div class="grid grid-cols-1 lg:grid-cols-2 gap-5">
                            <!-- 4. Recolhimento Interno -->
                            <div class="border border-blue-900 rounded-xl overflow-hidden shadow-sm">
                                <div class="bg-blue-950 text-white font-extrabold px-3 py-2 text-xs">
                                    4. Recolhimento Interno
                                </div>
                                <div class="max-h-48 overflow-y-auto">
                                    <table class="w-full text-xs text-left">
                                        <thead class="bg-slate-100 text-slate-800 font-bold border-b border-slate-300">
                                            <tr>
                                                <th class="p-2 border-r border-slate-300">Setor</th>
                                                <th class="p-2 text-center w-36">Kg Recolhido</th>
                                            </tr>
                                        </thead>
                                        <tbody id="tbodyDiarioRecolhimentoInterno" class="divide-y divide-slate-100">
                                            <!-- Dinâmico -->
                                        </tbody>
                                    </table>
                                </div>
                            </div>

                            <!-- 5 e 6: Devoluções e Recolhimento Sujo -->
                            <div class="space-y-4">
                                <div class="border border-slate-300 rounded-xl overflow-hidden shadow-sm">
                                    <div class="bg-slate-800 text-white font-extrabold px-3 py-2 text-xs">
                                        5. Devolução de Roupas Processadas que Apresentam Inconformidades (KG)
                                    </div>
                                    <div class="p-3 bg-slate-50 flex justify-between items-center text-xs">
                                        <span class="font-bold text-slate-600">Valor Total (KG):</span>
                                        <span id="diarioDevolucaoKg" class="font-extrabold text-rose-700 text-sm">0.00 kg</span>
                                    </div>
                                </div>

                                <div class="border border-blue-900 rounded-xl overflow-hidden shadow-sm">
                                    <div class="bg-blue-900 text-white font-extrabold px-3 py-2 text-xs">
                                        6. Recolhimento Sujo
                                    </div>
                                    <div class="p-3 bg-blue-50 flex justify-between items-center text-xs">
                                        <span class="font-bold text-blue-900">Valor Total (KG):</span>
                                        <span id="diarioRecolhimentoSujoKg" class="font-extrabold text-blue-950 text-sm">0.00 kg</span>
                                    </div>
                                </div>
                            </div>
                        </div>

                        <!-- 8. Ocorrências do Turno -->
                        <div class="border border-slate-200 rounded-xl p-4 bg-slate-50 space-y-2">
                            <span class="text-xs font-black text-slate-800 uppercase block">8. Ocorrências do Turno</span>
                            <div id="diarioOcorrenciasTurno" class="text-xs text-slate-700 leading-relaxed italic bg-white p-3 rounded-lg border border-slate-200 min-h-[60px]">
                                Nenhuma ocorrência registrada para este plantão.
                            </div>
                        </div>
                    </div>
                </div>

                <!-- SUB-VISÃO 3: PLANILHA DE ENXOVAL ENTREGUE NAS ÁREAS (IMAGEM 2) -->
                <div id="subview-rel-enxoval" class="space-y-4 hidden">
                    <div class="bg-white p-4 sm:p-5 rounded-2xl border border-slate-200 shadow-sm space-y-3" id="printableEnxovalEntregue">
                        <!-- Cabeçalho Oficial DOC 01 -->
                        <div class="border-b-2 border-slate-900 pb-2 flex flex-wrap justify-between items-center gap-2">
                            <div class="flex items-center gap-2.5">
                                <div class="w-10 h-10 bg-amber-500 rounded-xl flex items-center justify-center text-white text-lg font-black shadow-sm">
                                    📦
                                </div>
                                <div>
                                    <span class="text-[10px] font-black uppercase text-amber-800 tracking-wider block">Prefeitura Municipal de São Luís • SEMUS</span>
                                    <h2 class="text-base font-black text-slate-900 uppercase tracking-tight leading-tight">Planilha de Enxoval Entregue nas Áreas</h2>
                                    <p class="text-[10px] text-slate-500 font-bold">Secretaria Municipal de Saúde • Pronto Socorro do Anil (PS Anil)</p>
                                </div>
                            </div>
                            <div class="text-right text-[10px]">
                                <span class="block font-extrabold text-slate-700">DOC Nº: SEMUS/SUPDAM /Plan.Enx.Entr.Áreas./ 01</span>
                                <div class="flex items-center justify-end gap-2 mt-0.5">
                                    <span class="text-slate-500 font-semibold">Unidade: PS Anil</span>
                                    <span class="font-black text-blue-900 bg-blue-50 px-2 py-0.5 rounded border border-blue-200" id="enxovalDataRef">DATA: 01/08/2026</span>
                                </div>
                            </div>
                        </div>

                        <!-- Assinatura do Responsável pela Entrega -->
                        <div class="px-3 py-1.5 bg-slate-50 rounded-lg border border-slate-200 text-[10px] flex justify-between items-center">
                            <div>
                                <span class="font-bold text-slate-700 uppercase text-[9px]">Assinatura e Matrícula do Responsável pela Entrega do Enxoval:</span>
                                <span class="ml-2 italic text-slate-800">____________________________________________________________</span>
                            </div>
                        </div>

                        <!-- Os 4 Quadros Lado a Lado (Grid 4 Colunas para caber na folha A4) -->
                        <div class="w-full overflow-x-auto pb-1">
                            <div id="containerQuadrantesEnxoval" class="grid grid-cols-4 gap-2 min-w-[760px] lg:min-w-0" style="display: grid !important; grid-template-columns: repeat(4, minmax(0, 1fr)) !important;">
                                <!-- Gerado dinamicamente com os 4 quadros lado a lado -->
                            </div>
                        </div>
                    </div>
                </div>

                <!-- SUB-VISÃO 4: PLANILHA DE PESO DO ENXOVAL SUJO (IMAGEM 3) -->
                <div id="subview-rel-peso" class="space-y-6 hidden">
                    <div class="bg-white p-6 rounded-2xl border border-slate-200 shadow-sm space-y-5" id="printablePesoSujo">
                        <!-- Cabeçalho Oficial DOC 01 Peso Sujo -->
                        <div class="border-b-2 border-slate-800 pb-4 flex flex-wrap justify-between items-center gap-4">
                            <div class="flex items-center gap-3">
                                <div class="w-12 h-12 bg-blue-900 rounded-xl flex items-center justify-center text-white text-xl font-black shadow-sm">
                                    ⚖️
                                </div>
                                <div>
                                    <span class="text-[11px] font-black uppercase text-blue-900 tracking-wider block">Secretaria Municipal de Saúde - SEMUS • SUPDAM</span>
                                    <h2 class="text-lg font-black text-slate-900 uppercase tracking-tight">Planilha de Peso do Enxoval Sujo</h2>
                                    <p class="text-[11px] text-slate-500 font-bold">Superintendência de Apoio e Manutenção • PS Anil</p>
                                </div>
                            </div>
                            <div class="text-right">
                                <span class="block text-[11px] font-extrabold text-slate-700">DOC Nº: SEMUS/SUPDAM/Plan.Peso.Enx.Sujo./01</span>
                                <span class="text-xs font-black text-blue-900" id="pesoSujoDataRef">DATA: 01/08/2026</span>
                            </div>
                        </div>

                        <div class="grid grid-cols-1 sm:grid-cols-2 gap-3 p-3 bg-slate-50 rounded-xl border border-slate-200 text-xs">
                            <div>
                                <span class="font-bold text-slate-600 uppercase text-[10px] block">Responsável pela Pesagem (SEMUS):</span>
                                <span class="italic text-slate-800">___________________________________</span>
                            </div>
                            <div>
                                <span class="font-bold text-slate-600 uppercase text-[10px] block">Matrícula:</span>
                                <span class="italic text-slate-800">_____________________</span>
                            </div>
                        </div>

                        <!-- Tabela Oficial de Peso Sujo -->
                        <div class="border border-slate-300 rounded-xl overflow-hidden shadow-sm">
                            <div class="overflow-x-auto">
                                <table class="w-full text-xs text-left border-collapse">
                                    <thead class="bg-gradient-to-r from-blue-900 to-indigo-900 text-white font-bold border-b">
                                        <tr>
                                            <th class="p-2.5 text-center w-12 border-r border-blue-800">ORD.</th>
                                            <th class="p-2.5 border-r border-blue-800">DATA</th>
                                            <th class="p-2.5 text-center border-r border-blue-800">QTD HAMPER</th>
                                            <th class="p-2.5 text-center border-r border-blue-800">PESO (KG)</th>
                                            <th class="p-2.5 border-r border-blue-800">ASSINATURA E MATRÍCULA RESP. ENTREGA (SEMUS)</th>
                                            <th class="p-2.5">ASSINATURA RESP. EMPRESA TERCEIRIZADA (LAVARE)</th>
                                        </tr>
                                    </thead>
                                    <tbody id="tbodyPesoSujoRegistros" class="divide-y divide-slate-200">
                                        <!-- Dinâmico -->
                                    </tbody>
                                    <tfoot class="bg-amber-100 font-black border-t-2 border-slate-400 text-slate-900">
                                        <tr>
                                            <td colspan="2" class="p-2.5 text-right uppercase">TOTAL GERAL:</td>
                                            <td class="p-2.5 text-center text-blue-900 text-sm" id="pesoSujoTotalHampers">0</td>
                                            <td class="p-2.5 text-center text-emerald-800 text-sm" id="pesoSujoTotalKg">0.00 kg</td>
                                            <td colspan="2" class="p-2.5 text-slate-500 italic text-[11px]">Conferido e pesado conforme pesagem em balança</td>
                                        </tr>
                                    </tfoot>
                                </table>
                            </div>
                        </div>
                    </div>
                </div>
            </section>

            <!-- =============================================================== -->
            <!-- ABA 4: BASE DE DADOS (TABELAS DE CONSULTA E FILTROS)            -->
            <!-- =============================================================== -->
            <section id="aba-historico" class="space-y-6 hidden">
                <div class="bg-white p-5 rounded-2xl border border-slate-200 shadow-sm space-y-4">
                    <!-- Seletor de Tabela, Busca Geral e Ações -->
                    <div class="flex flex-wrap justify-between items-center gap-4">
                        <div class="flex items-center gap-2">
                            <button onclick="alternarTabelaHistorico('manutencao')" id="btnHistManut" class="px-4 py-2.5 text-xs font-bold rounded-xl bg-blue-900 text-white shadow-sm transition flex items-center gap-2">
                                <span>🛠️</span> Base Manutenção
                            </button>
                            <button onclick="alternarTabelaHistorico('rouparia')" id="btnHistRoupa" class="px-4 py-2.5 text-xs font-bold rounded-xl bg-slate-100 text-slate-700 hover:bg-slate-200 transition flex items-center gap-2">
                                <span>🧺</span> Base Rouparia
                            </button>
                        </div>

                        <div class="flex flex-wrap items-center gap-2 flex-1 max-w-lg">
                            <div class="relative w-full">
                                <span class="absolute inset-y-0 left-0 pl-3 flex items-center pointer-events-none text-slate-400">
                                    <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M21 21l-6-6m2-5a7 7 0 11-14 0 7 7 0 0114 0z"></path></svg>
                                </span>
                                <input type="text" id="filtroBuscaGeral" oninput="filtrarHistoricoComDebounce()" placeholder="Buscar por código, item, setor, chamado, responsável..." class="w-full pl-9 pr-8 py-2 text-xs bg-slate-50 border border-slate-300 rounded-xl focus:ring-2 focus:ring-blue-500 focus:bg-white focus:outline-none transition">
                                <button onclick="limparBuscaTexto()" class="absolute inset-y-0 right-0 pr-2.5 flex items-center text-slate-400 hover:text-slate-600 text-xs">✕</button>
                            </div>
                        </div>

                        <button onclick="limparTodosFiltros()" class="px-3.5 py-2 text-xs font-bold text-slate-600 hover:text-rose-600 bg-slate-100 hover:bg-rose-50 border border-slate-200 hover:border-rose-200 rounded-xl transition flex items-center gap-1.5" title="Redefinir todos os filtros aplicados">
                            <span>🔄</span> Limpar Filtros
                        </button>
                    </div>

                    <!-- FILTROS AVANÇADOS: MANUTENÇÃO -->
                    <div id="painelFiltrosManut" class="pt-3 border-t border-slate-100 grid grid-cols-2 sm:grid-cols-4 lg:grid-cols-7 gap-3 text-xs">
                        <div>
                            <label class="block font-bold text-slate-600 text-[10px] uppercase mb-1">Mês / Ano</label>
                            <select id="filtroManutMesAno" onchange="carregarHistorico()" class="w-full bg-slate-50 border border-slate-200 rounded-lg p-2 text-xs font-semibold focus:ring-2 focus:ring-blue-500 cursor-pointer">
                                <option value="Todos">Todos os Meses</option>
                            </select>
                        </div>
                        <div>
                            <label class="block font-bold text-slate-600 text-[10px] uppercase mb-1">Turno</label>
                            <select id="filtroManutTurno" onchange="carregarHistorico()" class="w-full bg-slate-50 border border-slate-200 rounded-lg p-2 text-xs font-semibold focus:ring-2 focus:ring-blue-500 cursor-pointer">
                                <option value="Todos">Todos os Turnos</option>
                                <option value="Matutino">Matutino</option>
                                <option value="Vespertino">Vespertino</option>
                                <option value="Noturno">Noturno</option>
                            </select>
                        </div>
                        <div>
                            <label class="block font-bold text-slate-600 text-[10px] uppercase mb-1">Status / Situação</label>
                            <select id="filtroManutStatus" onchange="carregarHistorico()" class="w-full bg-slate-50 border border-slate-200 rounded-lg p-2 text-xs font-semibold focus:ring-2 focus:ring-blue-500 cursor-pointer">
                                <option value="Todos">Todos os Status</option>
                                <option value="Pendente">Pendente</option>
                                <option value="Solicitado">Solicitado</option>
                                <option value="Ok">Ok</option>
                                <option value="Concluído">Concluído</option>
                            </select>
                        </div>
                        <div>
                            <label class="block font-bold text-slate-600 text-[10px] uppercase mb-1">Tipo de Manutenção</label>
                            <select id="filtroManutTipo" onchange="carregarHistorico()" class="w-full bg-slate-50 border border-slate-200 rounded-lg p-2 text-xs font-semibold focus:ring-2 focus:ring-blue-500 cursor-pointer">
                                <option value="Todos">Todos os Tipos</option>
                                <option value="Corretiva">Corretiva</option>
                                <option value="Preventiva">Preventiva</option>
                            </select>
                        </div>
                        <div>
                            <label class="block font-bold text-slate-600 text-[10px] uppercase mb-1">Categoria</label>
                            <select id="filtroManutCategoria" onchange="carregarHistorico()" class="w-full bg-slate-50 border border-slate-200 rounded-lg p-2 text-xs font-semibold focus:ring-2 focus:ring-blue-500 cursor-pointer">
                                <option value="Todas">Todas as Categorias</option>
                            </select>
                        </div>
                        <div>
                            <label class="block font-bold text-slate-600 text-[10px] uppercase mb-1">Setor / Local</label>
                            <select id="filtroManutSetor" onchange="carregarHistorico()" class="w-full bg-slate-50 border border-slate-200 rounded-lg p-2 text-xs font-semibold focus:ring-2 focus:ring-blue-500 cursor-pointer">
                                <option value="Todos">Todos os Setores</option>
                            </select>
                        </div>
                        <div>
                            <label class="block font-bold text-slate-600 text-[10px] uppercase mb-1">Retrabalho</label>
                            <select id="filtroManutRetrabalho" onchange="carregarHistorico()" class="w-full bg-slate-50 border border-slate-200 rounded-lg p-2 text-xs font-semibold focus:ring-2 focus:ring-blue-500 cursor-pointer">
                                <option value="Todos">Todos</option>
                                <option value="Sim">🔄 Apenas Retrabalhos</option>
                                <option value="Não">Sem Retrabalho</option>
                            </select>
                        </div>
                    </div>

                    <!-- FILTROS AVANÇADOS: ROUPARIA -->
                    <div id="painelFiltrosRoupa" class="pt-3 border-t border-slate-100 grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-5 gap-3 text-xs hidden">
                        <div>
                            <label class="block font-bold text-slate-600 text-[10px] uppercase mb-1">Mês / Ano</label>
                            <select id="filtroRoupaMesAno" onchange="carregarHistorico()" class="w-full bg-slate-50 border border-slate-200 rounded-lg p-2 text-xs font-semibold focus:ring-2 focus:ring-blue-500 cursor-pointer">
                                <option value="Todos">Todos os Meses</option>
                            </select>
                        </div>
                        <div>
                            <label class="block font-bold text-slate-600 text-[10px] uppercase mb-1">Tipo de Movimentação</label>
                            <select id="filtroRoupaTipo" onchange="carregarHistorico()" class="w-full bg-slate-50 border border-slate-200 rounded-lg p-2 text-xs font-semibold focus:ring-2 focus:ring-blue-500 cursor-pointer">
                                <option value="Todos">Todos os Tipos</option>
                                <option value="Recebidos Limpo">Recebidos Limpo</option>
                                <option value="Estoque">Estoque</option>
                                <option value="Recolhimento Sujo">Recolhimento Sujo</option>
                                <option value="Entrega aos Setores">Entrega aos Setores</option>
                                <option value="Devolução">Devolução por Não Conformidade</option>
                            </select>
                        </div>
                        <div>
                            <label class="block font-bold text-slate-600 text-[10px] uppercase mb-1">Item Enxoval</label>
                            <select id="filtroRoupaItem" onchange="carregarHistorico()" class="w-full bg-slate-50 border border-slate-200 rounded-lg p-2 text-xs font-semibold focus:ring-2 focus:ring-blue-500 cursor-pointer">
                                <option value="Todos">Todos os Itens</option>
                            </select>
                        </div>
                        <div>
                            <label class="block font-bold text-slate-600 text-[10px] uppercase mb-1">Setor / Local</label>
                            <select id="filtroRoupaSetor" onchange="carregarHistorico()" class="w-full bg-slate-50 border border-slate-200 rounded-lg p-2 text-xs font-semibold focus:ring-2 focus:ring-blue-500 cursor-pointer">
                                <option value="Todos">Todos os Setores</option>
                            </select>
                        </div>
                        <div>
                            <label class="block font-bold text-slate-600 text-[10px] uppercase mb-1">Turno</label>
                            <select id="filtroRoupaTurno" onchange="carregarHistorico()" class="w-full bg-slate-50 border border-slate-200 rounded-lg p-2 text-xs font-semibold focus:ring-2 focus:ring-blue-500 cursor-pointer">
                                <option value="Todos">Todos os Turnos</option>
                                <option value="Matutino">Matutino</option>
                                <option value="Vespertino">Vespertino</option>
                                <option value="Noturno">Noturno</option>
                            </select>
                        </div>
                    </div>
                </div>

                <!-- Tabela Manutenção -->
                <div id="containerHistManut" class="bg-white rounded-2xl border border-slate-200 shadow-sm overflow-hidden">
                    <div class="px-5 py-3 border-b bg-slate-50/50 flex justify-between items-center">
                        <span class="text-xs font-extrabold text-slate-700 uppercase">Registros de Manutenção</span>
                        <span id="contadorManut" class="text-xs font-bold text-slate-500">0 registros</span>
                    </div>
                    <div class="overflow-x-auto">
                        <table class="w-full text-left text-xs text-slate-600">
                            <thead class="bg-slate-100 text-slate-700 uppercase font-bold text-[11px] border-b">
                                <tr>
                                    <th class="p-3 text-center">ID</th>
                                    <th class="p-3">Data</th>
                                    <th class="p-3">Turno</th>
                                    <th class="p-3">Categoria</th>
                                    <th class="p-3">Tipo</th>
                                    <th class="p-3">Item</th>
                                    <th class="p-3">Setor</th>
                                    <th class="p-3">Qtd</th>
                                    <th class="p-3 text-center">Status</th>
                                    <th class="p-3 text-center">Ações</th>
                                </tr>
                            </thead>
                            <tbody id="tabelaManutCorpo" class="divide-y divide-slate-100">
                                <!-- Dinâmico -->
                            </tbody>
                        </table>
                    </div>
                </div>

                <!-- Tabela Rouparia -->
                <div id="containerHistRoupa" class="bg-white rounded-2xl border border-slate-200 shadow-sm overflow-hidden hidden">
                    <div class="px-5 py-3 border-b bg-slate-50/50 flex justify-between items-center">
                        <span class="text-xs font-extrabold text-slate-700 uppercase">Registros de Rouparia</span>
                        <span id="contadorRoupa" class="text-xs font-bold text-slate-500">0 registros</span>
                    </div>
                    <div class="overflow-x-auto">
                        <table class="w-full text-left text-xs text-slate-600">
                            <thead class="bg-slate-100 text-slate-700 uppercase font-bold text-[11px] border-b">
                                <tr>
                                    <th class="p-3 text-center">ID</th>
                                    <th class="p-3">Data</th>
                                    <th class="p-3">Turno</th>
                                    <th class="p-3">Tipo Movimentação</th>
                                    <th class="p-3">Item Enxoval</th>
                                    <th class="p-3">Setor</th>
                                    <th class="p-3">Peças</th>
                                    <th class="p-3">Peso (KG)</th>
                                    <th class="p-3">Resp. Unidade</th>
                                    <th class="p-3 text-center">Ações</th>
                                </tr>
                            </thead>
                            <tbody id="tabelaRoupaCorpo" class="divide-y divide-slate-100">
                                <!-- Dinâmico -->
                            </tbody>
                        </table>
                    </div>
                </div>
            </section>

            <!-- =============================================================== -->
            <!-- ABA 5: CONFIGURAÇÕES & CADASTROS                                -->
            <!-- =============================================================== -->
            <section id="aba-configuracoes" class="space-y-6 hidden">
                <div class="bg-white p-6 rounded-2xl border border-slate-200 shadow-sm space-y-6">
                    <div>
                        <h2 class="text-base font-extrabold text-slate-900">Configurações & Cadastros Base do Sistema</h2>
                        <p class="text-xs text-slate-500">Adicione ou gerencie novos itens, insumos, enxovais, setores e parâmetros que alimentam os formulários</p>
                    </div>

                    <!-- PARÂMETROS GERAIS (Custo de Lavanderia, etc.) -->
                    <div class="bg-slate-50 border border-slate-200 rounded-xl p-4 space-y-3">
                        <h3 class="text-xs font-extrabold text-blue-900 uppercase">Parâmetros de Custo & Lavanderia</h3>
                        <div class="flex flex-wrap items-center gap-4">
                            <div class="flex-1 min-w-[200px]">
                                <label class="block text-[11px] font-bold text-slate-600 uppercase mb-1">Custo por Quilo de Lavanderia (R$/KG)</label>
                                <input type="number" id="cfgCustoKg" step="0.01" class="text-xs border border-slate-300 rounded-lg p-2 w-full font-bold">
                            </div>
                            <button onclick="salvarParametroCusto()" class="mt-4 px-4 py-2 bg-blue-900 hover:bg-blue-800 text-white text-xs font-bold rounded-lg transition">
                                Salvar Parâmetro
                            </button>
                        </div>
                    </div>

                    <!-- CADASTROS BASE EM ABAS / SEÇÕES -->
                    <div class="grid grid-cols-1 md:grid-cols-2 gap-6">
                        
                        <!-- 1. Itens e Insumos de Manutenção -->
                        <div class="border border-slate-200 rounded-xl p-4 space-y-3 bg-white">
                            <div class="flex justify-between items-center border-b pb-2">
                                <h4 class="text-xs font-bold text-slate-800 uppercase">Itens & Insumos de Manutenção</h4>
                                <span id="badgeCountItens" class="text-[10px] bg-slate-100 text-slate-600 px-2 py-0.5 rounded-full font-bold">0</span>
                            </div>
                            <div class="flex gap-2">
                                <input type="text" id="novoItemManut" placeholder="Novo item / equipamento..." class="text-xs border border-slate-300 rounded-lg p-2 flex-1 focus:ring-2 focus:ring-blue-500">
                                <button onclick="adicionarConfigItem('item_manutencao', 'novoItemManut')" class="px-3 py-2 bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-bold rounded-lg transition">
                                    + Adicionar
                                </button>
                            </div>
                            <ul id="listaItensManut" class="divide-y divide-slate-100 text-xs max-h-48 overflow-y-auto">
                                <!-- Dinâmico -->
                            </ul>
                        </div>

                        <!-- 2. Itens de Enxoval (Rouparia) -->
                        <div class="border border-slate-200 rounded-xl p-4 space-y-3 bg-white">
                            <div class="flex justify-between items-center border-b pb-2">
                                <h4 class="text-xs font-bold text-slate-800 uppercase">Itens de Enxoval (Rouparia)</h4>
                                <span id="badgeCountEnxoval" class="text-[10px] bg-slate-100 text-slate-600 px-2 py-0.5 rounded-full font-bold">0</span>
                            </div>
                            <div class="flex gap-2">
                                <input type="text" id="novoItemEnxoval" placeholder="Novo item de enxoval..." class="text-xs border border-slate-300 rounded-lg p-2 flex-1 focus:ring-2 focus:ring-blue-500">
                                <button onclick="adicionarConfigItem('item_enxoval', 'novoItemEnxoval')" class="px-3 py-2 bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-bold rounded-lg transition">
                                    + Adicionar
                                </button>
                            </div>
                            <ul id="listaItensEnxoval" class="divide-y divide-slate-100 text-xs max-h-48 overflow-y-auto">
                                <!-- Dinâmico -->
                            </ul>
                        </div>

                        <!-- 3. Setores / Locais do Hospital -->
                        <div class="border border-slate-200 rounded-xl p-4 space-y-3 bg-white">
                            <div class="flex justify-between items-center border-b pb-2">
                                <h4 class="text-xs font-bold text-slate-800 uppercase">Setores & Locais do Hospital</h4>
                                <span id="badgeCountSetores" class="text-[10px] bg-slate-100 text-slate-600 px-2 py-0.5 rounded-full font-bold">0</span>
                            </div>
                            <div class="flex gap-2">
                                <input type="text" id="novoSetor" placeholder="Novo setor / local..." class="text-xs border border-slate-300 rounded-lg p-2 flex-1 focus:ring-2 focus:ring-blue-500">
                                <button onclick="adicionarConfigItem('setor', 'novoSetor')" class="px-3 py-2 bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-bold rounded-lg transition">
                                    + Adicionar
                                </button>
                            </div>
                            <ul id="listaSetores" class="divide-y divide-slate-100 text-xs max-h-48 overflow-y-auto">
                                <!-- Dinâmico -->
                            </ul>
                        </div>

                        <!-- 4. Categorias de Serviço -->
                        <div class="border border-slate-200 rounded-xl p-4 space-y-3 bg-white">
                            <div class="flex justify-between items-center border-b pb-2">
                                <h4 class="text-xs font-bold text-slate-800 uppercase">Categorias de Manutenção</h4>
                                <span id="badgeCountCategorias" class="text-[10px] bg-slate-100 text-slate-600 px-2 py-0.5 rounded-full font-bold">0</span>
                            </div>
                            <div class="flex gap-2">
                                <input type="text" id="novaCategoria" placeholder="Nova categoria..." class="text-xs border border-slate-300 rounded-lg p-2 flex-1 focus:ring-2 focus:ring-blue-500">
                                <button onclick="adicionarConfigItem('categoria_manutencao', 'novaCategoria')" class="px-3 py-2 bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-bold rounded-lg transition">
                                    + Adicionar
                                </button>
                            </div>
                            <ul id="listaCategorias" class="divide-y divide-slate-100 text-xs max-h-48 overflow-y-auto">
                                <!-- Dinâmico -->
                            </ul>
                        </div>

                    </div>
                </div>
            </section>

        </main>
        </div><!-- Fim de telaSistema -->

        <!-- MODAL: VISUALIZAR DETALHES DE MANUTENÇÃO -->
        <div id="modalVisualizarManutencao" class="fixed inset-0 z-50 hidden overflow-y-auto bg-slate-900/60 backdrop-blur-sm flex items-center justify-center p-4">
            <div class="bg-white rounded-2xl shadow-2xl border border-slate-200 max-w-2xl w-full overflow-hidden">
                <div class="px-6 py-4 bg-gradient-to-r from-slate-900 to-blue-900 text-white flex justify-between items-center">
                    <div>
                        <div class="flex items-center gap-2">
                            <span class="text-lg">📋</span>
                            <h3 class="text-base font-bold">Ficha de Detalhes do Chamado</h3>
                            <span id="viewManutCodigo" class="bg-blue-600 px-2.5 py-0.5 rounded-full text-xs font-black tracking-wider">#</span>
                        </div>
                        <p class="text-xs text-blue-200 mt-0.5" id="viewManutCategoriaSub">Manutenção</p>
                    </div>
                    <button onclick="fecharModalVisualizarManut()" class="text-white/70 hover:text-white transition text-lg p-1">✕</button>
                </div>
                
                <div class="p-6 space-y-4 max-h-[80vh] overflow-y-auto text-xs">
                    <!-- Badges de Status e Retrabalho -->
                    <div class="flex flex-wrap items-center justify-between gap-2 p-3 bg-slate-50 border border-slate-200 rounded-xl">
                        <div class="flex items-center gap-2">
                            <span class="font-bold text-slate-500 uppercase text-[10px]">Status:</span>
                            <span id="viewManutStatusBadge" class="px-2.5 py-1 rounded-full text-xs font-black bg-blue-100 text-blue-900">Pendente</span>
                        </div>
                        <div id="viewManutBadgesExtras" class="flex items-center gap-2">
                            <!-- Retrabalho e Editado badges inseridos via JS -->
                        </div>
                    </div>

                    <!-- Grid de Informações -->
                    <div class="grid grid-cols-2 sm:grid-cols-3 gap-3">
                        <div class="bg-slate-50 p-3 rounded-xl border border-slate-100">
                            <span class="block text-[10px] font-bold text-slate-400 uppercase">Data & Horário</span>
                            <span id="viewManutDataHora" class="font-bold text-slate-800 text-xs">-</span>
                        </div>
                        <div class="bg-slate-50 p-3 rounded-xl border border-slate-100">
                            <span class="block text-[10px] font-bold text-slate-400 uppercase">Turno</span>
                            <span id="viewManutTurno" class="font-bold text-slate-800 text-xs">-</span>
                        </div>
                        <div class="bg-slate-50 p-3 rounded-xl border border-slate-100">
                            <span class="block text-[10px] font-bold text-slate-400 uppercase">Tipo de Manutenção</span>
                            <span id="viewManutTipo" class="font-bold text-slate-800 text-xs">-</span>
                        </div>
                        <div class="bg-slate-50 p-3 rounded-xl border border-slate-100 sm:col-span-2">
                            <span class="block text-[10px] font-bold text-slate-400 uppercase">Item / Equipamento</span>
                            <span id="viewManutItem" class="font-bold text-blue-900 text-xs">-</span>
                        </div>
                        <div class="bg-slate-50 p-3 rounded-xl border border-slate-100">
                            <span class="block text-[10px] font-bold text-slate-400 uppercase">Setor / Local</span>
                            <span id="viewManutSetor" class="font-bold text-slate-800 text-xs">-</span>
                        </div>
                        <div class="bg-slate-50 p-3 rounded-xl border border-slate-100">
                            <span class="block text-[10px] font-bold text-slate-400 uppercase">Quantidade</span>
                            <span id="viewManutQuantidade" class="font-bold text-slate-800 text-xs">-</span>
                        </div>
                        <div class="bg-slate-50 p-3 rounded-xl border border-slate-100">
                            <span class="block text-[10px] font-bold text-slate-400 uppercase">Nº Chamado / Pedido</span>
                            <span id="viewManutChamado" class="font-bold text-blue-900 text-xs">-</span>
                        </div>
                        <div class="bg-slate-50 p-3 rounded-xl border border-slate-100">
                            <span class="block text-[10px] font-bold text-slate-400 uppercase">Conclusão / Retorno</span>
                            <span id="viewManutConclusao" class="font-bold text-emerald-800 text-xs">-</span>
                        </div>
                        <div class="bg-slate-50 p-3 rounded-xl border border-slate-100">
                            <span class="block text-[10px] font-bold text-slate-400 uppercase">Resp. Entrega / Vistoria</span>
                            <span id="viewManutRespEntrega" class="font-semibold text-slate-700">-</span>
                        </div>
                        <div class="bg-slate-50 p-3 rounded-xl border border-slate-100 sm:col-span-2">
                            <span class="block text-[10px] font-bold text-slate-400 uppercase">Resp. Recebimento / Acompanhamento</span>
                            <span id="viewManutRespRecebimento" class="font-semibold text-slate-700">-</span>
                        </div>
                    </div>

                    <!-- Descrição / Especificação -->
                    <div class="bg-slate-50 p-3.5 rounded-xl border border-slate-200">
                        <span class="block text-[10px] font-bold text-slate-500 uppercase mb-1">Descrição / Especificação do Serviço</span>
                        <p id="viewManutDescricao" class="text-slate-800 text-xs font-medium whitespace-pre-line leading-relaxed">-</p>
                    </div>

                    <!-- Observações Gerais / Histórico de Edições -->
                    <div class="bg-blue-50/50 p-3.5 rounded-xl border border-blue-100">
                        <span class="block text-[10px] font-bold text-blue-900 uppercase mb-1 flex items-center gap-1.5">
                            <span>📝</span> Observações Gerais & Histórico
                        </span>
                        <div id="viewManutObservacoes" class="text-slate-700 text-xs whitespace-pre-line leading-relaxed font-normal">-</div>
                    </div>
                </div>

                <div class="p-4 bg-slate-50 border-t border-slate-200 flex justify-between items-center">
                    <span class="text-[11px] text-slate-400 font-medium" id="viewManutRodapeRef">Registro no Banco de Dados</span>
                    <div class="flex gap-2">
                        <button type="button" onclick="fecharModalVisualizarManut()" class="px-3.5 py-2 bg-slate-200 hover:bg-slate-300 text-slate-700 font-bold text-xs rounded-xl transition">Fechar</button>
                        <button type="button" onclick="imprimirFichaManutencao()" class="px-3.5 py-2 bg-emerald-700 hover:bg-emerald-800 text-white font-bold text-xs rounded-xl shadow-sm transition flex items-center gap-1.5" title="Imprimir Ficha Oficial com Assinaturas">
                            <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M17 17h2a2 2 0 002-2v-4a2 2 0 00-2-2H5a2 2 0 00-2 2v4a2 2 0 002 2h2m2 4h6a2 2 0 002-2v-4a2 2 0 00-2-2H9a2 2 0 00-2 2v4a2 2 0 002 2zm8-12V5a2 2 0 00-2-2H9a2 2 0 00-2 2v4h10z"></path></svg>
                            <span>Imprimir Ficha</span>
                        </button>
                        <button type="button" id="btnEditarFromViewManut" class="px-3.5 py-2 bg-blue-900 hover:bg-blue-800 text-white font-bold text-xs rounded-xl shadow-sm transition flex items-center gap-1.5">
                            <span>✏️</span> Editar Chamado
                        </button>
                    </div>
                </div>
            </div>
        </div>

        <!-- MODAL: VISUALIZAR DETALHES DE ROUPARIA -->
        <div id="modalVisualizarRouparia" class="fixed inset-0 z-50 hidden overflow-y-auto bg-slate-900/60 backdrop-blur-sm flex items-center justify-center p-4">
            <div class="bg-white rounded-2xl shadow-2xl border border-slate-200 max-w-2xl w-full overflow-hidden">
                <div class="px-6 py-4 bg-gradient-to-r from-slate-900 to-indigo-900 text-white flex justify-between items-center">
                    <div>
                        <div class="flex items-center gap-2">
                            <span class="text-lg">🧺</span>
                            <h3 class="text-base font-bold">Detalhes da Movimentação de Rouparia</h3>
                            <span id="viewRoupaCodigo" class="bg-indigo-600 px-2.5 py-0.5 rounded-full text-xs font-black tracking-wider">#</span>
                        </div>
                        <p class="text-xs text-indigo-200 mt-0.5" id="viewRoupaTipoSub">Movimentação</p>
                    </div>
                    <button onclick="fecharModalVisualizarRoupa()" class="text-white/70 hover:text-white transition text-lg p-1">✕</button>
                </div>
                
                <div class="p-6 space-y-4 max-h-[80vh] overflow-y-auto text-xs">
                    <div class="grid grid-cols-2 sm:grid-cols-3 gap-3">
                        <div class="bg-slate-50 p-3 rounded-xl border border-slate-100">
                            <span class="block text-[10px] font-bold text-slate-400 uppercase">Data & Horário</span>
                            <span id="viewRoupaDataHora" class="font-bold text-slate-800 text-xs">-</span>
                        </div>
                        <div class="bg-slate-50 p-3 rounded-xl border border-slate-100">
                            <span class="block text-[10px] font-bold text-slate-400 uppercase">Turno</span>
                            <span id="viewRoupaTurno" class="font-bold text-slate-800 text-xs">-</span>
                        </div>
                        <div class="bg-slate-50 p-3 rounded-xl border border-slate-100">
                            <span class="block text-[10px] font-bold text-slate-400 uppercase">Tipo Movimentação</span>
                            <span id="viewRoupaTipoMov" class="font-bold text-blue-900 text-xs">-</span>
                        </div>
                        <div class="bg-slate-50 p-3 rounded-xl border border-slate-100 sm:col-span-2">
                            <span class="block text-[10px] font-bold text-slate-400 uppercase">Item do Enxoval</span>
                            <span id="viewRoupaItem" class="font-bold text-slate-900 text-xs">-</span>
                        </div>
                        <div class="bg-slate-50 p-3 rounded-xl border border-slate-100">
                            <span class="block text-[10px] font-bold text-slate-400 uppercase">Setor / Local</span>
                            <span id="viewRoupaSetor" class="font-bold text-slate-800 text-xs">-</span>
                        </div>
                        <div class="bg-slate-50 p-3 rounded-xl border border-slate-100">
                            <span class="block text-[10px] font-bold text-slate-400 uppercase">Quantidade de Peças</span>
                            <span id="viewRoupaQtd" class="font-bold text-slate-800 text-xs">-</span>
                        </div>
                        <div class="bg-slate-50 p-3 rounded-xl border border-slate-100">
                            <span class="block text-[10px] font-bold text-slate-400 uppercase">Peso (kg)</span>
                            <span id="viewRoupaPeso" class="font-bold text-slate-800 text-xs">-</span>
                        </div>
                        <div class="bg-slate-50 p-3 rounded-xl border border-slate-100">
                            <span class="block text-[10px] font-bold text-slate-400 uppercase">Pontualidade</span>
                            <span id="viewRoupaPontual" class="font-bold text-emerald-800 text-xs">-</span>
                        </div>
                        <div class="bg-slate-50 p-3 rounded-xl border border-slate-100 sm:col-span-1">
                            <span class="block text-[10px] font-bold text-slate-400 uppercase">Resp. Unidade</span>
                            <span id="viewRoupaRespUnidade" class="font-semibold text-slate-700">-</span>
                        </div>
                        <div class="bg-slate-50 p-3 rounded-xl border border-slate-100 sm:col-span-2">
                            <span class="block text-[10px] font-bold text-slate-400 uppercase">Resp. Terceirizada</span>
                            <span id="viewRoupaRespTerceirizada" class="font-semibold text-slate-700">-</span>
                        </div>
                    </div>

                    <!-- Inconformidade / Ocorrência / Observações -->
                    <div class="bg-amber-50/60 p-3.5 rounded-xl border border-amber-200">
                        <span class="block text-[10px] font-bold text-amber-900 uppercase mb-1 flex items-center gap-1.5">
                            <span>⚠️</span> Inconformidade / Ocorrências & Histórico
                        </span>
                        <div id="viewRoupaInconformidade" class="text-slate-800 text-xs whitespace-pre-line leading-relaxed font-medium">-</div>
                    </div>
                </div>

                <div class="p-4 bg-slate-50 border-t border-slate-200 flex justify-between items-center">
                    <span class="text-[11px] text-slate-400 font-medium">Registro de Rouparia</span>
                    <div class="flex gap-2">
                        <button type="button" onclick="fecharModalVisualizarRoupa()" class="px-3.5 py-2 bg-slate-200 hover:bg-slate-300 text-slate-700 font-bold text-xs rounded-xl transition">Fechar</button>
                        <button type="button" onclick="imprimirFichaRouparia()" class="px-3.5 py-2 bg-emerald-700 hover:bg-emerald-800 text-white font-bold text-xs rounded-xl shadow-sm transition flex items-center gap-1.5" title="Imprimir Comprovante Oficial de Movimentação">
                            <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M17 17h2a2 2 0 002-2v-4a2 2 0 00-2-2H5a2 2 0 00-2 2v4a2 2 0 002 2h2m2 4h6a2 2 0 002-2v-4a2 2 0 00-2-2H9a2 2 0 00-2 2v4a2 2 0 002 2zm8-12V5a2 2 0 00-2-2H9a2 2 0 00-2 2v4h10z"></path></svg>
                            <span>Imprimir Ficha</span>
                        </button>
                        <button type="button" id="btnEditarFromViewRoupa" class="px-3.5 py-2 bg-blue-900 hover:bg-blue-800 text-white font-bold text-xs rounded-xl shadow-sm transition flex items-center gap-1.5">
                            <span>✏️</span> Editar Movimentação
                        </button>
                    </div>
                </div>
            </div>
        </div>

        <!-- MODAL: EDITAR MANUTENÇÃO -->
        <div id="modalEditarManutencao" class="fixed inset-0 z-50 hidden overflow-y-auto bg-slate-900/60 backdrop-blur-sm flex items-center justify-center p-4">
            <div class="bg-white rounded-2xl shadow-2xl border border-slate-200 max-w-2xl w-full overflow-hidden">
                <div class="px-6 py-4 bg-gradient-to-r from-blue-900 to-indigo-900 text-white flex justify-between items-center">
                    <div>
                        <h3 class="text-base font-bold flex items-center gap-2">
                            <span>✏️</span> Editar Chamado de Manutenção <span id="editManutCodigo" class="bg-white/20 px-2.5 py-0.5 rounded-full text-xs">#</span>
                        </h3>
                        <p class="text-xs text-blue-200">Altere os dados necessários e salve para atualizar no banco de dados</p>
                    </div>
                    <button onclick="fecharModalEditarManut()" class="text-white/70 hover:text-white transition text-lg">✕</button>
                </div>
                <form id="formEditarManut" onsubmit="salvarEdicaoManutencao(event)" class="p-6 space-y-4 max-h-[80vh] overflow-y-auto text-xs">
                    <input type="hidden" id="editManutId">
                    <div class="grid grid-cols-3 gap-3">
                        <div>
                            <label class="block font-bold text-slate-700 uppercase mb-1">Data *</label>
                            <input type="date" id="editManutData" onclick="this.showPicker()" class="w-full border border-slate-300 rounded-lg p-2.5 focus:ring-2 focus:ring-blue-500 cursor-pointer bg-white" required>
                        </div>
                        <div>
                            <label class="block font-bold text-slate-700 uppercase mb-1">Horário</label>
                            <div class="flex gap-1.5">
                                <input type="time" id="editManutHorario" class="w-full border border-slate-300 rounded-lg p-2.5 focus:ring-2 focus:ring-blue-500">
                                <button type="button" onclick="preencherHoraAtual('editManutHorario')" class="px-2.5 py-2 bg-blue-100 hover:bg-blue-200 text-blue-900 border border-blue-300 font-bold text-xs rounded-lg transition whitespace-nowrap shadow-sm active:scale-95 flex items-center gap-1" title="Preencher com a hora atual">
                                    ⏰ Agora
                                </button>
                            </div>
                        </div>
                        <div>
                            <label class="block font-bold text-slate-700 uppercase mb-1">Turno</label>
                            <select id="editManutTurno" class="w-full border border-slate-300 rounded-lg p-2.5 focus:ring-2 focus:ring-blue-500">
                                <option value="Matutino">Matutino</option>
                                <option value="Vespertino">Vespertino</option>
                                <option value="Noturno">Noturno</option>
                            </select>
                        </div>
                    </div>
                    <div class="grid grid-cols-2 gap-3">
                        <div>
                            <label class="block font-bold text-slate-700 uppercase mb-1">Categoria *</label>
                            <select id="editManutCategoria" class="w-full border border-slate-300 rounded-lg p-2.5 focus:ring-2 focus:ring-blue-500" required></select>
                        </div>
                        <div>
                            <label class="block font-bold text-slate-700 uppercase mb-1">Tipo de Manutenção *</label>
                            <select id="editManutTipo" class="w-full border border-slate-300 rounded-lg p-2.5 font-bold focus:ring-2 focus:ring-blue-500">
                                <option value="Corretiva">Corretiva</option>
                                <option value="Preventiva">Preventiva</option>
                            </select>
                        </div>
                    </div>
                    <div class="grid grid-cols-2 gap-3">
                        <div>
                            <label class="block font-bold text-slate-700 uppercase mb-1">Item / Equipamento / Insumo *</label>
                            <select id="editManutItem" class="w-full border border-slate-300 rounded-lg p-2.5 focus:ring-2 focus:ring-blue-500" required></select>
                        </div>
                        <div>
                            <label class="block font-bold text-slate-700 uppercase mb-1">Setor / Local *</label>
                            <select id="editManutSetor" class="w-full border border-slate-300 rounded-lg p-2.5 focus:ring-2 focus:ring-blue-500" required></select>
                        </div>
                    </div>
                    <div>
                        <label class="block font-bold text-slate-700 uppercase mb-1">Descrição / Especificação *</label>
                        <textarea id="editManutDescricao" rows="2" class="w-full border border-slate-300 rounded-lg p-2.5 focus:ring-2 focus:ring-blue-500" required></textarea>
                    </div>
                    <div class="grid grid-cols-2 gap-3">
                        <div>
                            <label class="block font-bold text-slate-700 uppercase mb-1">Quantidade *</label>
                            <input type="number" id="editManutQuantidade" step="1" min="1" class="w-full border border-slate-300 rounded-lg p-2.5 focus:ring-2 focus:ring-blue-500" required>
                        </div>
                        <div>
                            <label class="block font-bold text-slate-700 uppercase mb-1">Nº Chamado / Pedido</label>
                            <input type="text" id="editManutNumeroChamado" class="w-full border border-slate-300 rounded-lg p-2.5 focus:ring-2 focus:ring-blue-500">
                        </div>
                    </div>
                    <div class="grid grid-cols-2 gap-3">
                        <div>
                            <label class="block font-bold text-slate-700 uppercase mb-1">Resp. Entrega / Vistoria</label>
                            <input type="text" id="editManutRespEntrega" class="w-full border border-slate-300 rounded-lg p-2.5 focus:ring-2 focus:ring-blue-500">
                        </div>
                        <div>
                            <label class="block font-bold text-slate-700 uppercase mb-1">Resp. Recebimento / Acomp.</label>
                            <input type="text" id="editManutRespRecebimento" class="w-full border border-slate-300 rounded-lg p-2.5 focus:ring-2 focus:ring-blue-500">
                        </div>
                    </div>
                    <div class="grid grid-cols-2 gap-3">
                        <div>
                            <label class="block font-bold text-slate-700 uppercase mb-1">Status / Situação *</label>
                            <select id="editManutStatus" class="w-full border border-slate-300 rounded-lg p-2.5 font-bold focus:ring-2 focus:ring-blue-500">
                                <option value="Pendente">Pendente</option>
                                <option value="Solicitado">Solicitado</option>
                                <option value="Ok">Ok</option>
                                <option value="Concluído">Concluído</option>
                            </select>
                        </div>
                        <div>
                            <label class="block font-bold text-slate-700 uppercase mb-1">Data Conclusão / Retorno</label>
                            <input type="date" id="editManutDataConclusao" onclick="this.showPicker()" class="w-full border border-slate-300 rounded-lg p-2.5 focus:ring-2 focus:ring-blue-500 cursor-pointer bg-white">
                        </div>
                    </div>
                    <div>
                        <label class="block font-bold text-slate-700 uppercase mb-1">Observações Gerais</label>
                        <textarea id="editManutObservacoes" rows="2" class="w-full border border-slate-300 rounded-lg p-2.5 focus:ring-2 focus:ring-blue-500"></textarea>
                    </div>
                    <div class="bg-amber-50 border border-amber-200 rounded-xl p-3 flex items-center justify-between">
                        <div>
                            <span class="font-bold text-amber-900 block">Retrabalho</span>
                            <span class="text-[11px] text-amber-700">Marque se este chamado for um retrabalho de serviço anterior</span>
                        </div>
                        <input type="checkbox" id="editManutRetrabalho" class="w-5 h-5 text-blue-600 rounded cursor-pointer">
                    </div>
                    <div class="pt-3 border-t flex justify-end gap-2">
                        <button type="button" onclick="fecharModalEditarManut()" class="px-4 py-2 bg-slate-100 hover:bg-slate-200 text-slate-700 font-bold rounded-lg transition">Cancelar</button>
                        <button type="submit" id="btnSalvarEditManut" class="px-5 py-2 bg-blue-900 hover:bg-blue-800 text-white font-bold rounded-lg shadow-sm transition">Salvar Alterações</button>
                    </div>
                </form>
            </div>
        </div>

        <!-- MODAL: EDITAR ROUPARIA -->
        <div id="modalEditarRouparia" class="fixed inset-0 z-50 hidden overflow-y-auto bg-slate-900/60 backdrop-blur-sm flex items-center justify-center p-4">
            <div class="bg-white rounded-2xl shadow-2xl border border-slate-200 max-w-2xl w-full overflow-hidden">
                <div class="px-6 py-4 bg-gradient-to-r from-blue-900 to-indigo-900 text-white flex justify-between items-center">
                    <div>
                        <h3 class="text-base font-bold flex items-center gap-2">
                            <span>✏️</span> Editar Movimentação de Rouparia <span id="editRoupaCodigo" class="bg-white/20 px-2.5 py-0.5 rounded-full text-xs">#</span>
                        </h3>
                        <p class="text-xs text-blue-200">Altere os dados da movimentação de enxoval e salve no banco</p>
                    </div>
                    <button onclick="fecharModalEditarRoupa()" class="text-white/70 hover:text-white transition text-lg">✕</button>
                </div>
                <form id="formEditarRoupa" onsubmit="salvarEdicaoRouparia(event)" class="p-6 space-y-4 max-h-[80vh] overflow-y-auto text-xs">
                    <input type="hidden" id="editRoupaId">
                    <div class="grid grid-cols-3 gap-3">
                        <div>
                            <label class="block font-bold text-slate-700 uppercase mb-1">Data *</label>
                            <input type="date" id="editRoupaData" onclick="this.showPicker()" class="w-full border border-slate-300 rounded-lg p-2.5 focus:ring-2 focus:ring-blue-500 cursor-pointer bg-white" required>
                        </div>
                        <div>
                            <label class="block font-bold text-slate-700 uppercase mb-1">Horário</label>
                            <div class="flex gap-1.5">
                                <input type="time" id="editRoupaHorario" class="w-full border border-slate-300 rounded-lg p-2.5 focus:ring-2 focus:ring-blue-500">
                                <button type="button" onclick="preencherHoraAtual('editRoupaHorario')" class="px-2.5 py-2 bg-blue-100 hover:bg-blue-200 text-blue-900 border border-blue-300 font-bold text-xs rounded-lg transition whitespace-nowrap shadow-sm active:scale-95 flex items-center gap-1" title="Preencher com a hora atual">
                                    ⏰ Agora
                                </button>
                            </div>
                        </div>
                        <div>
                            <label class="block font-bold text-slate-700 uppercase mb-1">Turno</label>
                            <select id="editRoupaTurno" class="w-full border border-slate-300 rounded-lg p-2.5 focus:ring-2 focus:ring-blue-500">
                                <option value="Matutino">Matutino</option>
                                <option value="Vespertino">Vespertino</option>
                                <option value="Noturno">Noturno</option>
                            </select>
                        </div>
                    </div>
                    <div class="grid grid-cols-2 gap-3">
                        <div>
                            <label class="block font-bold text-slate-700 uppercase mb-1">Tipo de Movimentação *</label>
                            <select id="editRoupaTipo" class="w-full border border-slate-300 rounded-lg p-2.5 font-bold focus:ring-2 focus:ring-blue-500">
                                <option value="Recebidos Limpo">Recebidos Limpo</option>
                                <option value="Estoque">Estoque</option>
                                <option value="Recolhimento Sujo">Recolhimento Sujo</option>
                                <option value="Entrega aos Setores">Entrega aos Setores</option>
                                <option value="Devolução">Devolução por Não Conformidade</option>
                            </select>
                        </div>
                        <div>
                            <label class="block font-bold text-slate-700 uppercase mb-1">Item Enxoval *</label>
                            <select id="editRoupaItem" class="w-full border border-slate-300 rounded-lg p-2.5 focus:ring-2 focus:ring-blue-500" required></select>
                        </div>
                    </div>
                    <div class="grid grid-cols-3 gap-3">
                        <div>
                            <label class="block font-bold text-slate-700 uppercase mb-1">Setor / Local *</label>
                            <select id="editRoupaSetor" class="w-full border border-slate-300 rounded-lg p-2.5 focus:ring-2 focus:ring-blue-500" required></select>
                        </div>
                        <div>
                            <label class="block font-bold text-slate-700 uppercase mb-1">Qtd Peças</label>
                            <input type="number" id="editRoupaQtdPecas" step="1" class="w-full border border-slate-300 rounded-lg p-2.5 focus:ring-2 focus:ring-blue-500">
                        </div>
                        <div>
                            <label class="block font-bold text-slate-700 uppercase mb-1">Peso (KG)</label>
                            <input type="number" id="editRoupaPesoKg" step="0.01" class="w-full border border-slate-300 rounded-lg p-2.5 focus:ring-2 focus:ring-blue-500">
                        </div>
                    </div>
                    <div class="grid grid-cols-2 gap-3">
                        <div>
                            <label class="block font-bold text-slate-700 uppercase mb-1">Resp. Unidade</label>
                            <input type="text" id="editRoupaRespUnidade" class="w-full border border-slate-300 rounded-lg p-2.5 focus:ring-2 focus:ring-blue-500">
                        </div>
                        <div>
                            <label class="block font-bold text-slate-700 uppercase mb-1">Resp. Terceirizada</label>
                            <input type="text" id="editRoupaRespTerceirizada" class="w-full border border-slate-300 rounded-lg p-2.5 focus:ring-2 focus:ring-blue-500">
                        </div>
                    </div>
                    <div>
                        <label class="block font-bold text-slate-700 uppercase mb-1">Inconformidade / Ocorrência</label>
                        <textarea id="editRoupaInconformidade" rows="2" class="w-full border border-slate-300 rounded-lg p-2.5 focus:ring-2 focus:ring-blue-500"></textarea>
                    </div>
                    <div class="pt-3 border-t flex justify-end gap-2">
                        <button type="button" onclick="fecharModalEditarRoupa()" class="px-4 py-2 bg-slate-100 hover:bg-slate-200 text-slate-700 font-bold rounded-lg transition">Cancelar</button>
                        <button type="submit" id="btnSalvarEditRoupa" class="px-5 py-2 bg-blue-900 hover:bg-blue-800 text-white font-bold rounded-lg shadow-sm transition">Salvar Alterações</button>
                    </div>
                </form>
            </div>
        </div>

        <!-- TOAST FEEDBACK NOTIFICAÇÃO -->
        <div id="toastBox" class="fixed bottom-6 right-6 z-50 transform translate-y-20 opacity-0 transition-all duration-300 pointer-events-none">
            <div id="toastContent" class="bg-slate-900 text-white text-xs font-semibold px-4 py-3 rounded-xl shadow-xl flex items-center gap-2 border border-slate-700">
                <span id="toastIcon">✅</span>
                <span id="toastMsg">Operação realizada com sucesso!</span>
            </div>
        </div>

        <!-- JAVASCRIPT DA SPA -->
        <script>
            let debounceTimer = null;
            let chartStatusInstance = null;
            let chartCategoriaInstance = null;

            // Funções de formatação e preenchimento de Data e Horário
            function formatarDataIsoHoje() {
                const d = new Date();
                const ano = d.getFullYear();
                const mes = String(d.getMonth() + 1).padStart(2, '0');
                const dia = String(d.getDate()).padStart(2, '0');
                return `${ano}-${mes}-${dia}`;
            }

            function obterHoraAtual() {
                const d = new Date();
                const h = String(d.getHours()).padStart(2, '0');
                const m = String(d.getMinutes()).padStart(2, '0');
                return `${h}:${m}`;
            }

            function preencherHoraAtual(inputId) {
                const el = document.getElementById(inputId);
                if (el) {
                    el.value = obterHoraAtual();
                    mostrarToast(`Horário atual definido: ${el.value}`);
                }
            }

            function dataParaIso(dStr) {
                if (!dStr) return '';
                if (dStr.includes('-')) return dStr.split('T')[0];
                const p = dStr.split('/');
                if (p.length === 3) {
                    return `${p[2]}-${p[1].padStart(2, '0')}-${p[0].padStart(2, '0')}`;
                }
                return dStr;
            }

            function dataParaBr(dStr) {
                if (!dStr) return '';
                if (dStr.includes('/')) return dStr;
                const p = dStr.split('-');
                if (p.length === 3) {
                    return `${p[2]}/${p[1]}/${p[0]}`;
                }
                return dStr;
            }

            // Toast
            function mostrarToast(msg, tipo = 'sucesso') {
                const box = document.getElementById('toastBox');
                const tMsg = document.getElementById('toastMsg');
                const tIcon = document.getElementById('toastIcon');
                tMsg.textContent = msg;
                tIcon.textContent = tipo === 'sucesso' ? '✅' : tipo === 'aviso' ? '⚠️' : '❌';
                box.classList.remove('translate-y-20', 'opacity-0');
                box.classList.add('translate-y-0', 'opacity-100');
                setTimeout(() => {
                    box.classList.remove('translate-y-0', 'opacity-100');
                    box.classList.add('translate-y-20', 'opacity-0');
                }, 3500);
            }

            // Alternar Abas Principais
            function mudarAba(abaId) {
                const abas = ['lancamentos', 'consolidado', 'relatorios', 'historico', 'configuracoes'];
                abas.forEach(a => {
                    const el = document.getElementById('aba-' + a);
                    const btn = document.getElementById('tabBtn-' + a);
                    if (a === abaId) {
                        el.classList.remove('hidden');
                        btn.classList.add('active');
                    } else {
                        el.classList.add('hidden');
                        btn.classList.remove('active');
                    }
                });

                if (abaId === 'consolidado') carregarConsolidado();
                if (abaId === 'relatorios') {
                    if (typeof subAbaRelatorioAtiva !== 'undefined' && subAbaRelatorioAtiva !== 'manutencao') {
                        carregarDocumentosRouparia();
                    } else {
                        carregarRelatorios();
                    }
                }
                if (abaId === 'historico') carregarHistorico();
                if (abaId === 'configuracoes') carregarConfiguracoes();
            }

            // Alternar Módulo de Lançamento (Manutenção vs Rouparia)
            function alternarModuloLancamento() {
                const mod = document.getElementById('seletorModulo').value;
                const fManut = document.getElementById('formManutencao');
                const fRoupa = document.getElementById('formRouparia');

                if (mod === 'Manutenção') {
                    fManut.classList.remove('hidden');
                    fRoupa.classList.add('hidden');
                } else {
                    fManut.classList.add('hidden');
                    fRoupa.classList.remove('hidden');
                }
            }

            // Carregar Dropdowns de Configuração
            async function carregarDropdownsFormulario() {
                try {
                    const res = await fetch('/api/configuracoes');
                    const itens = await res.json();

                    const selCat = document.getElementById('manutCategoria');
                    const selItemManut = document.getElementById('manutItem');
                    const selSetorManut = document.getElementById('manutSetor');
                    const selItemRoupa = document.getElementById('roupaItem');
                    const selSetorRoupa = document.getElementById('roupaSetor');

                    const editCat = document.getElementById('editManutCategoria');
                    const editItemManut = document.getElementById('editManutItem');
                    const editSetorManut = document.getElementById('editManutSetor');
                    const editItemRoupa = document.getElementById('editRoupaItem');
                    const editSetorRoupa = document.getElementById('editRoupaSetor');

                    // Filtros da Base de Dados (Aba 4)
                    const fCat = document.getElementById('filtroManutCategoria');
                    const fSetorManut = document.getElementById('filtroManutSetor');
                    const fItemRoupa = document.getElementById('filtroRoupaItem');
                    const fSetorRoupa = document.getElementById('filtroRoupaSetor');

                    selCat.innerHTML = '';
                    selItemManut.innerHTML = '';
                    selSetorManut.innerHTML = '';
                    selItemRoupa.innerHTML = '';
                    selSetorRoupa.innerHTML = '';

                    if (editCat) editCat.innerHTML = '';
                    if (editItemManut) editItemManut.innerHTML = '';
                    if (editSetorManut) editSetorManut.innerHTML = '';
                    if (editItemRoupa) editItemRoupa.innerHTML = '';
                    if (editSetorRoupa) editSetorRoupa.innerHTML = '';

                    if (fCat) fCat.innerHTML = '<option value="Todas">Todas as Categorias</option>';
                    if (fSetorManut) fSetorManut.innerHTML = '<option value="Todos">Todos os Setores</option>';
                    if (fItemRoupa) fItemRoupa.innerHTML = '<option value="Todos">Todos os Itens</option>';
                    if (fSetorRoupa) fSetorRoupa.innerHTML = '<option value="Todos">Todos os Setores</option>';

                    itens.forEach(i => {
                        const opt = `<option value="${i.nome}">${i.nome}</option>`;
                        if (i.tipo === 'categoria_manutencao') {
                            selCat.innerHTML += opt;
                            if (editCat) editCat.innerHTML += opt;
                            if (fCat) fCat.innerHTML += opt;
                        }
                        if (i.tipo === 'item_manutencao') {
                            selItemManut.innerHTML += opt;
                            if (editItemManut) editItemManut.innerHTML += opt;
                        }
                        if (i.tipo === 'setor') {
                            selSetorManut.innerHTML += opt;
                            selSetorRoupa.innerHTML += opt;
                            if (editSetorManut) editSetorManut.innerHTML += opt;
                            if (editSetorRoupa) editSetorRoupa.innerHTML += opt;
                            if (fSetorManut) fSetorManut.innerHTML += opt;
                            if (fSetorRoupa) fSetorRoupa.innerHTML += opt;
                        }
                        if (i.tipo === 'item_enxoval') {
                            selItemRoupa.innerHTML += opt;
                            if (editItemRoupa) editItemRoupa.innerHTML += opt;
                            if (fItemRoupa) fItemRoupa.innerHTML += opt;
                        }
                    });

                    // Setar data de hoje nos inputs com calendário e horário atual
                    const hojeIso = formatarDataIsoHoje();
                    const horaAtual = obterHoraAtual();
                    document.getElementById('manutData').value = hojeIso;
                    document.getElementById('manutHorario').value = horaAtual;
                    document.getElementById('roupaData').value = hojeIso;
                    document.getElementById('roupaHorario').value = horaAtual;

                } catch (e) {
                    console.error('Erro ao carregar dropdowns:', e);
                }
            }

            // Salvar Manutenção
            async function salvarManutencao(e) {
                e.preventDefault();
                const btn = document.getElementById('btnSalvarManutencao');
                btn.disabled = true;

                const payload = {
                    data: dataParaBr(document.getElementById('manutData').value.trim()),
                    horario: document.getElementById('manutHorario').value.trim(),
                    turno: document.getElementById('manutTurno').value,
                    categoria: document.getElementById('manutCategoria').value,
                    tipo_manutencao: document.getElementById('manutTipo').value,
                    item: document.getElementById('manutItem').value,
                    setor: document.getElementById('manutSetor').value,
                    descricao: document.getElementById('manutDescricao').value.trim(),
                    quantidade: parseFloat(document.getElementById('manutQuantidade').value) || 1.0,
                    numero_chamado: document.getElementById('manutNumeroChamado').value.trim(),
                    responsavel_entrega: document.getElementById('manutRespEntrega').value.trim(),
                    responsavel_recebimento: document.getElementById('manutRespRecebimento').value.trim(),
                    status: document.getElementById('manutStatus').value,
                    data_conclusao: dataParaBr(document.getElementById('manutDataConclusao').value.trim()),
                    observacoes: document.getElementById('manutObservacoes').value.trim(),
                    retrabalho: document.getElementById('manutRetrabalho').checked ? 1 : 0
                };

                try {
                    const res = await fetch('/api/manutencao', {
                        method: 'POST',
                        headers: {'Content-Type': 'application/json'},
                        body: JSON.stringify(payload)
                    });
                    const d = await res.json();
                    if (res.ok) {
                        mostrarToast(d.mensagem);
                        document.getElementById('formManutencao').reset();
                        carregarDropdownsFormulario();
                    } else {
                        mostrarToast('Erro ao salvar registro de manutenção.', 'erro');
                    }
                } catch (err) {
                    mostrarToast('Erro de conexão.', 'erro');
                } finally {
                    btn.disabled = false;
                }
            }

            // Salvar Rouparia
            async function salvarRouparia(e) {
                e.preventDefault();
                const btn = document.getElementById('btnSalvarRouparia');
                btn.disabled = true;

                const payload = {
                    data: dataParaBr(document.getElementById('roupaData').value.trim()),
                    horario: document.getElementById('roupaHorario').value.trim(),
                    turno: document.getElementById('roupaTurno').value,
                    tipo_registro: document.getElementById('roupaTipo').value,
                    item_enxoval: document.getElementById('roupaItem').value,
                    setor: document.getElementById('roupaSetor').value,
                    qtd_pecas: parseFloat(document.getElementById('roupaQtdPecas').value) || 0.0,
                    peso_kg: parseFloat(document.getElementById('roupaPesoKg').value) || 0.0,
                    responsavel_unidade: document.getElementById('roupaRespUnidade').value.trim(),
                    responsavel_terceirizada: document.getElementById('roupaRespTerceirizada').value.trim(),
                    inconformidade: document.getElementById('roupaInconformidade').value.trim()
                };

                try {
                    const res = await fetch('/api/rouparia', {
                        method: 'POST',
                        headers: {'Content-Type': 'application/json'},
                        body: JSON.stringify(payload)
                    });
                    const d = await res.json();
                    if (res.ok) {
                        mostrarToast(d.mensagem);
                        document.getElementById('formRouparia').reset();
                        carregarDropdownsFormulario();
                    } else {
                        mostrarToast('Erro ao salvar movimentação de rouparia.', 'erro');
                    }
                } catch (err) {
                    mostrarToast('Erro de conexão.', 'erro');
                } finally {
                    btn.disabled = false;
                }
            }

            // Carregar Consolidado de Indicadores
            async function carregarConsolidado() {
                const mes = document.getElementById('consolidadoFiltroMes').value;
                try {
                    const res = await fetch(`/api/consolidado?mes_ano=${encodeURIComponent(mes)}`);
                    const d = await res.json();

                    // Preencher Dropdown de Meses se necessário
                    const selMes = document.getElementById('consolidadoFiltroMes');
                    const selRel = document.getElementById('relatorioFiltroMes');
                    const fMesManut = document.getElementById('filtroManutMesAno');
                    const fMesRoupa = document.getElementById('filtroRoupaMesAno');
                    if (d.meses_disponiveis) {
                        if (selMes && selMes.options.length <= 1) {
                            d.meses_disponiveis.forEach(m => {
                                const opt = `<option value="${m}">${m}</option>`;
                                selMes.innerHTML += opt;
                                if (selRel) selRel.innerHTML += opt;
                            });
                            selMes.value = d.mes_ano_selecionado;
                        }
                        if (fMesManut && fMesManut.options.length <= 1) {
                            d.meses_disponiveis.forEach(m => {
                                fMesManut.innerHTML += `<option value="${m}">${m}</option>`;
                            });
                        }
                        if (fMesRoupa && fMesRoupa.options.length <= 1) {
                            d.meses_disponiveis.forEach(m => {
                                fMesRoupa.innerHTML += `<option value="${m}">${m}</option>`;
                            });
                        }
                    }

                    document.getElementById('consolidadoLabelMes').textContent = `Mês/Ano: ${d.mes_ano_selecionado}`;

                    // Manutenção
                    document.getElementById('kpiPrevRealizadas').textContent = d.manutencao.preventivas_realizadas;
                    document.getElementById('kpiPrevTotal').textContent = d.manutencao.preventivas_agendadas;
                    document.getElementById('kpiAderenciaPrev').textContent = d.manutencao.aderencia_preventiva;
                    document.getElementById('kpiCorrSolicitadas').textContent = d.manutencao.corretivas_solicitadas;
                    document.getElementById('kpiCorrConcluidas').textContent = d.manutencao.corretivas_concluidas;
                    document.getElementById('kpiResolucaoCorr').textContent = d.manutencao.resolucao_corretivas;
                    document.getElementById('kpiTotalInsumos').textContent = d.manutencao.total_pecas_insumos;
                    document.getElementById('kpiTaxaRetrabalho').textContent = d.manutencao.taxa_retrabalho;

                    // Rouparia
                    document.getElementById('kpiPecasEntregues').textContent = d.rouparia.pecas_entregues;
                    document.getElementById('kpiSujoKg').textContent = d.rouparia.enxoval_sujo_kg.toLocaleString('pt-BR', {minimumFractionDigits: 2});
                    document.getElementById('kpiLimpoKg').textContent = d.rouparia.enxoval_limpo_kg.toLocaleString('pt-BR', {minimumFractionDigits: 2});
                    document.getElementById('kpiDevolucoes').textContent = d.rouparia.devolucoes_inconformidade;
                    document.getElementById('kpiNivelAtendimento').textContent = d.rouparia.nivel_atendimento;
                    document.getElementById('kpiPontualidade').textContent = d.rouparia.pontualidade;
                    document.getElementById('kpiCustoKg').textContent = d.rouparia.custo_kg;
                    document.getElementById('kpiCustoTotal').textContent = d.rouparia.custo_total_lavanderia;

                } catch (e) {
                    console.error('Erro ao carregar consolidado:', e);
                }
            }

            function baixarConsolidadoExcel() {
                const mes = document.getElementById('consolidadoFiltroMes').value || 'Todos';
                window.open(`/api/exportar/consolidado/excel?mes_ano=${encodeURIComponent(mes)}`, '_blank');
            }

            async function enviarPlanilhaExcel(event) {
                const file = event.target.files[0];
                if (!file) return;

                const formData = new FormData();
                formData.append("arquivo", file);

                mostrarToast("Lendo e importando planilha... aguarde!", "aviso");
                try {
                    const res = await fetch("/api/importar/excel", {
                        method: "POST",
                        body: formData
                    });
                    const d = await res.json();
                    if (res.ok) {
                        mostrarToast(d.mensagem);
                        carregarConsolidado();
                        carregarRelatorios();
                        carregarHistorico();
                        carregarDropdownsFormulario();
                    } else {
                        mostrarToast(d.detail || "Erro na importação", "erro");
                    }
                } catch(err) {
                    mostrarToast("Falha de comunicação ao importar arquivo", "erro");
                } finally {
                    event.target.value = "";
                }
            }

            // Carregar Relatórios & Gráficos (Imagem 4)
            async function carregarRelatorios() {
                const mes = document.getElementById('relatorioFiltroMes').value;
                try {
                    const res = await fetch(`/api/relatorios/manutencao?mes_ano=${encodeURIComponent(mes)}`);
                    const d = await res.json();

                    // 1. Tabela e Gráfico de Status
                    const tbStatus = document.getElementById('tabelaRelatorioStatus');
                    tbStatus.innerHTML = '';
                    const statusLabels = Object.keys(d.status);
                    const statusData = Object.values(d.status);

                    statusLabels.forEach(s => {
                        tbStatus.innerHTML += `
                            <tr>
                                <td class="p-2.5 font-medium text-slate-700">${s}</td>
                                <td class="p-2.5 text-right font-extrabold text-blue-900">${d.status[s]}</td>
                            </tr>
                        `;
                    });
                    tbStatus.innerHTML += `
                        <tr class="bg-slate-50 font-bold border-t">
                            <td class="p-2.5">Total geral</td>
                            <td class="p-2.5 text-right font-black text-slate-900">${d.total_geral}</td>
                        </tr>
                    `;

                    // Renderizar Chart Status
                    if (chartStatusInstance) chartStatusInstance.destroy();
                    const ctxStatus = document.getElementById('chartStatus').getContext('2d');
                    chartStatusInstance = new Chart(ctxStatus, {
                        type: 'doughnut',
                        data: {
                            labels: statusLabels,
                            datasets: [{
                                data: statusData,
                                backgroundColor: ['#0284c7', '#10b981', '#f59e0b', '#1e3a8a']
                            }]
                        },
                        options: {
                            responsive: true,
                            maintainAspectRatio: false,
                            plugins: { legend: { position: 'bottom', labels: { boxWidth: 10, font: { size: 10 } } } }
                        }
                    });

                    // 2. Tabela e Gráfico de Categorias
                    const tbCat = document.getElementById('tabelaRelatorioCategoria');
                    tbCat.innerHTML = '';
                    const catLabels = Object.keys(d.categorias);
                    const catData = Object.values(d.categorias);

                    catLabels.forEach(c => {
                        tbCat.innerHTML += `
                            <tr>
                                <td class="p-2.5 font-medium text-slate-700">${c}</td>
                                <td class="p-2.5 text-right font-extrabold text-blue-900">${d.categorias[c]}</td>
                            </tr>
                        `;
                    });

                    if (chartCategoriaInstance) chartCategoriaInstance.destroy();
                    const ctxCat = document.getElementById('chartCategoria').getContext('2d');
                    chartCategoriaInstance = new Chart(ctxCat, {
                        type: 'bar',
                        data: {
                            labels: catLabels,
                            datasets: [{
                                label: 'Manutenções',
                                data: catData,
                                backgroundColor: '#1d4ed8',
                                borderRadius: 4
                            }]
                        },
                        options: {
                            responsive: true,
                            maintainAspectRatio: false,
                            plugins: { legend: { display: false } },
                            scales: { x: { display: false } }
                        }
                    });

                    // 3. Tabela Ranking Setores
                    const tbSet = document.getElementById('tabelaRelatorioSetores');
                    tbSet.innerHTML = '';
                    const maxSet = d.setores_ranking.length > 0 ? d.setores_ranking[0].quantidade : 1;
                    d.setores_ranking.forEach(item => {
                        const pct = Math.round((item.quantidade / maxSet) * 100);
                        tbSet.innerHTML += `
                            <tr>
                                <td class="p-2.5 font-semibold text-slate-900">${item.setor}</td>
                                <td class="p-2.5 text-right font-black text-blue-900">${item.quantidade}</td>
                                <td class="p-2.5">
                                    <div class="w-full bg-slate-100 rounded-full h-2 overflow-hidden max-w-xs">
                                        <div class="bg-blue-600 h-2 rounded-full" style="width: ${pct}%"></div>
                                    </div>
                                </td>
                            </tr>
                        `;
                    });

                    // 4. Matriz Cruzada (Setor x Item)
                    const tbMatriz = document.getElementById('tabelaMatrizCruzada');
                    tbMatriz.innerHTML = '';
                    let headerHtml = `<thead class="bg-slate-900 text-white font-bold text-[10px]"><tr><th class="p-2 border border-slate-700 sticky left-0 bg-slate-900">SETOR / LOCAL</th>`;
                    d.itens_unicos.forEach(it => {
                        headerHtml += `<th class="p-2 border border-slate-700 text-center whitespace-nowrap">${it}</th>`;
                    });
                    headerHtml += `</tr></thead>`;

                    let bodyHtml = `<tbody class="divide-y divide-slate-200">`;
                    const setoresMatriz = Object.keys(d.matriz).sort();
                    setoresMatriz.forEach(st => {
                        bodyHtml += `<tr><td class="p-2 border border-slate-200 font-bold text-slate-800 bg-slate-50 sticky left-0 whitespace-nowrap">${st}</td>`;
                        d.itens_unicos.forEach(it => {
                            const val = d.matriz[st][it] || '';
                            const cor = val ? 'bg-blue-50 font-extrabold text-blue-900' : 'text-slate-300';
                            bodyHtml += `<td class="p-2 border border-slate-200 text-center ${cor}">${val}</td>`;
                        });
                        bodyHtml += `</tr>`;
                    });
                    bodyHtml += `</tbody>`;
                    tbMatriz.innerHTML = headerHtml + bodyHtml;

                } catch (e) {
                    console.error('Erro ao carregar relatórios:', e);
                }
            }

            // ===============================================================
            // SUB-ABAS DE RELATÓRIOS & DOCUMENTOS OFICIAIS DE ROUPARIA
            // ===============================================================
            let subAbaRelatorioAtiva = 'manutencao';

            function alternarSubRelatorio(sub) {
                subAbaRelatorioAtiva = sub;
                const subs = ['manutencao', 'rouparia_diario', 'rouparia_enxoval', 'rouparia_peso'];
                const btns = {
                    'manutencao': document.getElementById('btnSubRelManut'),
                    'rouparia_diario': document.getElementById('btnSubRelDiario'),
                    'rouparia_enxoval': document.getElementById('btnSubRelEnxoval'),
                    'rouparia_peso': document.getElementById('btnSubRelPeso')
                };
                const views = {
                    'manutencao': document.getElementById('subview-rel-manut'),
                    'rouparia_diario': document.getElementById('subview-rel-diario'),
                    'rouparia_enxoval': document.getElementById('subview-rel-enxoval'),
                    'rouparia_peso': document.getElementById('subview-rel-peso')
                };

                subs.forEach(s => {
                    if (s === sub) {
                        if (views[s]) views[s].classList.remove('hidden');
                        if (btns[s]) btns[s].className = 'px-3.5 py-2 text-xs font-bold rounded-xl bg-blue-900 text-white shadow-sm transition flex items-center gap-1.5';
                    } else {
                        if (views[s]) views[s].classList.add('hidden');
                        if (btns[s]) btns[s].className = 'px-3.5 py-2 text-xs font-bold rounded-xl bg-slate-100 text-slate-700 hover:bg-slate-200 transition flex items-center gap-1.5';
                    }
                });

                const wMes = document.getElementById('wrapperFiltroMesManut');
                const wData = document.getElementById('wrapperFiltroDataRouparia');
                if (sub === 'manutencao') {
                    if (wMes) wMes.classList.remove('hidden');
                    if (wData) wData.classList.add('hidden');
                    carregarRelatorios();
                } else {
                    if (wMes) wMes.classList.add('hidden');
                    if (wData) wData.classList.remove('hidden');
                    carregarDocumentosRouparia();
                }
            }

            // Carregar Documentos Oficiais da Rouparia (Imagens 1, 2 e 3)
            async function carregarDocumentosRouparia() {
                const selData = document.getElementById('relatorioFiltroDataRoupa');
                const dataVal = (selData && selData.value) ? selData.value : '01/08/2026';
                try {
                    const res = await fetch(`/api/relatorios/rouparia/documentos?data=${encodeURIComponent(dataVal)}`);
                    if (!res.ok) throw new Error('Falha ao carregar documentos');
                    const d = await res.json();

                    // Preencher o seletor de datas se necessário
                    if (selData && (selData.options.length <= 1 || !Array.from(selData.options).some(o => o.value === d.data_selecionada))) {
                        selData.innerHTML = '';
                        d.datas_disponiveis.forEach(dt => {
                            const opt = document.createElement('option');
                            opt.value = dt;
                            opt.textContent = dt;
                            if (dt === d.data_selecionada) opt.selected = true;
                            selData.appendChild(opt);
                        });
                    }

                    // Atualizar referências visuais de data
                    const elD1 = document.getElementById('diarioDataRef');
                    if (elD1) elD1.textContent = d.data_selecionada;
                    const elD2 = document.getElementById('enxovalDataRef');
                    if (elD2) elD2.textContent = 'DATA: ' + d.data_selecionada;
                    const elD3 = document.getElementById('pesoSujoDataRef');
                    if (elD3) elD3.textContent = 'DATA: ' + d.data_selecionada;

                    // 1. Renderizar Relatório Diário (Imagem 1)
                    renderizarRelatorioDiario(d.relatorio_diario, d.itens_nomes);

                    // 2. Renderizar Enxoval Entregue nas Áreas (Imagem 2)
                    renderizarEnxovalEntregue(d.enxoval_entregue, d.itens_nomes);

                    // 3. Renderizar Peso Sujo (Imagem 3)
                    renderizarPesoSujo(d.peso_sujo);

                } catch (e) {
                    console.error('Erro ao carregar documentos de rouparia:', e);
                }
            }

            // Renderizar Relatório Diário (Imagem 1)
            function renderizarRelatorioDiario(rd, itensNomes) {
                // 1. Quantitativo Recebido
                const tbRec = document.getElementById('tbodyDiarioRecebidos');
                if (tbRec) {
                    tbRec.innerHTML = '';
                    let totRec = 0;
                    itensNomes.forEach(it => {
                        const qtd = rd.recebidos[it] || 0;
                        totRec += qtd;
                        tbRec.innerHTML += `
                            <tr class="hover:bg-amber-50/50">
                                <td class="p-2 border-r border-amber-100 font-medium text-slate-800">${it}</td>
                                <td class="p-2 text-center font-bold ${qtd > 0 ? 'text-amber-800 text-sm' : 'text-red-500'}">${qtd}</td>
                            </tr>
                        `;
                    });
                    const bRec = document.getElementById('totRecebidosPecas');
                    if (bRec) bRec.textContent = `${totRec} peças recebidas`;
                }

                // 7. Estoque
                const tbEst = document.getElementById('tbodyDiarioEstoque');
                if (tbEst) {
                    tbEst.innerHTML = '';
                    itensNomes.forEach(it => {
                        const qtd = rd.estoque[it] || 0;
                        tbEst.innerHTML += `
                            <tr class="hover:bg-emerald-50/50">
                                <td class="p-2 border-r border-emerald-100 font-medium text-slate-800">${it}</td>
                                <td class="p-2 text-center font-bold ${qtd < 0 ? 'text-rose-600' : qtd > 0 ? 'text-emerald-800 font-black' : 'text-slate-400'}">${qtd}</td>
                            </tr>
                        `;
                    });
                }

                // 2. Intercorrências
                const tbInt = document.getElementById('tbodyDiarioIntercorrencias');
                if (tbInt) {
                    tbInt.innerHTML = '';
                    const setoresInt = Object.keys(rd.intercorrencias);
                    if (setoresInt.length === 0) {
                        tbInt.innerHTML = `<tr><td colspan="3" class="p-4 text-center text-slate-400 italic">Nenhuma intercorrência registrada nesta data.</td></tr>`;
                    } else {
                        setoresInt.forEach(st => {
                            rd.intercorrencias[st].forEach(oc => {
                                tbInt.innerHTML += `
                                    <tr class="hover:bg-yellow-50/50">
                                        <td class="p-2 border-r border-yellow-100 font-bold text-slate-800">${st}</td>
                                        <td class="p-2 text-slate-700">${oc.item}: <span class="italic text-rose-700">${oc.inconformidade}</span></td>
                                        <td class="p-2 text-center font-bold text-amber-900">${oc.qtd}</td>
                                    </tr>
                                `;
                            });
                        });
                    }
                }

                // 3. Distribuição por Setor
                const tbDist = document.getElementById('tbodyDiarioDistribuicao');
                if (tbDist) {
                    tbDist.innerHTML = '';
                    let totDist = 0;
                    if (rd.distribuicao.length === 0) {
                        tbDist.innerHTML = `<tr><td colspan="3" class="p-4 text-center text-slate-400 italic">Nenhuma entrega para setores nesta data.</td></tr>`;
                    } else {
                        rd.distribuicao.forEach(ds => {
                            totDist += ds.qtd;
                            tbDist.innerHTML += `
                                <tr class="hover:bg-cyan-50/50">
                                    <td class="p-2 border-r border-cyan-100 font-bold text-slate-800">${ds.setor}</td>
                                    <td class="p-2 border-r border-cyan-100 text-slate-700">${ds.item}</td>
                                    <td class="p-2 text-center font-bold text-cyan-900">${ds.qtd}</td>
                                </tr>
                            `;
                        });
                    }
                    const bDist = document.getElementById('totDistribuidasPecas');
                    if (bDist) bDist.textContent = `${totDist} peças enviadas`;
                }

                // 4. Recolhimento Interno
                const tbRecInt = document.getElementById('tbodyDiarioRecolhimentoInterno');
                if (tbRecInt) {
                    tbRecInt.innerHTML = '';
                    if (rd.recolhimento_interno.length === 0) {
                        tbRecInt.innerHTML = `<tr><td colspan="2" class="p-4 text-center text-slate-400 italic">Nenhum recolhimento interno registrado nesta data.</td></tr>`;
                    } else {
                        rd.recolhimento_interno.forEach(ri => {
                            tbRecInt.innerHTML += `
                                <tr class="hover:bg-slate-50">
                                    <td class="p-2 border-r border-slate-200 font-medium text-slate-800">${ri.setor}</td>
                                    <td class="p-2 text-center font-bold text-blue-900">${ri.peso_kg.toFixed(2)} kg</td>
                                </tr>
                            `;
                        });
                    }
                }

                // 5 e 6. Devolução e Recolhimento Sujo
                const elDev = document.getElementById('diarioDevolucaoKg');
                if (elDev) elDev.textContent = `${rd.devolucao_inconf_kg.toFixed(2)} kg`;
                const elSujo = document.getElementById('diarioRecolhimentoSujoKg');
                if (elSujo) elSujo.textContent = `${rd.recolhimento_sujo_kg.toFixed(2)} kg`;

                // 8. Ocorrências
                const elOcor = document.getElementById('diarioOcorrenciasTurno');
                if (elOcor) {
                    if (rd.ocorrencias.length > 0) {
                        elOcor.innerHTML = rd.ocorrencias.map(o => `• ${o}`).join('<br>');
                    } else {
                        elOcor.textContent = 'Nenhuma ocorrência registrada para este plantão.';
                    }
                }
            }

            // Renderizar Enxoval Entregue nas Áreas (Imagem 2) - 4 Quadros Lado a Lado
            function renderizarEnxovalEntregue(entregas, itensNomes) {
                const container = document.getElementById('containerQuadrantesEnxoval');
                if (!container) return;
                container.innerHTML = '';

                // Cores idênticas aos 4 quadrantes da planilha oficial física
                const cores = [
                    { border: 'border-amber-500', headerBg: 'bg-amber-500', headerText: 'text-amber-950', theadBg: 'bg-amber-400', theadText: 'text-amber-950' },
                    { border: 'border-yellow-400', headerBg: 'bg-yellow-400', headerText: 'text-yellow-950', theadBg: 'bg-yellow-300', theadText: 'text-yellow-950' },
                    { border: 'border-cyan-500', headerBg: 'bg-cyan-500', headerText: 'text-cyan-950', theadBg: 'bg-cyan-400', theadText: 'text-cyan-950' },
                    { border: 'border-emerald-500', headerBg: 'bg-emerald-500', headerText: 'text-emerald-950', theadBg: 'bg-emerald-400', theadText: 'text-emerald-950' }
                ];

                const listaSetores = [...entregas];
                const nomesPadrao = ['Triagem', 'Sala de Observação', 'Medicação', 'CME'];
                while (listaSetores.length < 4) {
                    const idx = listaSetores.length;
                    listaSetores.push({
                        setor: nomesPadrao[idx] || `Setor ${idx + 1}`,
                        horario: '',
                        responsavel_recebimento: '',
                        itens: {}
                    });
                }

                listaSetores.slice(0, 4).forEach((sec, idx) => {
                    const c = cores[idx % cores.length];
                    let rowsHtml = '';
                    itensNomes.forEach(it => {
                        const qtd = sec.itens[it] || 0;
                        rowsHtml += `
                            <tr class="border-b border-slate-300">
                                <td class="px-1 py-0.5 border-r border-slate-300 text-slate-900 text-[8.5px] font-medium truncate max-w-[125px]" title="${it}">${it}</td>
                                <td class="px-1 py-0.5 text-center font-bold text-[8.5px] ${qtd > 0 ? 'text-blue-900 bg-blue-50/70 font-black' : 'text-red-500'}">${qtd}</td>
                            </tr>
                        `;
                    });

                    container.innerHTML += `
                        <div class="border-2 ${c.border} rounded-lg overflow-hidden shadow-sm flex flex-col justify-between bg-white text-[9px] w-full">
                            <div>
                                <div class="${c.headerBg} ${c.headerText} px-1.5 py-1 font-black text-[9px] border-b border-slate-300">
                                    <div class="truncate uppercase font-black">SETOR: ${sec.setor.toUpperCase()}</div>
                                    <div class="flex justify-between items-center text-[8.5px] mt-0.5 font-bold">
                                        <span>HORÁRIO:</span>
                                        <span>${sec.horario || '__:__'}</span>
                                    </div>
                                </div>
                                <table class="w-full text-left border-collapse">
                                    <thead class="${c.theadBg} ${c.theadText} font-black text-[8px] border-b border-slate-300">
                                        <tr>
                                            <th class="px-1 py-0.5 border-r border-slate-300 uppercase tracking-tight">DESCRIMINAÇÃO</th>
                                            <th class="px-1 py-0.5 text-center w-7 uppercase">QTD</th>
                                        </tr>
                                    </thead>
                                    <tbody>
                                        ${rowsHtml}
                                    </tbody>
                                </table>
                            </div>
                            <div class="px-1 py-1 bg-slate-50 border-t border-slate-200 text-[8px] font-bold text-slate-700">
                                <span class="block text-[7px] uppercase text-slate-500">RESPONSÁVEL PELO RECEBIMENTO:</span>
                                <span class="italic text-slate-900 block truncate mt-0.5">${sec.responsavel_recebimento || '_____________________'}</span>
                            </div>
                        </div>
                    `;
                });
            }

            // Renderizar Peso do Enxoval Sujo (Imagem 3)
            function renderizarPesoSujo(ps) {
                const tb = document.getElementById('tbodyPesoSujoRegistros');
                if (!tb) return;
                tb.innerHTML = '';

                const regs = ps.registros || [];
                regs.forEach(r => {
                    tb.innerHTML += `
                        <tr class="hover:bg-blue-50/50">
                            <td class="p-2 text-center font-bold text-slate-700 border-r border-slate-200">${r.ord}</td>
                            <td class="p-2 font-medium text-slate-800 border-r border-slate-200">${r.data}</td>
                            <td class="p-2 text-center font-bold text-slate-900 border-r border-slate-200">${r.qtd_hamper || '-'}</td>
                            <td class="p-2 text-center font-bold text-emerald-800 border-r border-slate-200">${r.peso_kg > 0 ? r.peso_kg.toFixed(2) + ' kg' : '-'}</td>
                            <td class="p-2 text-slate-700 border-r border-slate-200 italic">${r.resp_entrega}</td>
                            <td class="p-2 text-slate-700 italic">${r.resp_recebimento}</td>
                        </tr>
                    `;
                });

                const totalLinhas = Math.max(20, regs.length);
                for (let i = regs.length + 1; i <= totalLinhas; i++) {
                    tb.innerHTML += `
                        <tr class="h-6">
                            <td class="p-1.5 text-center text-slate-400 border-r border-slate-200">${i}</td>
                            <td class="border-r border-slate-200"></td>
                            <td class="border-r border-slate-200"></td>
                            <td class="border-r border-slate-200"></td>
                            <td class="border-r border-slate-200"></td>
                            <td></td>
                        </tr>
                    `;
                }

                const elTotH = document.getElementById('pesoSujoTotalHampers');
                if (elTotH) elTotH.textContent = ps.total_hampers;
                const elTotKg = document.getElementById('pesoSujoTotalKg');
                if (elTotKg) elTotKg.textContent = `${ps.total_peso_kg.toFixed(2)} kg`;
            }

            // ===============================================================
            // MOTOR DE IMPRESSÃO OFICIAL (HTML / A4 COM CSS DEDICADO)
            // ===============================================================
            function imprimirDocumentoHtml(conteudoHtml, titulo, orientacao = 'portrait') {
                const frame = document.createElement('iframe');
                frame.style.position = 'fixed';
                frame.style.right = '0';
                frame.style.bottom = '0';
                frame.style.width = '0';
                frame.style.height = '0';
                frame.style.border = '0';
                document.body.appendChild(frame);

                const pageSize = orientacao === 'landscape' ? 'A4 landscape' : 'A4 portrait';
                const doc = frame.contentWindow.document;
                doc.open();
                doc.write('<!DOCTYPE html><html lang="pt-BR"><head><meta charset="utf-8">' +
                    '<title>' + (titulo || 'Documento Oficial - SEMUS / PS ANIL') + '</title>' +
                    '<link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/tailwindcss@2.2.19/dist/tailwind.min.css">' +
                    '<style>' +
                    '@page { size: ' + pageSize + '; margin: 4mm 5mm; } ' +
                    'body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Arial, sans-serif; background: #fff; color: #0f172a; font-size: 9px; line-height: 1.15; } ' +
                    'table { width: 100%; border-collapse: collapse; } ' +
                    'th, td { border: 1px solid #cbd5e1; padding: 1px 3px; } ' +
                    '.no-print { display: none !important; } ' +
                    '#containerQuadrantesEnxoval { display: grid !important; grid-template-columns: repeat(4, 1fr) !important; gap: 6px !important; width: 100% !important; } ' +
                    '@media print { body { print-color-adjust: exact; -webkit-print-color-adjust: exact; } }' +
                    '</style>' +
                    '</head><body class="p-1">' + conteudoHtml + '</body></html>');
                doc.close();

                setTimeout(() => {
                    frame.contentWindow.focus();
                    frame.contentWindow.print();
                    setTimeout(() => { frame.remove(); }, 2000);
                }, 500);
            }

            // Imprimir o Relatório ou Documento Ativo na tela
            function imprimirRelatorioAtivo() {
                if (subAbaRelatorioAtiva === 'manutencao') {
                    const el = document.getElementById('subview-rel-manut');
                    if (el) {
                        const html = `
                            <div class="border-b-2 border-slate-900 pb-3 mb-4 flex justify-between items-center">
                                <div>
                                    <span class="text-xs font-black uppercase text-blue-900 block">Prefeitura Municipal de São Luís • SEMUS</span>
                                    <h1 class="text-lg font-black uppercase">Relatório de Manutenções e Insumos Hospitalares</h1>
                                    <p class="text-xs text-slate-500">Pronto Socorro do Anil (PS Anil) • Relatório Sintético de Gestão</p>
                                </div>
                                <div class="text-right text-xs">
                                    <span class="font-bold">Mês de Referência: ${document.getElementById('relatorioFiltroMes').value}</span><br>
                                    <span class="text-slate-400">Gerado em: ${new Date().toLocaleDateString('pt-BR')} às ${obterHoraAtual()}</span>
                                </div>
                            </div>
                            ${el.innerHTML}
                        `;
                        imprimirDocumentoHtml(html, 'Relatorio-Manutencoes-PS-ANIL');
                    }
                } else if (subAbaRelatorioAtiva === 'rouparia_diario') {
                    const el = document.getElementById('printableRelatorioDiario');
                    if (el) imprimirDocumentoHtml(el.innerHTML, 'Relatorio-Diario-Rouparia-PS-ANIL');
                } else if (subAbaRelatorioAtiva === 'rouparia_enxoval') {
                    const el = document.getElementById('printableEnxovalEntregue');
                    if (el) imprimirDocumentoHtml(el.innerHTML, 'Planilha-Enxoval-Entregue-PS-ANIL', 'landscape');
                } else if (subAbaRelatorioAtiva === 'rouparia_peso') {
                    const el = document.getElementById('printablePesoSujo');
                    if (el) imprimirDocumentoHtml(el.innerHTML, 'Planilha-Peso-Enxoval-Sujo-PS-ANIL');
                }
            }

            // Imprimir Ficha de Manutenção (Aba Detalhes do BD)
            function imprimirFichaManutencao() {
                const codigo = document.getElementById('viewManutCodigo').textContent;
                const status = document.getElementById('viewManutStatusBadge').textContent;
                const dataHora = document.getElementById('viewManutDataHora').textContent;
                const turno = document.getElementById('viewManutTurno').textContent;
                const tipo = document.getElementById('viewManutTipo').textContent;
                const categoria = document.getElementById('viewManutCategoria').textContent;
                const item = document.getElementById('viewManutItem').textContent;
                const setor = document.getElementById('viewManutSetor').textContent;
                const quantidade = document.getElementById('viewManutQtd').textContent;
                const chamado = document.getElementById('viewManutChamado').textContent;
                const conclusao = document.getElementById('viewManutConclusao').textContent;
                const respEntrega = document.getElementById('viewManutRespEntrega').textContent;
                const respRecebimento = document.getElementById('viewManutRespRecebimento').textContent;
                const descricao = document.getElementById('viewManutDescricao').textContent;
                const observacoes = document.getElementById('viewManutObservacoes').textContent;

                const html = `
                    <div class="p-6 space-y-4 max-w-3xl mx-auto">
                        <!-- Cabeçalho Oficial -->
                        <div class="border-b-2 border-slate-900 pb-3 flex justify-between items-center">
                            <div>
                                <span class="text-xs font-black uppercase text-blue-900 block">Prefeitura de São Luís • SEMUS</span>
                                <h1 class="text-xl font-black uppercase text-slate-900">Ficha Técnica de Atendimento de Manutenção</h1>
                                <p class="text-xs text-slate-500 font-bold">Pronto Socorro do Anil (PS Anil) • Superintendência de Apoio e Manutenção</p>
                            </div>
                            <div class="text-right">
                                <span class="text-base font-black text-blue-900 block">${codigo}</span>
                                <span class="inline-block px-2.5 py-0.5 rounded text-xs font-bold border border-slate-300 bg-slate-100">Status: ${status}</span>
                            </div>
                        </div>

                        <!-- Tabela de Dados do Chamado -->
                        <table class="w-full text-xs">
                            <tr>
                                <td class="bg-slate-100 font-bold w-1/4">Data & Horário:</td>
                                <td class="w-1/4">${dataHora}</td>
                                <td class="bg-slate-100 font-bold w-1/4">Turno:</td>
                                <td class="w-1/4">${turno}</td>
                            </tr>
                            <tr>
                                <td class="bg-slate-100 font-bold">Categoria do Serviço:</td>
                                <td>${categoria}</td>
                                <td class="bg-slate-100 font-bold">Tipo de Manutenção:</td>
                                <td>${tipo}</td>
                            </tr>
                            <tr>
                                <td class="bg-slate-100 font-bold">Item / Sub-Categoria:</td>
                                <td>${item}</td>
                                <td class="bg-slate-100 font-bold">Setor / Local:</td>
                                <td class="font-bold">${setor}</td>
                            </tr>
                            <tr>
                                <td class="bg-slate-100 font-bold">Nº Pedido / Chamado:</td>
                                <td class="font-bold text-blue-900">${chamado || '-'}</td>
                                <td class="bg-slate-100 font-bold">Quantidade / Volume:</td>
                                <td>${quantidade}</td>
                            </tr>
                            <tr>
                                <td class="bg-slate-100 font-bold">Data Conclusão / Retorno:</td>
                                <td>${conclusao || '-'}</td>
                                <td class="bg-slate-100 font-bold">Retrabalho:</td>
                                <td>${status.includes('Retrabalho') ? 'SIM' : 'NÃO'}</td>
                            </tr>
                            <tr>
                                <td class="bg-slate-100 font-bold">Resp. Entrega / Vistoria:</td>
                                <td>${respEntrega || '-'}</td>
                                <td class="bg-slate-100 font-bold">Resp. Recebimento:</td>
                                <td>${respRecebimento || '-'}</td>
                            </tr>
                        </table>

                        <!-- Descrição do Serviço -->
                        <div class="border border-slate-300 p-3 rounded-lg bg-slate-50">
                            <span class="block text-[10px] font-bold uppercase text-slate-500 mb-1">Descrição / Especificação da Demanda:</span>
                            <p class="whitespace-pre-line text-xs font-medium text-slate-800 leading-relaxed">${descricao || 'Sem descrição cadastrada.'}</p>
                        </div>

                        <!-- Observações Gerais e Auditoria -->
                        <div class="border border-blue-200 p-3 rounded-lg bg-blue-50/50">
                            <span class="block text-[10px] font-bold uppercase text-blue-900 mb-1">Observações Gerais & Histórico de Edição:</span>
                            <div class="whitespace-pre-line text-xs text-slate-700 leading-relaxed">${observacoes || 'Nenhuma observação registrada.'}</div>
                        </div>

                        <!-- Assinaturas -->
                        <div class="pt-10 grid grid-cols-2 gap-8 text-center text-xs">
                            <div>
                                <div class="border-t border-slate-900 pt-2 font-bold">${respRecebimento || 'Responsável pelo Setor / Solicitante'}</div>
                                <span class="text-[10px] text-slate-500 block">Assinatura / Matrícula</span>
                            </div>
                            <div>
                                <div class="border-t border-slate-900 pt-2 font-bold">${respEntrega || 'Técnico Responsável / Executante'}</div>
                                <span class="text-[10px] text-slate-500 block">Assinatura / Matrícula</span>
                            </div>
                        </div>
                    </div>
                `;
                imprimirDocumentoHtml(html, `Ficha-${codigo}-PS-ANIL`);
            }

            // Imprimir Ficha de Rouparia (Aba Detalhes do BD)
            function imprimirFichaRouparia() {
                const codigo = document.getElementById('viewRoupaCodigo').textContent;
                const dataHora = document.getElementById('viewRoupaDataHora').textContent;
                const turno = document.getElementById('viewRoupaTurno').textContent;
                const tipo = document.getElementById('viewRoupaTipoMov').textContent;
                const item = document.getElementById('viewRoupaItem').textContent;
                const setor = document.getElementById('viewRoupaSetor').textContent;
                const qtd = document.getElementById('viewRoupaQtd').textContent;
                const peso = document.getElementById('viewRoupaPeso').textContent;
                const pontual = document.getElementById('viewRoupaPontual').textContent;
                const respUnidade = document.getElementById('viewRoupaRespUnidade').textContent;
                const respTerc = document.getElementById('viewRoupaRespTerceirizada').textContent;
                const inconf = document.getElementById('viewRoupaInconformidade').textContent;

                const html = `
                    <div class="p-6 space-y-4 max-w-3xl mx-auto">
                        <div class="border-b-2 border-slate-900 pb-3 flex justify-between items-center">
                            <div>
                                <span class="text-xs font-black uppercase text-blue-900 block">Prefeitura de São Luís • SEMUS</span>
                                <h1 class="text-xl font-black uppercase text-slate-900">Comprovante de Movimentação de Rouparia</h1>
                                <p class="text-xs text-slate-500 font-bold">Pronto Socorro do Anil (PS Anil) • Gestão de Enxoval Hospitalar</p>
                            </div>
                            <div class="text-right">
                                <span class="text-base font-black text-indigo-900 block">${codigo}</span>
                                <span class="inline-block px-2.5 py-0.5 rounded text-xs font-bold border border-slate-300 bg-slate-100">${tipo}</span>
                            </div>
                        </div>

                        <table class="w-full text-xs">
                            <tr>
                                <td class="bg-slate-100 font-bold w-1/4">Data & Horário:</td>
                                <td class="w-1/4">${dataHora}</td>
                                <td class="bg-slate-100 font-bold w-1/4">Turno:</td>
                                <td class="w-1/4">${turno}</td>
                            </tr>
                            <tr>
                                <td class="bg-slate-100 font-bold">Item do Enxoval:</td>
                                <td class="font-bold text-slate-900">${item}</td>
                                <td class="bg-slate-100 font-bold">Setor / Local:</td>
                                <td class="font-bold">${setor}</td>
                            </tr>
                            <tr>
                                <td class="bg-slate-100 font-bold">Quantidade de Peças:</td>
                                <td class="font-bold text-blue-900">${qtd}</td>
                                <td class="bg-slate-100 font-bold">Peso Registrado:</td>
                                <td>${peso}</td>
                            </tr>
                            <tr>
                                <td class="bg-slate-100 font-bold">Pontualidade (07h-19h):</td>
                                <td>${pontual}</td>
                                <td class="bg-slate-100 font-bold">Tipo Movimentação:</td>
                                <td>${tipo}</td>
                            </tr>
                            <tr>
                                <td class="bg-slate-100 font-bold">Resp. Unidade (SEMUS):</td>
                                <td>${respUnidade || '-'}</td>
                                <td class="bg-slate-100 font-bold">Resp. Terceirizada:</td>
                                <td>${respTerc || 'LAVARE'}</td>
                            </tr>
                        </table>

                        <div class="border border-amber-200 p-3 rounded-lg bg-amber-50/50">
                            <span class="block text-[10px] font-bold uppercase text-amber-900 mb-1">Inconformidade / Ocorrências & Histórico:</span>
                            <div class="whitespace-pre-line text-xs text-slate-800 leading-relaxed">${inconf || 'Nenhuma inconformidade ou observação registrada.'}</div>
                        </div>

                        <div class="pt-10 grid grid-cols-2 gap-8 text-center text-xs">
                            <div>
                                <div class="border-t border-slate-900 pt-2 font-bold">${respUnidade || 'Responsável Unidade (SEMUS)'}</div>
                                <span class="text-[10px] text-slate-500 block">Assinatura / Matrícula</span>
                            </div>
                            <div>
                                <div class="border-t border-slate-900 pt-2 font-bold">${respTerc || 'Responsável Empresa Terceirizada (LAVARE)'}</div>
                                <span class="text-[10px] text-slate-500 block">Assinatura / Recebimento</span>
                            </div>
                        </div>
                    </div>
                `;
                imprimirDocumentoHtml(html, `Ficha-${codigo}-Rouparia-PS-ANIL`);
            }

            // Alternar Histórico
            function alternarTabelaHistorico(tipo) {
                const bManut = document.getElementById('btnHistManut');
                const bRoupa = document.getElementById('btnHistRoupa');
                const cManut = document.getElementById('containerHistManut');
                const cRoupa = document.getElementById('containerHistRoupa');
                const pManut = document.getElementById('painelFiltrosManut');
                const pRoupa = document.getElementById('painelFiltrosRoupa');

                if (tipo === 'manutencao') {
                    bManut.className = 'px-4 py-2.5 text-xs font-bold rounded-xl bg-blue-900 text-white shadow-sm transition flex items-center gap-2';
                    bRoupa.className = 'px-4 py-2.5 text-xs font-bold rounded-xl bg-slate-100 text-slate-700 hover:bg-slate-200 transition flex items-center gap-2';
                    cManut.classList.remove('hidden');
                    cRoupa.classList.add('hidden');
                    if (pManut) pManut.classList.remove('hidden');
                    if (pRoupa) pRoupa.classList.add('hidden');
                } else {
                    bRoupa.className = 'px-4 py-2.5 text-xs font-bold rounded-xl bg-blue-900 text-white shadow-sm transition flex items-center gap-2';
                    bManut.className = 'px-4 py-2.5 text-xs font-bold rounded-xl bg-slate-100 text-slate-700 hover:bg-slate-200 transition flex items-center gap-2';
                    cRoupa.classList.remove('hidden');
                    cManut.classList.add('hidden');
                    if (pManut) pManut.classList.add('hidden');
                    if (pRoupa) pRoupa.classList.remove('hidden');
                }
                carregarHistorico();
            }

            // Carregar Histórico
            async function carregarHistorico() {
                const busca = document.getElementById('filtroBuscaGeral') ? document.getElementById('filtroBuscaGeral').value.trim() : '';

                // Filtros Manutenção
                const mMesAno = document.getElementById('filtroManutMesAno') ? document.getElementById('filtroManutMesAno').value : 'Todos';
                const mTurno = document.getElementById('filtroManutTurno') ? document.getElementById('filtroManutTurno').value : 'Todos';
                const mStatus = document.getElementById('filtroManutStatus') ? document.getElementById('filtroManutStatus').value : 'Todos';
                const mTipo = document.getElementById('filtroManutTipo') ? document.getElementById('filtroManutTipo').value : 'Todos';
                const mCategoria = document.getElementById('filtroManutCategoria') ? document.getElementById('filtroManutCategoria').value : 'Todas';
                const mSetor = document.getElementById('filtroManutSetor') ? document.getElementById('filtroManutSetor').value : 'Todos';
                const mRetrabalho = document.getElementById('filtroManutRetrabalho') ? document.getElementById('filtroManutRetrabalho').value : 'Todos';

                // Filtros Rouparia
                const rMesAno = document.getElementById('filtroRoupaMesAno') ? document.getElementById('filtroRoupaMesAno').value : 'Todos';
                const rTipo = document.getElementById('filtroRoupaTipo') ? document.getElementById('filtroRoupaTipo').value : 'Todos';
                const rItem = document.getElementById('filtroRoupaItem') ? document.getElementById('filtroRoupaItem').value : 'Todos';
                const rSetor = document.getElementById('filtroRoupaSetor') ? document.getElementById('filtroRoupaSetor').value : 'Todos';
                const rTurno = document.getElementById('filtroRoupaTurno') ? document.getElementById('filtroRoupaTurno').value : 'Todos';
                
                // Manutenção
                try {
                    const paramsM = new URLSearchParams();
                    if (busca) paramsM.append('busca', busca);
                    if (mMesAno && mMesAno !== 'Todos') paramsM.append('mes_ano', mMesAno);
                    if (mTurno && mTurno !== 'Todos') paramsM.append('turno', mTurno);
                    if (mStatus && mStatus !== 'Todos') paramsM.append('status', mStatus);
                    if (mTipo && mTipo !== 'Todos') paramsM.append('tipo', mTipo);
                    if (mCategoria && mCategoria !== 'Todas' && mCategoria !== 'Todos') paramsM.append('categoria', mCategoria);
                    if (mSetor && mSetor !== 'Todos') paramsM.append('setor', mSetor);
                    if (mRetrabalho && mRetrabalho !== 'Todos') paramsM.append('retrabalho', mRetrabalho);

                    const resM = await fetch(`/api/manutencao?${paramsM.toString()}`);
                    if (!resM.ok) throw new Error('Falha ao obter registros de manutenção');
                    const dadosM = await resM.json();
                    const tbM = document.getElementById('tabelaManutCorpo');
                    
                    if (Array.isArray(dadosM)) {
                        document.getElementById('contadorManut').textContent = `${dadosM.length} registro(s)`;
                        let htmlM = '';
                        if (dadosM.length === 0) {
                            htmlM = `<tr><td colspan="10" class="p-4 text-center text-slate-400 italic">Nenhum registro de manutenção encontrado para os filtros selecionados.</td></tr>`;
                        } else {
                            dadosM.forEach(r => {
                                let badgeCor = r.status === 'Concluído' ? 'bg-emerald-100 text-emerald-800' : r.status === 'Ok' ? 'bg-sky-100 text-sky-800' : 'bg-amber-100 text-amber-800';
                                let badgeRetrabalho = r.retrabalho ? `<span class="ml-1.5 inline-flex items-center px-1.5 py-0.5 rounded text-[10px] font-extrabold bg-amber-100 text-amber-900 border border-amber-300" title="Identificado como Retrabalho">🔄 Retrabalho</span>` : '';
                                let badgeEditado = (r.observacoes && typeof r.observacoes === 'string' && r.observacoes.includes('[Editado em')) ? `<span class="ml-1.5 inline-flex items-center px-1.5 py-0.5 rounded text-[10px] font-bold bg-indigo-50 text-indigo-700 border border-indigo-200" title="Item com alteração registrada">✏️ Editado</span>` : '';
                                htmlM += `
                                    <tr class="hover:bg-slate-50 transition">
                                        <td class="p-3 text-center font-bold text-slate-500">${r.codigo_id || r.id}</td>
                                        <td class="p-3 whitespace-nowrap">${r.data || '-'}</td>
                                        <td class="p-3">${r.turno || '-'}</td>
                                        <td class="p-3">${r.categoria || '-'}</td>
                                        <td class="p-3 font-semibold">${r.tipo_manutencao || '-'}${badgeRetrabalho}${badgeEditado}</td>
                                        <td class="p-3 font-bold text-slate-900">${r.item || '-'}</td>
                                        <td class="p-3">${r.setor || '-'}</td>
                                        <td class="p-3 font-bold text-center">${r.quantidade || 1}</td>
                                        <td class="p-3 text-center"><span class="px-2 py-0.5 rounded-full text-[10px] font-bold ${badgeCor}">${r.status || '-'}</span></td>
                                        <td class="p-3 text-center">
                                            <div class="flex items-center justify-center gap-1.5">
                                                <button onclick="abrirModalVisualizarManut(${r.id})" class="p-1 text-slate-400 hover:text-indigo-600 hover:bg-indigo-50 rounded transition" title="Visualizar todos os detalhes deste chamado">
                                                    <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M15 12a3 3 0 11-6 0 3 3 0 016 0z"></path><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M2.458 12C3.732 7.943 7.523 5 12 5c4.478 0 8.268 2.943 9.542 7-1.274 4.057-5.064 7-9.542 7-4.477 0-8.268-2.943-9.542-7z"></path></svg>
                                                </button>
                                                <button onclick="abrirModalEditarManut(${r.id})" class="p-1 text-slate-400 hover:text-blue-600 hover:bg-blue-50 rounded transition" title="Editar este chamado">
                                                    <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M11 5H6a2 2 0 00-2 2v11a2 2 0 002 2h11a2 2 0 002-2v-5m-1.414-9.414a2 2 0 112.828 2.828L11.828 15H9v-2.828l8.586-8.586z"></path></svg>
                                                </button>
                                                <button onclick="excluirManutencao(${r.id})" class="p-1 text-slate-400 hover:text-rose-600 hover:bg-rose-50 rounded transition" title="Excluir">
                                                    <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M19 7l-.867 12.142A2 2 0 0116.138 21H7.862a2 2 0 01-1.995-1.858L5 7m5 4v6m4-6v6m1-10V4a1 1 0 00-1-1h-4a1 1 0 00-1 1v3M4 7h16"></path></svg>
                                                </button>
                                            </div>
                                        </td>
                                    </tr>
                                `;
                            });
                        }
                        tbM.innerHTML = htmlM;
                    }
                } catch(e) {
                    console.error('Erro ao carregar histórico de manutenção:', e);
                }

                // Rouparia
                try {
                    const paramsR = new URLSearchParams();
                    if (busca) paramsR.append('busca', busca);
                    if (rMesAno && rMesAno !== 'Todos') paramsR.append('mes_ano', rMesAno);
                    if (rTipo && rTipo !== 'Todos') paramsR.append('tipo_registro', rTipo);
                    if (rItem && rItem !== 'Todos') paramsR.append('item', rItem);
                    if (rSetor && rSetor !== 'Todos') paramsR.append('setor', rSetor);
                    if (rTurno && rTurno !== 'Todos') paramsR.append('turno', rTurno);

                    const resR = await fetch(`/api/rouparia?${paramsR.toString()}`);
                    if (!resR.ok) throw new Error('Falha ao obter registros de rouparia');
                    const dadosR = await resR.json();
                    const tbR = document.getElementById('tabelaRoupaCorpo');

                    if (Array.isArray(dadosR)) {
                        document.getElementById('contadorRoupa').textContent = `${dadosR.length} registro(s)`;
                        let htmlR = '';
                        if (dadosR.length === 0) {
                            htmlR = `<tr><td colspan="10" class="p-4 text-center text-slate-400 italic">Nenhum registro de rouparia encontrado para os filtros selecionados.</td></tr>`;
                        } else {
                            dadosR.forEach(r => {
                                let badgeEditadoR = (r.inconformidade && typeof r.inconformidade === 'string' && r.inconformidade.includes('[Editado em')) ? `<span class="ml-1.5 inline-flex items-center px-1.5 py-0.5 rounded text-[10px] font-bold bg-indigo-50 text-indigo-700 border border-indigo-200" title="Item com alteração registrada">✏️ Editado</span>` : '';
                                const pVal = parseFloat(r.peso_kg) || 0;
                                const pesoFmt = pVal > 0 ? `${pVal.toFixed(2)} kg` : '-';
                                htmlR += `
                                    <tr class="hover:bg-slate-50 transition">
                                        <td class="p-3 text-center font-bold text-slate-500">${r.codigo_id || r.id}</td>
                                        <td class="p-3 whitespace-nowrap">${r.data || '-'}</td>
                                        <td class="p-3">${r.turno || '-'}</td>
                                        <td class="p-3 font-semibold text-blue-900">${r.tipo_registro || '-'}${badgeEditadoR}</td>
                                        <td class="p-3 font-bold text-slate-900">${r.item_enxoval || '-'}</td>
                                        <td class="p-3">${r.setor || '-'}</td>
                                        <td class="p-3 font-bold text-center">${r.qtd_pecas || 0}</td>
                                        <td class="p-3 text-center">${pesoFmt}</td>
                                        <td class="p-3">${r.responsavel_unidade || '-'}</td>
                                        <td class="p-3 text-center">
                                            <div class="flex items-center justify-center gap-1.5">
                                                <button onclick="abrirModalVisualizarRoupa(${r.id})" class="p-1 text-slate-400 hover:text-indigo-600 hover:bg-indigo-50 rounded transition" title="Visualizar todos os detalhes desta movimentação">
                                                    <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M15 12a3 3 0 11-6 0 3 3 0 016 0z"></path><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M2.458 12C3.732 7.943 7.523 5 12 5c4.478 0 8.268 2.943 9.542 7-1.274 4.057-5.064 7-9.542 7-4.477 0-8.268-2.943-9.542-7z"></path></svg>
                                                </button>
                                                <button onclick="abrirModalEditarRoupa(${r.id})" class="p-1 text-slate-400 hover:text-blue-600 hover:bg-blue-50 rounded transition" title="Editar esta movimentação">
                                                    <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M11 5H6a2 2 0 00-2 2v11a2 2 0 002 2h11a2 2 0 002-2v-5m-1.414-9.414a2 2 0 112.828 2.828L11.828 15H9v-2.828l8.586-8.586z"></path></svg>
                                                </button>
                                                <button onclick="excluirRouparia(${r.id})" class="p-1 text-slate-400 hover:text-rose-600 hover:bg-rose-50 rounded transition" title="Excluir">
                                                    <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M19 7l-.867 12.142A2 2 0 0116.138 21H7.862a2 2 0 01-1.995-1.858L5 7m5 4v6m4-6v6m1-10V4a1 1 0 00-1-1h-4a1 1 0 00-1 1v3M4 7h16"></path></svg>
                                                </button>
                                            </div>
                                        </td>
                                    </tr>
                                `;
                            });
                        }
                        tbR.innerHTML = htmlR;
                    }
                } catch(e) {
                    console.error('Erro ao carregar histórico de rouparia:', e);
                }
            }

            function filtrarHistoricoComDebounce() {
                clearTimeout(debounceTimer);
                debounceTimer = setTimeout(carregarHistorico, 300);
            }

            function limparBuscaTexto() {
                const el = document.getElementById('filtroBuscaGeral');
                if (el) el.value = '';
                carregarHistorico();
            }

            function limparTodosFiltros() {
                const elBusca = document.getElementById('filtroBuscaGeral');
                if (elBusca) elBusca.value = '';

                const selIds = [
                    'filtroManutMesAno', 'filtroManutTurno', 'filtroManutStatus', 'filtroManutTipo',
                    'filtroManutCategoria', 'filtroManutSetor', 'filtroManutRetrabalho',
                    'filtroRoupaMesAno', 'filtroRoupaTipo', 'filtroRoupaItem',
                    'filtroRoupaSetor', 'filtroRoupaTurno'
                ];

                selIds.forEach(id => {
                    const sel = document.getElementById(id);
                    if (sel) {
                        sel.selectedIndex = 0;
                    }
                });

                carregarHistorico();
                mostrarToast('Filtros redefinidos com sucesso!', 'sucesso');
            }

            async function excluirManutencao(id) {
                if (!confirm(`Deseja realmente excluir este registro de manutenção?`)) return;
                await fetch(`/api/manutencao/${id}`, { method: 'DELETE' });
                mostrarToast('Registro de manutenção excluído com sucesso!');
                carregarHistorico();
            }

            async function excluirRouparia(id) {
                if (!confirm(`Deseja realmente excluir esta movimentação de rouparia?`)) return;
                await fetch(`/api/rouparia/${id}`, { method: 'DELETE' });
                mostrarToast('Movimentação de rouparia excluída com sucesso!');
                carregarHistorico();
            }

            // Carregar Configurações e Parâmetros
            async function carregarConfiguracoes() {
                try {
                    // Parâmetros
                    const resP = await fetch('/api/parametros');
                    const params = await resP.json();
                    document.getElementById('cfgCustoKg').value = params['custo_kg_lavanderia'] || '4.50';

                    // Itens
                    const resC = await fetch('/api/configuracoes');
                    const itens = await resC.json();

                    const lManut = document.getElementById('listaItensManut');
                    const lEnxoval = document.getElementById('listaItensEnxoval');
                    const lSetores = document.getElementById('listaSetores');
                    const lCategorias = document.getElementById('listaCategorias');

                    lManut.innerHTML = '';
                    lEnxoval.innerHTML = '';
                    lSetores.innerHTML = '';
                    lCategorias.innerHTML = '';

                    let cM = 0, cE = 0, cS = 0, cC = 0;

                    itens.forEach(i => {
                        const li = `
                            <li class="py-2 flex justify-between items-center">
                                <span class="font-medium text-slate-700">${i.nome}</span>
                                <button onclick="removerConfigItem(${i.id})" class="text-slate-300 hover:text-rose-600 transition" title="Excluir">✕</button>
                            </li>
                        `;
                        if (i.tipo === 'item_manutencao') { lManut.innerHTML += li; cM++; }
                        if (i.tipo === 'item_enxoval') { lEnxoval.innerHTML += li; cE++; }
                        if (i.tipo === 'setor') { lSetores.innerHTML += li; cS++; }
                        if (i.tipo === 'categoria_manutencao') { lCategorias.innerHTML += li; cC++; }
                    });

                    document.getElementById('badgeCountItens').textContent = cM;
                    document.getElementById('badgeCountEnxoval').textContent = cE;
                    document.getElementById('badgeCountSetores').textContent = cS;
                    document.getElementById('badgeCountCategorias').textContent = cC;

                } catch(e) {
                    console.error('Erro ao carregar configurações:', e);
                }
            }

            async function salvarParametroCusto() {
                const val = document.getElementById('cfgCustoKg').value;
                await fetch('/api/parametros', {
                    method: 'POST',
                    headers: {'Content-Type': 'application/json'},
                    body: JSON.stringify({ chave: 'custo_kg_lavanderia', valor: val })
                });
                mostrarToast('Custo por quilo salvo com sucesso!');
            }

            async function adicionarConfigItem(tipo, inputId) {
                const input = document.getElementById(inputId);
                const nome = input.value.trim();
                if (!nome) return;

                const res = await fetch('/api/configuracoes', {
                    method: 'POST',
                    headers: {'Content-Type': 'application/json'},
                    body: JSON.stringify({ tipo: tipo, nome: nome })
                });

                if (res.ok) {
                    input.value = '';
                    mostrarToast('Item adicionado com sucesso!');
                    carregarConfiguracoes();
                    carregarDropdownsFormulario();
                } else {
                    const err = await res.json();
                    mostrarToast(err.detail || 'Erro ao adicionar item', 'erro');
                }
            }

            async function removerConfigItem(id) {
                if (!confirm('Deseja desativar este item?')) return;
                await fetch(`/api/configuracoes/${id}`, { method: 'DELETE' });
                mostrarToast('Item desativado!');
                carregarConfiguracoes();
                carregarDropdownsFormulario();
            }

            // ===============================================================
            // AUTENTICAÇÃO E LOGIN
            // ===============================================================
            function verificarAutenticacao() {
                const usuarioStr = localStorage.getItem('usuario_ps_anil');
                const telaLogin = document.getElementById('telaLogin');
                const telaSistema = document.getElementById('telaSistema');
                
                if (usuarioStr) {
                    try {
                        const usuario = JSON.parse(usuarioStr);
                        telaLogin.classList.add('hidden');
                        telaSistema.classList.remove('hidden');
                        document.getElementById('nomeUsuarioLogado').textContent = usuario.nome || usuario.username;
                        carregarDropdownsFormulario();
                        carregarConsolidado();
                    } catch(e) {
                        fazerLogout();
                    }
                } else {
                    telaLogin.classList.remove('hidden');
                    telaSistema.classList.add('hidden');
                }
            }

            function preencherLogin(u, p) {
                document.getElementById('loginUsuario').value = u;
                document.getElementById('loginSenha').value = p;
                const err = document.getElementById('loginErro');
                if (err) err.classList.add('hidden');
                document.getElementById('btnLogin').focus();
            }

            function alternarVerSenha() {
                const inp = document.getElementById('loginSenha');
                if (inp) inp.type = inp.type === 'password' ? 'text' : 'password';
            }

            async function executarLogin(e) {
                e.preventDefault();
                const btn = document.getElementById('btnLogin');
                const erroBox = document.getElementById('loginErro');
                const erroMsg = document.getElementById('loginErroMsg');
                erroBox.classList.add('hidden');
                btn.disabled = true;

                const username = document.getElementById('loginUsuario').value.trim();
                const password = document.getElementById('loginSenha').value.trim();

                try {
                    const res = await fetch('/api/login', {
                        method: 'POST',
                        headers: {'Content-Type': 'application/json'},
                        body: JSON.stringify({ username, password })
                    });
                    const d = await res.json();
                    if (res.ok && d.status === 'sucesso') {
                        localStorage.setItem('usuario_ps_anil', JSON.stringify(d.usuario));
                        mostrarToast(`Bem-vindo(a), ${d.usuario.nome}!`);
                        verificarAutenticacao();
                    } else {
                        erroMsg.textContent = d.detail || 'Usuário ou senha incorretos.';
                        erroBox.classList.remove('hidden');
                    }
                } catch(err) {
                    erroMsg.textContent = 'Erro ao conectar ao servidor.';
                    erroBox.classList.remove('hidden');
                } finally {
                    btn.disabled = false;
                }
            }

            function fazerLogout() {
                localStorage.removeItem('usuario_ps_anil');
                document.getElementById('telaSistema').classList.add('hidden');
                document.getElementById('telaLogin').classList.remove('hidden');
                document.getElementById('loginSenha').value = '';
                mostrarToast('Sessão encerrada.');
            }

            // ===============================================================
            // VISUALIZAÇÃO DE DETALHES DE REGISTROS (MODAIS)
            // ===============================================================
            async function abrirModalVisualizarManut(id) {
                try {
                    const res = await fetch(`/api/manutencao/${id}`);
                    if (!res.ok) throw new Error('Não encontrado');
                    const reg = await res.json();

                    document.getElementById('viewManutCodigo').textContent = reg.codigo_id || `#${reg.id}`;
                    document.getElementById('viewManutCategoriaSub').textContent = `${reg.categoria} • ${reg.tipo_manutencao}`;
                    document.getElementById('viewManutDataHora').textContent = `${reg.data} ${reg.horario ? 'às ' + reg.horario : ''}`;
                    document.getElementById('viewManutTurno').textContent = reg.turno || '-';
                    document.getElementById('viewManutTipo').textContent = reg.tipo_manutencao || '-';
                    document.getElementById('viewManutItem').textContent = reg.item || '-';
                    document.getElementById('viewManutSetor').textContent = reg.setor || '-';
                    document.getElementById('viewManutQuantidade').textContent = reg.quantidade || 1;
                    document.getElementById('viewManutChamado').textContent = reg.numero_chamado || 'Sem chamado';
                    document.getElementById('viewManutConclusao').textContent = reg.data_conclusao || 'Em aberto';
                    document.getElementById('viewManutRespEntrega').textContent = reg.responsavel_entrega || '-';
                    document.getElementById('viewManutRespRecebimento').textContent = reg.responsavel_recebimento || '-';
                    document.getElementById('viewManutDescricao').textContent = reg.descricao || 'Nenhuma descrição detalhada informada.';
                    document.getElementById('viewManutObservacoes').textContent = reg.observacoes || 'Nenhuma observação registrada.';
                    document.getElementById('viewManutRodapeRef').textContent = `Mês de Referência: ${reg.mes_ano || '-'}`;

                    // Badge de Status
                    const badgeSt = document.getElementById('viewManutStatusBadge');
                    badgeSt.textContent = reg.status || 'Pendente';
                    badgeSt.className = reg.status === 'Concluído' ? 'px-2.5 py-1 rounded-full text-xs font-black bg-emerald-100 text-emerald-800' :
                                        reg.status === 'Ok' ? 'px-2.5 py-1 rounded-full text-xs font-black bg-sky-100 text-sky-800' :
                                        'px-2.5 py-1 rounded-full text-xs font-black bg-amber-100 text-amber-800';

                    // Badges Extras (Retrabalho e Editado)
                    const extras = document.getElementById('viewManutBadgesExtras');
                    extras.innerHTML = '';
                    if (reg.retrabalho) {
                        extras.innerHTML += `<span class="px-2.5 py-1 rounded-full text-xs font-extrabold bg-amber-100 text-amber-900 border border-amber-300">🔄 Retrabalho</span>`;
                    }
                    if (reg.observacoes && reg.observacoes.includes('[Editado em')) {
                        extras.innerHTML += `<span class="px-2.5 py-1 rounded-full text-xs font-bold bg-indigo-100 text-indigo-900 border border-indigo-200">✏️ Histórico de Edição</span>`;
                    }

                    // Botão Editar a partir da visualização
                    const btnEd = document.getElementById('btnEditarFromViewManut');
                    btnEd.onclick = () => {
                        fecharModalVisualizarManut();
                        abrirModalEditarManut(id);
                    };

                    document.getElementById('modalVisualizarManutencao').classList.remove('hidden');
                } catch(e) {
                    mostrarToast('Erro ao carregar detalhes do chamado.', 'erro');
                }
            }

            function fecharModalVisualizarManut() {
                document.getElementById('modalVisualizarManutencao').classList.add('hidden');
            }

            async function abrirModalVisualizarRoupa(id) {
                try {
                    const res = await fetch(`/api/rouparia/${id}`);
                    if (!res.ok) throw new Error('Não encontrado');
                    const reg = await res.json();

                    document.getElementById('viewRoupaCodigo').textContent = reg.codigo_id || `#${reg.id}`;
                    document.getElementById('viewRoupaTipoSub').textContent = `Módulo Rouparia • ${reg.tipo_registro}`;
                    document.getElementById('viewRoupaDataHora').textContent = `${reg.data} ${reg.horario ? 'às ' + reg.horario : ''}`;
                    document.getElementById('viewRoupaTurno').textContent = reg.turno || '-';
                    document.getElementById('viewRoupaTipoMov').textContent = reg.tipo_registro || '-';
                    document.getElementById('viewRoupaItem').textContent = reg.item_enxoval || '-';
                    document.getElementById('viewRoupaSetor').textContent = reg.setor || '-';
                    document.getElementById('viewRoupaQtd').textContent = `${reg.qtd_pecas || 0} peça(s)`;
                    document.getElementById('viewRoupaPeso').textContent = reg.peso_kg > 0 ? `${reg.peso_kg.toFixed(2)} kg` : '-';
                    document.getElementById('viewRoupaPontual').textContent = reg.pontual === 1 ? 'Sim' : reg.pontual === 0 ? 'Não' : '-';
                    document.getElementById('viewRoupaRespUnidade').textContent = reg.responsavel_unidade || '-';
                    document.getElementById('viewRoupaRespTerceirizada').textContent = reg.responsavel_terceirizada || '-';
                    document.getElementById('viewRoupaInconformidade').textContent = reg.inconformidade || 'Nenhuma inconformidade ou observação registrada.';

                    const btnEdR = document.getElementById('btnEditarFromViewRoupa');
                    btnEdR.onclick = () => {
                        fecharModalVisualizarRoupa();
                        abrirModalEditarRoupa(id);
                    };

                    document.getElementById('modalVisualizarRouparia').classList.remove('hidden');
                } catch(e) {
                    mostrarToast('Erro ao carregar detalhes da rouparia.', 'erro');
                }
            }

            function fecharModalVisualizarRoupa() {
                document.getElementById('modalVisualizarRouparia').classList.add('hidden');
            }

            // ===============================================================
            // EDIÇÃO DE REGISTROS DO BANCO DE DADOS (MODAIS)
            // ===============================================================
            async function abrirModalEditarManut(id) {
                try {
                    const res = await fetch(`/api/manutencao/${id}`);
                    if (!res.ok) throw new Error('Não encontrado');
                    const reg = await res.json();
                    
                    document.getElementById('editManutId').value = reg.id;
                    document.getElementById('editManutCodigo').textContent = reg.codigo_id || `#${reg.id}`;
                    document.getElementById('editManutData').value = dataParaIso(reg.data || '');
                    document.getElementById('editManutHorario').value = reg.horario || '';
                    document.getElementById('editManutTurno').value = reg.turno || 'Matutino';
                    document.getElementById('editManutCategoria').value = reg.categoria || '';
                    document.getElementById('editManutTipo').value = reg.tipo_manutencao || 'Corretiva';
                    document.getElementById('editManutItem').value = reg.item || '';
                    document.getElementById('editManutSetor').value = reg.setor || '';
                    document.getElementById('editManutDescricao').value = reg.descricao || '';
                    document.getElementById('editManutQuantidade').value = reg.quantidade || 1;
                    document.getElementById('editManutNumeroChamado').value = reg.numero_chamado || '';
                    document.getElementById('editManutRespEntrega').value = reg.responsavel_entrega || '';
                    document.getElementById('editManutRespRecebimento').value = reg.responsavel_recebimento || '';
                    document.getElementById('editManutStatus').value = reg.status || 'Pendente';
                    document.getElementById('editManutDataConclusao').value = dataParaIso(reg.data_conclusao || '');
                    document.getElementById('editManutObservacoes').value = reg.observacoes || '';
                    document.getElementById('editManutRetrabalho').checked = !!reg.retrabalho;

                    document.getElementById('modalEditarManutencao').classList.remove('hidden');
                } catch(e) {
                    mostrarToast('Erro ao carregar dados do registro de manutenção.', 'erro');
                }
            }

            function fecharModalEditarManut() {
                document.getElementById('modalEditarManutencao').classList.add('hidden');
            }

            async function salvarEdicaoManutencao(e) {
                e.preventDefault();
                const id = document.getElementById('editManutId').value;
                const btn = document.getElementById('btnSalvarEditManut');
                btn.disabled = true;

                // Recuperar usuário logado para auditoria
                let nomeOperador = 'Gestão';
                try {
                    const u = JSON.parse(localStorage.getItem('usuario_ps_anil') || '{}');
                    nomeOperador = u.nome || u.username || 'Gestão';
                } catch(err) {}

                const agoraAudit = `${new Date().toLocaleDateString('pt-BR')} às ${obterHoraAtual()}`;
                const notaAuditoria = `[Editado em ${agoraAudit} por ${nomeOperador}]`;

                let obsForm = document.getElementById('editManutObservacoes').value.trim();
                if (obsForm && !obsForm.includes(notaAuditoria)) {
                    obsForm = `${obsForm}\n${notaAuditoria}`;
                } else if (!obsForm) {
                    obsForm = notaAuditoria;
                }

                const payload = {
                    data: dataParaBr(document.getElementById('editManutData').value.trim()),
                    horario: document.getElementById('editManutHorario').value.trim(),
                    turno: document.getElementById('editManutTurno').value,
                    categoria: document.getElementById('editManutCategoria').value,
                    tipo_manutencao: document.getElementById('editManutTipo').value,
                    item: document.getElementById('editManutItem').value,
                    setor: document.getElementById('editManutSetor').value,
                    descricao: document.getElementById('editManutDescricao').value.trim(),
                    quantidade: parseFloat(document.getElementById('editManutQuantidade').value) || 1.0,
                    numero_chamado: document.getElementById('editManutNumeroChamado').value.trim(),
                    responsavel_entrega: document.getElementById('editManutRespEntrega').value.trim(),
                    responsavel_recebimento: document.getElementById('editManutRespRecebimento').value.trim(),
                    status: document.getElementById('editManutStatus').value,
                    data_conclusao: dataParaBr(document.getElementById('editManutDataConclusao').value.trim()),
                    observacoes: obsForm,
                    retrabalho: document.getElementById('editManutRetrabalho').checked ? 1 : 0
                };

                try {
                    const res = await fetch(`/api/manutencao/${id}`, {
                        method: 'PUT',
                        headers: {'Content-Type': 'application/json'},
                        body: JSON.stringify(payload)
                    });
                    if (res.ok) {
                        mostrarToast('Registro atualizado com sucesso e observação registrada!');
                        fecharModalEditarManut();
                        carregarHistorico();
                        carregarConsolidado();
                        carregarRelatorios();
                    } else {
                        const err = await res.json();
                        mostrarToast(err.detail || 'Erro ao salvar alterações.', 'erro');
                    }
                } catch(err) {
                    mostrarToast('Falha na comunicação ao salvar edição.', 'erro');
                } finally {
                    btn.disabled = false;
                }
            }

            async function abrirModalEditarRoupa(id) {
                try {
                    const res = await fetch(`/api/rouparia/${id}`);
                    if (!res.ok) throw new Error('Não encontrado');
                    const reg = await res.json();

                    document.getElementById('editRoupaId').value = reg.id;
                    document.getElementById('editRoupaCodigo').textContent = reg.codigo_id || `#${reg.id}`;
                    document.getElementById('editRoupaData').value = dataParaIso(reg.data || '');
                    document.getElementById('editRoupaHorario').value = reg.horario || '';
                    document.getElementById('editRoupaTurno').value = reg.turno || 'Matutino';
                    document.getElementById('editRoupaTipo').value = reg.tipo_registro || 'Estoque';
                    document.getElementById('editRoupaItem').value = reg.item_enxoval || '';
                    document.getElementById('editRoupaSetor').value = reg.setor || '';
                    document.getElementById('editRoupaQtdPecas').value = reg.qtd_pecas || 0;
                    document.getElementById('editRoupaPesoKg').value = reg.peso_kg || 0;
                    document.getElementById('editRoupaRespUnidade').value = reg.responsavel_unidade || '';
                    document.getElementById('editRoupaRespTerceirizada').value = reg.responsavel_terceirizada || 'LAVARE';
                    document.getElementById('editRoupaInconformidade').value = reg.inconformidade || '';

                    document.getElementById('modalEditarRouparia').classList.remove('hidden');
                } catch(e) {
                    mostrarToast('Erro ao carregar movimentação de rouparia.', 'erro');
                }
            }

            function fecharModalEditarRoupa() {
                document.getElementById('modalEditarRouparia').classList.add('hidden');
            }

            async function salvarEdicaoRouparia(e) {
                e.preventDefault();
                const id = document.getElementById('editRoupaId').value;
                const btn = document.getElementById('btnSalvarEditRoupa');
                btn.disabled = true;

                let nomeOperador = 'Gestão';
                try {
                    const u = JSON.parse(localStorage.getItem('usuario_ps_anil') || '{}');
                    nomeOperador = u.nome || u.username || 'Gestão';
                } catch(err) {}

                const agoraAudit = `${new Date().toLocaleDateString('pt-BR')} às ${obterHoraAtual()}`;
                const notaAuditoria = `[Editado em ${agoraAudit} por ${nomeOperador}]`;

                let inconfForm = document.getElementById('editRoupaInconformidade').value.trim();
                if (inconfForm && !inconfForm.includes(notaAuditoria)) {
                    inconfForm = `${inconfForm}\n${notaAuditoria}`;
                } else if (!inconfForm) {
                    inconfForm = notaAuditoria;
                }

                const payload = {
                    data: dataParaBr(document.getElementById('editRoupaData').value.trim()),
                    horario: document.getElementById('editRoupaHorario').value.trim(),
                    turno: document.getElementById('editRoupaTurno').value,
                    tipo_registro: document.getElementById('editRoupaTipo').value,
                    item_enxoval: document.getElementById('editRoupaItem').value,
                    setor: document.getElementById('editRoupaSetor').value,
                    qtd_pecas: parseFloat(document.getElementById('editRoupaQtdPecas').value) || 0.0,
                    peso_kg: parseFloat(document.getElementById('editRoupaPesoKg').value) || 0.0,
                    responsavel_unidade: document.getElementById('editRoupaRespUnidade').value.trim(),
                    responsavel_terceirizada: document.getElementById('editRoupaRespTerceirizada').value.trim(),
                    inconformidade: inconfForm
                };

                try {
                    const res = await fetch(`/api/rouparia/${id}`, {
                        method: 'PUT',
                        headers: {'Content-Type': 'application/json'},
                        body: JSON.stringify(payload)
                    });
                    if (res.ok) {
                        mostrarToast('Movimentação de rouparia atualizada com sucesso no banco!');
                        fecharModalEditarRoupa();
                        carregarHistorico();
                        carregarConsolidado();
                        carregarRelatorios();
                    } else {
                        const err = await res.json();
                        mostrarToast(err.detail || 'Erro ao salvar alterações.', 'erro');
                    }
                } catch(err) {
                    mostrarToast('Falha na comunicação ao salvar edição.', 'erro');
                } finally {
                    btn.disabled = false;
                }
            }

            // Inicialização com Verificação de Login
            window.addEventListener('DOMContentLoaded', () => {
                verificarAutenticacao();
            });
        </script>
    </body>
    </html>
    """

if __name__ == "__main__":
    import uvicorn
    # Obter porta do Render ($PORT) ou 10000 no Render ou 8000 local
    porta = int(os.environ.get("PORT", 10000 if os.environ.get("RENDER") else 8000))
    print(f"🚀 Iniciando servidor FastAPI na porta {porta} (0.0.0.0)...")
    uvicorn.run("main:app", host="0.0.0.0", port=porta, reload=False)
