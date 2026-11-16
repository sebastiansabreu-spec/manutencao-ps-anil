"""
Script de Migração dos Dados Locais (hospital.db) para o Neon PostgreSQL
Uso:
  python migrar_para_neon.py "sua-url-do-neon"
  ou defina a variável de ambiente DATABASE_URL
"""
import sys
import os
import sqlite3
from sqlalchemy import create_engine
from sqlalchemy.orm import sessionmaker

# Importar modelos do app
import main
from main import Base, Manutencao, Rouparia, ConfiguracaoItem, Parametro, Usuario

def migrar(neon_url=None):
    url_destino = neon_url or os.getenv("DATABASE_URL")
    if not url_destino or "sqlite" in url_destino:
        print("❌ Por favor, informe a URL de conexão do Neon PostgreSQL.")
        print('Exemplo: python migrar_para_neon.py "postgresql://user:senha@ep-xyz.us-east-2.aws.neon.tech/neondb?sslmode=require"')
        return

    if url_destino.startswith("postgres://"):
        url_destino = url_destino.replace("postgres://", "postgresql://", 1)

    print("🔌 Conectando ao banco de dados SQLite local (hospital.db)...")
    origem_engine = create_engine("sqlite:///./hospital.db", connect_args={"check_same_thread": False})
    SessionOrigem = sessionmaker(bind=origem_engine)
    db_origem = SessionOrigem()

    print("☁️ Conectando ao Neon PostgreSQL...")
    try:
        destino_engine = create_engine(url_destino, pool_pre_ping=True)
        Base.metadata.create_all(bind=destino_engine)
        SessionDestino = sessionmaker(bind=destino_engine)
        db_destino = SessionDestino()
        print("✅ Conectado ao Neon com sucesso! Tabelas criadas/verificadas.")
    except Exception as e:
        print(f"❌ Erro ao conectar ao Neon: {e}")
        return

    print("\n📦 Iniciando transferência dos dados...")

    # 1. Configurações
    itens_cfg = db_origem.query(ConfiguracaoItem).all()
    db_destino.query(ConfiguracaoItem).delete()
    for item in itens_cfg:
        db_destino.add(ConfiguracaoItem(
            tipo=item.tipo,
            nome=item.nome,
            detalhe=item.detalhe,
            ativo=item.ativo
        ))
    db_destino.commit()
    print(f"  ✓ {len(itens_cfg)} itens de configuração transferidos.")

    # 2. Parâmetros
    params = db_origem.query(Parametro).all()
    db_destino.query(Parametro).delete()
    for p in params:
        db_destino.add(Parametro(
            chave=p.chave,
            valor=p.valor,
            descricao=p.descricao
        ))
    db_destino.commit()
    print(f"  ✓ {len(params)} parâmetros transferidos.")

    # 3. Manutenções
    manuts = db_origem.query(Manutencao).all()
    db_destino.query(Manutencao).delete()
    for m in manuts:
        db_destino.add(Manutencao(
            codigo_id=m.codigo_id,
            data=m.data,
            mes_ano=m.mes_ano,
            horario=m.horario,
            turno=m.turno,
            categoria=m.categoria,
            tipo_manutencao=m.tipo_manutencao,
            item=m.item,
            setor=m.setor,
            descricao=m.descricao,
            quantidade=m.quantidade,
            numero_chamado=m.numero_chamado,
            responsavel_entrega=m.responsavel_entrega,
            responsavel_recebimento=m.responsavel_recebimento,
            status=m.status,
            data_conclusao=m.data_conclusao,
            observacoes=m.observacoes,
            retrabalho=m.retrabalho
        ))
    db_destino.commit()
    print(f"  ✓ {len(manuts)} registros de Manutenção transferidos.")

    # 4. Rouparia
    roupas = db_origem.query(Rouparia).all()
    db_destino.query(Rouparia).delete()
    for r in roupas:
        db_destino.add(Rouparia(
            codigo_id=r.codigo_id,
            data=r.data,
            mes_ano=r.mes_ano,
            horario=r.horario,
            turno=r.turno,
            tipo_registro=r.tipo_registro,
            item_enxoval=r.item_enxoval,
            setor=r.setor,
            qtd_pecas=r.qtd_pecas,
            qtd_hamper=r.qtd_hamper,
            peso_kg=r.peso_kg,
            responsavel_unidade=r.responsavel_unidade,
            responsavel_terceirizada=r.responsavel_terceirizada,
            inconformidade=r.inconformidade,
            pontual=r.pontual
        ))
    db_destino.commit()
    print(f"  ✓ {len(roupas)} registros de Rouparia transferidos.")

    # 5. Usuários
    users = db_origem.query(Usuario).all()
    db_destino.query(Usuario).delete()
    for u in users:
        db_destino.add(Usuario(
            username=u.username,
            nome=u.nome,
            senha_hash=u.senha_hash,
            perfil=u.perfil,
            ativo=u.ativo
        ))
    db_destino.commit()
    print(f"  ✓ {len(users)} usuários transferidos.")

    db_origem.close()
    db_destino.close()
    print("\n🎉 MIGRAÇÃO CONCLUÍDA COM 100% DE SUCESSO NO NEON POSTGRESQL!")

if __name__ == "__main__":
    url_cli = sys.argv[1] if len(sys.argv) > 1 else None
    migrar(url_cli)
