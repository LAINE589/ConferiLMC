"""
Modelos do banco de dados – Sistema LMC SaaS
Hierarquia: Admin → Contabilidade → Posto
"""
from flask_sqlalchemy import SQLAlchemy
from flask_login import UserMixin
from werkzeug.security import generate_password_hash, check_password_hash
from datetime import datetime

db = SQLAlchemy()

PERFIL_ADMIN         = "admin"
PERFIL_CONTABILIDADE = "contabilidade"
PERFIL_POSTO         = "posto"

PLANO_MENSAL = "mensal"
PLANO_ANUAL  = "anual"
PLANO_UNICO  = "unico"


class Contabilidade(db.Model):
    __tablename__ = "contabilidades"
    id         = db.Column(db.Integer, primary_key=True)
    nome       = db.Column(db.String(200), nullable=False)
    cnpj       = db.Column(db.String(20), unique=True, nullable=False)
    email      = db.Column(db.String(200), unique=True, nullable=False)
    telefone   = db.Column(db.String(30))
    ativa      = db.Column(db.Boolean, default=True)
    criada_em  = db.Column(db.DateTime, default=datetime.utcnow)
    usuarios   = db.relationship("Usuario", back_populates="contabilidade",
                                 foreign_keys="Usuario.contabilidade_id")
    postos     = db.relationship("Posto", back_populates="contabilidade")


class Posto(db.Model):
    __tablename__ = "postos"
    id               = db.Column(db.Integer, primary_key=True)
    cnpj             = db.Column(db.String(20), unique=True, nullable=False)
    razao_social     = db.Column(db.String(200), nullable=False)
    nome_fantasia    = db.Column(db.String(200))
    email            = db.Column(db.String(200))
    telefone         = db.Column(db.String(30))
    cidade           = db.Column(db.String(100))
    estado           = db.Column(db.String(2))
    ativo            = db.Column(db.Boolean, default=True)
    plano            = db.Column(db.String(20), default=PLANO_MENSAL)
    licenca_ativa    = db.Column(db.Boolean, default=True)
    licenca_expira   = db.Column(db.DateTime, nullable=True)
    contabilidade_id = db.Column(db.Integer, db.ForeignKey("contabilidades.id"), nullable=True)
    contabilidade    = db.relationship("Contabilidade", back_populates="postos")
    criado_em        = db.Column(db.DateTime, default=datetime.utcnow)
    usuarios         = db.relationship("Usuario", back_populates="posto",
                                       foreign_keys="Usuario.posto_id")
    relatorios       = db.relationship("Relatorio", back_populates="posto",
                                       order_by="Relatorio.gerado_em.desc()")


class Usuario(db.Model, UserMixin):
    __tablename__ = "usuarios"
    id               = db.Column(db.Integer, primary_key=True)
    email            = db.Column(db.String(200), unique=True, nullable=False)
    senha_hash       = db.Column(db.String(256), nullable=False)
    nome             = db.Column(db.String(200), nullable=False)
    perfil           = db.Column(db.String(20), nullable=False)
    ativo            = db.Column(db.Boolean, default=True)
    criado_em        = db.Column(db.DateTime, default=datetime.utcnow)
    ultimo_acesso    = db.Column(db.DateTime, nullable=True)
    contabilidade_id = db.Column(db.Integer, db.ForeignKey("contabilidades.id"), nullable=True)
    posto_id         = db.Column(db.Integer, db.ForeignKey("postos.id"), nullable=True)
    contabilidade    = db.relationship("Contabilidade", back_populates="usuarios",
                                       foreign_keys=[contabilidade_id])
    posto            = db.relationship("Posto", back_populates="usuarios",
                                       foreign_keys=[posto_id])

    def set_senha(self, senha):
        self.senha_hash = generate_password_hash(senha)

    def check_senha(self, senha):
        return check_password_hash(self.senha_hash, senha)

    @property
    def is_admin(self):         return self.perfil == PERFIL_ADMIN
    @property
    def is_contabilidade(self): return self.perfil == PERFIL_CONTABILIDADE
    @property
    def is_posto(self):         return self.perfil == PERFIL_POSTO

    # Retorna nome da organização do usuário
    @property
    def organizacao(self):
        if self.is_posto and self.posto:
            return self.posto.razao_social
        if self.is_contabilidade and self.contabilidade:
            return self.contabilidade.nome
        return "Administração"


class Relatorio(db.Model):
    __tablename__ = "relatorios"
    id                 = db.Column(db.Integer, primary_key=True)
    posto_id           = db.Column(db.Integer, db.ForeignKey("postos.id"), nullable=False)
    posto              = db.relationship("Posto", back_populates="relatorios")
    competencia_ant    = db.Column(db.String(10), nullable=True)
    competencia_atu    = db.Column(db.String(10), nullable=False)
    gerado_em          = db.Column(db.DateTime, default=datetime.utcnow)
    gerado_por         = db.Column(db.String(200))
    total_divergencias = db.Column(db.Integer, default=0)
    tem_dac            = db.Column(db.Boolean, default=False)
    status_geral       = db.Column(db.String(20), default="ok")  # ok / alerta / critico
    arquivo_nome       = db.Column(db.String(300), nullable=True)

    @property
    def status_icone(self):
        return {"ok": "✅", "alerta": "⚠️", "critico": "❌"}.get(self.status_geral, "—")
