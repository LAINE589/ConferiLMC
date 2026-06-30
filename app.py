import os
import io
from datetime import datetime, timedelta
from functools import wraps

from flask import (Flask, render_template, request, redirect,
                   url_for, session, send_file, flash)
from werkzeug.security import generate_password_hash, check_password_hash

import openpyxl
try:
    from dac_reader import ler_dac, confrontar_dac_sped
    DAC_DISPONIVEL = True
except Exception:
    DAC_DISPONIVEL = False
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ─────────────────────────────────────────────────────────────────────────────
# CONFIG
# ─────────────────────────────────────────────────────────────────────────────
app = Flask(__name__)
# SECRET_KEY fixa garante que sessões funcionem com múltiplos workers
app.secret_key = os.environ.get("SECRET_KEY", "LMC-SPED-2026-K9x#mQpZ")
app.permanent_session_lifetime = timedelta(hours=8)

# ── Usuários ──────────────────────────────────────────────────────────────────
# Senhas definidas aqui diretamente.
# Para trocar: edite o segundo argumento de generate_password_hash()
USUARIOS = {
    "fiscal@cleodoncontabilidade.com.br": {
        "senha_hash": generate_password_hash("Cld@123"),
        "nome": "Fiscal",
    },
    "lucroreal@cleodoncontabilidade.com.br": {
        "senha_hash": generate_password_hash("Cld@123"),
        "nome": "Lucro Real",
    },
}

# ─────────────────────────────────────────────────────────────────────────────
# DECORADOR DE LOGIN
# ─────────────────────────────────────────────────────────────────────────────
def login_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if "usuario" not in session:
            flash("Faça login para continuar.", "warning")
            return redirect(url_for("login"))
        return f(*args, **kwargs)
    return decorated

# ─────────────────────────────────────────────────────────────────────────────
# ROTAS
# ─────────────────────────────────────────────────────────────────────────────
@app.route("/", methods=["GET", "POST"])
def login():
    if "usuario" in session:
        return redirect(url_for("index"))

    erro = None
    if request.method == "POST":
        user  = request.form.get("usuario", "").strip()
        senha = request.form.get("senha", "").strip()

        if user in USUARIOS and check_password_hash(USUARIOS[user]["senha_hash"], senha):
            session.permanent = True
            session["usuario"] = user
            session["nome"]    = USUARIOS[user]["nome"]
            return redirect(url_for("index"))

        erro = "Usuário ou senha incorretos. Verifique e tente novamente."

    return render_template("login.html", erro=erro)


@app.route("/logout")
def logout():
    session.clear()
    flash("Você saiu do sistema.", "success")
    return redirect(url_for("login"))


@app.route("/sistema")
@login_required
def index():
    return render_template("index.html", nome=session.get("nome"))


@app.route("/processar", methods=["POST"])
@login_required
def processar():
    arq_ant = request.files.get("ant")
    arq_atu = request.files.get("atu")

    # SPED da competência anterior agora é opcional.
    # Se ausente, o sistema confronta apenas o mês atual com o DAC,
    # mantendo consistência diária, negativos, versão/capacidade e ANP.
    tem_ant = bool(arq_ant and arq_ant.filename)

    if not arq_atu or not arq_atu.filename:
        flash("Selecione ao menos o arquivo SPED da competência atual.", "danger")
        return redirect(url_for("index"))

    try:
        bytes_atu = arq_atu.read()
        d_atu = ler_sped_bytes(bytes_atu)

        if tem_ant:
            bytes_ant = arq_ant.read()
            d_ant   = ler_sped_bytes(bytes_ant)
            neg_abr = verificar_negativos_bytes(bytes_ant)
        else:
            # Estrutura vazia equivalente — confronto_mensal já trata isso
            # retornando listas vazias quando não há dados do mês anterior.
            d_ant   = {"info": {}, "tanques": {}, "bicos": {}}
            neg_abr = {"tanques": [], "bicos": []}
            flash("SPED da competência anterior não enviado — gerando apenas a "
                  "conferência da competência atual (sem confronto entre meses).",
                  "warning")

        conf_m  = confronto_mensal(d_ant, d_atu)
        d_mai   = confronto_diario(d_atu)
        neg_mai = verificar_negativos_bytes(bytes_atu)
        vc_mai  = verificar_versao_capacidade(d_atu)

        # DAC opcional
        arq_dac = request.files.get("dac")
        conf_dac = None
        erro_dac = None
        if arq_dac and arq_dac.filename:
            try:
                bytes_dac = arq_dac.read()
                dac_dados = ler_dac(bytes_dac, arq_dac.filename)
                conf_dac  = confrontar_dac_sped(dac_dados, d_atu)
            except Exception as e:
                erro_dac = str(e)

        wb  = openpyxl.Workbook()
        ws1 = wb.active; ws1.title = "Resumo"
        aba_resumo(ws1, conf_m, d_mai, d_ant, d_atu, neg_abr, neg_mai, vc_mai)
        ws2 = wb.create_sheet("Confronto Meses")
        aba_mensal(ws2, conf_m, d_ant, d_atu)
        ws3 = wb.create_sheet("Comparativo Diário")
        aba_diario(ws3, d_mai)
        if conf_dac:
            ws4 = wb.create_sheet("DAC × SPED")
            aba_dac(ws4, conf_dac, d_atu)
            ws5 = wb.create_sheet("DAC do SPED")
            aba_dac_sped(ws5, d_atu, d_atu)
        else:
            # Sem DAC real: gerar DAC de acompanhamento a partir do SPED
            ws4 = wb.create_sheet("DAC do SPED")
            aba_dac_sped(ws4, d_atu, d_atu)
        if erro_dac:
            flash(f"Aviso DAC: {erro_dac}", "warning")

        buf = io.BytesIO()
        wb.save(buf); buf.seek(0)

        nome = f"Relatorio_LMC_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        return send_file(buf,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            as_attachment=True, download_name=nome)

    except Exception as e:
        flash(f"Erro ao processar os arquivos: {str(e)}", "danger")
        return redirect(url_for("index"))


# ─────────────────────────────────────────────────────────────────────────────from datetime import datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

C_AZUL_ESC = "1F3864"; C_AZUL_MED = "2E75B6"
C_VERDE_BG = "C6EFCE"; C_VERDE_FG = "375623"
C_VERM_BG  = "FFC7CE"; C_VERM_FG  = "9C0006"
C_AMAR_BG  = "FFEB9C"; C_AMAR_FG  = "7D6608"
C_CINZA    = "D9D9D9"
NF = "#,##0.000"
VERSAO_OBRIGATORIA = "020"

def _sk(x):
    return int(x) if str(x).isdigit() else x

def _nid(raw):
    try:
        return str(int(raw.strip()))
    except:
        return raw.strip()

def _brd():
    s = Side(style="thin", color="BFBFBF")
    return Border(left=s, right=s, top=s, bottom=s)

def _fl(v):
    return float(v.strip().replace(",", ".")) if v.strip() else 0.0

def _dt(s):
    s = s.strip()
    return datetime.strptime(s, "%d%m%Y").date() if len(s) == 8 else None

def _st(dif):
    if dif is None: return "⚠️ AUSENTE"
    return "✅ OK" if dif == 0 else "❌ DIVERGÊNCIA"

def _row_bg(status):
    return {"✅ OK": C_VERDE_BG, "❌ DIVERGÊNCIA": C_VERM_BG, "⚠️ AUSENTE": C_AMAR_BG}.get(status)

def _label(mapa, id_raw, tipo):
    """Retorna 'Tanque 1', 'Bico 3' etc. usando o mapa ordinal."""
    return f"{tipo} {mapa.get(id_raw, id_raw)}"

# ── LEITURA ────────────────────────────────────────────────────────────────────
def ler_sped(caminho):
    with open(caminho, encoding="latin-1", errors="replace") as f:
        text = f.read()
    tanques={}; bicos={}; info={}
    data_atual=None; vals_1300={}

    for linha in text.splitlines():
        c=linha.strip().split("|")
        if len(c)<2: continue
        tp=c[1]

        if tp=="0000":
            info={"versao":c[2].strip() if len(c)>2 else "",
                  "razao": c[6].strip() if len(c)>6 else "",
                  "cnpj":  c[7].strip() if len(c)>7 else "",
                  "dt_ini":c[4].strip() if len(c)>4 else "",
                  "dt_fin":c[5].strip() if len(c)>5 else ""}

        elif tp=="1300":
            data_atual=_dt(c[3]) if len(c)>3 else None
            if not data_atual: continue
            vals_1300={
                "data":data_atual,
                "est_abert":_fl(c[4]),"entrada":_fl(c[5]),"saida":_fl(c[7]),
                "evap":_fl(c[9]) if len(c)>9 else 0.0,
                "ajuste":_fl(c[10]) if len(c)>10 else 0.0,
                "est_fech":_fl(c[11]) if len(c)>11 else 0.0,
            }

        elif tp=="1310":
            if not vals_1300: continue
            t=_nid(c[2])
            tem_campos_proprios = len(c) > 10 and c[3].strip() != ""
            if tem_campos_proprios:
                try:
                    est_abert_1310 = _fl(c[3])
                    entrada_1310   = _fl(c[4]) if len(c)>4 else 0.0
                    saida_1310     = _fl(c[6]) if len(c)>6 else 0.0
                    evap_1310      = _fl(c[8]) if len(c)>8 else 0.0
                    ajuste_1310    = _fl(c[9]) if len(c)>9 else 0.0
                    est_fech_1310  = _fl(c[10]) if len(c)>10 else 0.0
                    cap            = _fl(c[11]) if len(c)>11 and c[11].strip() else None
                except Exception:
                    tem_campos_proprios = False

            if tem_campos_proprios:
                key=(t,vals_1300["data"])
                tanques[key]={
                    "tanque":t,"data":vals_1300["data"],
                    "est_abert":est_abert_1310,"entrada":entrada_1310,
                    "saida":saida_1310,"evap":evap_1310,
                    "ajuste":ajuste_1310,"est_fech":est_fech_1310,
                    "capacidade":cap,
                }
            else:
                cap=_fl(c[11]) if len(c)>11 and c[11].strip() else None
                key=(t,vals_1300["data"])
                tanques[key]={
                    "tanque":t,"data":vals_1300["data"],
                    "est_abert":vals_1300["est_abert"],"entrada":vals_1300["entrada"],
                    "saida":vals_1300["saida"],"evap":vals_1300["evap"],
                    "ajuste":vals_1300["ajuste"],"est_fech":vals_1300["est_fech"],
                    "capacidade":cap,
                }

        elif tp=="1320":
            b=_nid(c[2])
            if not data_atual: continue
            bicos[(b,data_atual)]={
                "bico":b,"data":data_atual,
                "enc_abert":_fl(c[9]) if len(c)>9 else 0.0,
                "enc_fech": _fl(c[8]) if len(c)>8 else 0.0,
            }

    return {"info":info,"tanques":tanques,"bicos":bicos}


def confronto_mensal(d_ant, d_atu):
    res = {"tanques": [], "bicos": []}
    dt_ant = sorted(set(d for (_,d) in d_ant["tanques"]))
    dt_atu = sorted(set(d for (_,d) in d_atu["tanques"]))
    if not dt_ant or not dt_atu: return res
    ult=dt_ant[-1]; pri=dt_atu[0]

    tanq_pri={}; tanq_ult={}
    for (t,d) in d_atu["tanques"]:
        if t not in tanq_pri or d<tanq_pri[t]: tanq_pri[t]=d
        if t not in tanq_ult or d>tanq_ult[t]: tanq_ult[t]=d

    for t in sorted(set([t for (t,d) in d_ant["tanques"] if d==ult]+
                        list(tanq_pri.keys())), key=_sk):
        fa   = d_ant["tanques"].get((t,ult),{})
        fech = fa.get("est_fech")
        pri_t=tanq_pri.get(t); ult_t=tanq_ult.get(t)
        aa   = d_atu["tanques"].get((t,pri_t),{}) if pri_t else {}
        fu   = d_atu["tanques"].get((t,ult_t),{}) if ult_t else {}
        aber     = aa.get("est_abert")
        fech_atu = fu.get("est_fech")
        dif = round(aber-fech,3) if fech is not None and aber is not None else None

        # Cálculo ANP
        dias_t = sorted(d for (tt,d) in d_atu["tanques"] if tt==t)
        est_ini    = d_atu["tanques"][(t,dias_t[0])]["est_abert"]  if dias_t else None
        est_fin_sp = d_atu["tanques"][(t,dias_t[-1])]["est_fech"]  if dias_t else None
        total_rec  = round(sum(d_atu["tanques"][(t,d)]["entrada"] for d in dias_t),3)
        total_sai  = round(sum(d_atu["tanques"][(t,d)]["saida"]   for d in dias_t),3)
        if est_ini is not None and est_fin_sp is not None:
            saldo_calc    = round(est_ini+total_rec-total_sai,3)
            diferenca_anp = round(saldo_calc-est_fin_sp,3)
            limite_anp    = round(total_rec*0.006,3)
            pct_anp       = round(abs(diferenca_anp)/total_rec*100,4) if total_rec>0 else 0
            if abs(diferenca_anp)<=limite_anp:
                status_anp="✅ DENTRO DO LIMITE"
            elif diferenca_anp>0:
                status_anp="⚠️ SOBRA ACIMA 0,6%"
            else:
                status_anp="❌ FALTA ACIMA 0,6%"
        else:
            diferenca_anp=limite_anp=pct_anp=None; status_anp="⚠️ AUSENTE"

        res["tanques"].append({
            "id":t,"dt_fech":ult,"fech":fech,"dt_aber":pri_t,"aber":aber,
            "dif":dif,"status":_st(dif),"dt_fech_atu":ult_t,"fech_atu":fech_atu,
            "total_rec":total_rec,"total_sai":total_sai,
            "diferenca_anp":diferenca_anp,"limite_anp":limite_anp,
            "pct_anp":pct_anp,"status_anp":status_anp,
        })

    bico_pri={}; bico_ult={}
    for (b,d) in d_atu["bicos"]:
        if b not in bico_pri or d<bico_pri[b]: bico_pri[b]=d
        if b not in bico_ult or d>bico_ult[b]: bico_ult[b]=d

    db_ant=sorted(set(d for (_,d) in d_ant["bicos"]))
    ult_b=db_ant[-1] if db_ant else ult

    for b in sorted(set([b for (b,d) in d_ant["bicos"] if d==ult_b]+
                        list(bico_pri.keys())), key=_sk):
        fa  = d_ant["bicos"].get((b,ult_b),{})
        fech= fa.get("enc_fech")
        pri_b=bico_pri.get(b); ult_b2=bico_ult.get(b)
        aa  = d_atu["bicos"].get((b,pri_b),{}) if pri_b else {}
        fu  = d_atu["bicos"].get((b,ult_b2),{}) if ult_b2 else {}
        aber    = aa.get("enc_abert")
        fech_atu= fu.get("enc_fech")
        dif = round(aber-fech,3) if fech is not None and aber is not None else None
        res["bicos"].append({
            "id":b,"dt_fech":ult_b,"fech":fech,"dt_aber":pri_b,"aber":aber,
            "dif":dif,"status":_st(dif),"dt_fech_atu":ult_b2,"fech_atu":fech_atu,
        })
    return res

def confronto_diario(dados):
    res_t = []; res_b = []
    for tanque in sorted(set(t for (t,_) in dados["tanques"]), key=lambda x: int(x) if x.isdigit() else x):
        dias = sorted(d for (t,d) in dados["tanques"] if t==tanque)
        for i in range(len(dias)-1):
            d1,d2 = dias[i],dias[i+1]
            r1=dados["tanques"][(tanque,d1)]; r2=dados["tanques"][(tanque,d2)]
            fech=r1["est_fech"]; aber=r2["est_abert"]; dif=round(aber-fech,3)
            res_t.append({"tanque":tanque,"dia_fech":d1,"fech":fech,"dia_aber":d2,"aber":aber,"dif":dif,"status":_st(dif),})
    for bico in sorted(set(b for (b,_) in dados["bicos"]), key=lambda x: int(x) if x.isdigit() else x):
        dias = sorted(d for (b,d) in dados["bicos"] if b==bico)
        for i in range(len(dias)-1):
            d1,d2 = dias[i],dias[i+1]
            r1=dados["bicos"][(bico,d1)]; r2=dados["bicos"][(bico,d2)]
            fech=r1["enc_fech"]; aber=r2["enc_abert"]; dif=round(aber-fech,3)
            res_b.append({"bico":bico,"dia_fech":d1,"fech":fech,"dia_aber":d2,"aber":aber,"dif":dif,"status":_st(dif),})
    return {"tanques": res_t, "bicos": res_b}

# ── NEGATIVOS ──────────────────────────────────────────────────────────────────
CAMPOS_1300 = {5:"Est. Abertura",6:"Entrada",7:"Est. Aber. Pós Entrada",8:"Saída",
               9:"Est. Fech. Pré Ajuste",10:"Evaporação",11:"Ajuste",12:"Est. Fechamento Final"}
CAMPOS_1320 = {8:"Enc. Fechamento",9:"Enc. Abertura",10:"Volume Vendido",11:"Diferença Encerrante"}

def verificar_negativos(caminho):
    neg_t=[]; neg_b=[]; data_atual=None
    with open(caminho, encoding="latin-1", errors="replace") as f:
        for n, linha in enumerate(f, 1):
            c = linha.strip().split("|")
            if len(c)<2: continue
            tp=c[1]
            if tp=="1300":
                data_atual=_dt(c[3]) if len(c)>3 else None
                tanque=_nid(c[2])
                for idx,nome in CAMPOS_1300.items():
                    if idx>=len(c): continue
                    try:
                        v=_fl(c[idx])
                        if v<0: neg_t.append({"tanque":tanque,"data":data_atual,"campo":nome,"valor":v,"linha":n})
                    except: pass
            elif tp=="1320":
                bico=_nid(c[2])
                for idx,nome in CAMPOS_1320.items():
                    if idx>=len(c): continue
                    try:
                        v=_fl(c[idx])
                        if v<0: neg_b.append({"bico":bico,"data":data_atual,"campo":nome,"valor":v,"linha":n})
                    except: pass
    return {"tanques": neg_t, "bicos": neg_b}

# ── VERSÃO E CAPACIDADE ────────────────────────────────────────────────────────
def verificar_versao_capacidade(dados):
    info=dados["info"]; versao=info.get("versao",""); dt_ini=info.get("dt_ini","")
    periodo=f"{dt_ini[2:4]}/{dt_ini[4:]}" if len(dt_ini)==8 else dt_ini
    ano=int(dt_ini[4:]) if len(dt_ini)==8 else 0
    cap_obrig = ano >= 2026
    tanques_ids=sorted(set(t for (t,_) in dados["tanques"]), key=lambda x: int(x) if x.isdigit() else x)
    cap_tanques=[]
    for t in tanques_ids:
        caps=set()
        for d in sorted(d for (tt,d) in dados["tanques"] if tt==t):
            c=dados["tanques"].get((t,d),{}).get("capacidade")
            if c is not None: caps.add(c)
        if not caps:
            st="❌ AUSENTE" if cap_obrig else "⚠️ NÃO DECLARADA"; obs="Capacidade não informada"
        elif len(caps)>1:
            st="⚠️ INCONSISTENTE"; obs=f"Valores distintos: {sorted(caps)}"
        elif list(caps)[0]<=0:
            st="❌ INVÁLIDA"; obs=f"Valor zero/negativo: {list(caps)[0]}"
        else:
            st="✅ OK"; obs=f"{list(caps)[0]:,.0f} L"
        cap_tanques.append({"tanque":t,"caps":sorted(caps),"status":st,"obs":obs})
    return {"versao":versao,"versao_ok":versao==VERSAO_OBRIGATORIA,
            "periodo":periodo,"cap_obrig":cap_obrig,"tanques":cap_tanques}

# ── HELPERS EXCEL ──────────────────────────────────────────────────────────────
def _brd_cel(ws,r,c,val,bg=None,fg="000000",bold=False,sz=10,align="center",wrap=False,fmt=None):
    cel=ws.cell(row=r,column=c,value=val)
    cel.font=Font(name="Arial",size=sz,bold=bold,color=fg)
    cel.alignment=Alignment(horizontal=align,vertical="center",wrap_text=wrap)
    cel.border=_brd()
    if bg: cel.fill=PatternFill("solid",start_color=bg)
    if fmt: cel.number_format=fmt
    return cel

def _ch(ws,r,c,val,bg=C_AZUL_ESC,fg="FFFFFF",sz=10):
    cel=ws.cell(row=r,column=c,value=val)
    cel.font=Font(name="Arial",bold=True,color=fg,size=sz)
    cel.fill=PatternFill("solid",start_color=bg)
    cel.alignment=Alignment(horizontal="center",vertical="center",wrap_text=True)
    cel.border=_brd(); return cel

def _dc(ws,r,c,val,fmt=None,bg=None):
    return _brd_cel(ws,r,c,val,bg=bg,fmt=fmt)

def _sc(ws,r,c,status):
    mapa={"✅ OK":(C_VERDE_BG,C_VERDE_FG),"❌ DIVERGÊNCIA":(C_VERM_BG,C_VERM_FG),
          "⚠️ AUSENTE":(C_AMAR_BG,C_AMAR_FG),"❌ AUSENTE":(C_VERM_BG,C_VERM_FG),
          "⚠️ INCONSISTENTE":(C_AMAR_BG,C_AMAR_FG),"❌ INVÁLIDA":(C_VERM_BG,C_VERM_FG)}
    bg,fg=mapa.get(status,(C_CINZA,"000000"))
    cel=ws.cell(row=r,column=c,value=status)
    cel.font=Font(name="Arial",bold=True,size=10,color=fg)
    cel.fill=PatternFill("solid",start_color=bg)
    cel.alignment=Alignment(horizontal="center",vertical="center")
    cel.border=_brd(); return cel

def _titulo(ws,r,texto,n,sz=12):
    ws.merge_cells(start_row=r,start_column=1,end_row=r,end_column=n)
    c=ws.cell(row=r,column=1,value=texto)
    c.font=Font(name="Arial",bold=True,size=sz,color="FFFFFF")
    c.fill=PatternFill("solid",start_color=C_AZUL_ESC)
    c.alignment=Alignment(horizontal="center",vertical="center")
    ws.row_dimensions[r].height=26; return c

def _subtit(ws,r,texto,n):
    ws.merge_cells(start_row=r,start_column=1,end_row=r,end_column=n)
    c=ws.cell(row=r,column=1,value=texto)
    c.font=Font(name="Arial",bold=True,size=10,color="FFFFFF")
    c.fill=PatternFill("solid",start_color=C_AZUL_MED)
    c.alignment=Alignment(horizontal="center",vertical="center")
    ws.row_dimensions[r].height=20; return c

# ── ABA RESUMO ─────────────────────────────────────────────────────────────────
def aba_resumo(ws, conf_m, d_mai, info_ant, info_atu, neg_abr, neg_mai, vc_mai):
    ws.sheet_view.showGridLines=False; N=7
    _titulo(ws,1,"CONFERÊNCIA LMC – LIVRO DE MOVIMENTAÇÃO DE COMBUSTÍVEIS",N,sz=13)
    ws.row_dimensions[1].height=30
    ia=info_ant["info"]; iu=info_atu["info"]
    razao_emp = ia.get('razao','') or iu.get('razao','')
    cnpj_emp  = ia.get('cnpj','')  or iu.get('cnpj','')
    comp_ant_txt = (f"{ia.get('dt_ini','')} a {ia.get('dt_fin','')}"
                    if ia.get('dt_ini') else "Não enviado (apenas competência atual)")
    for i,(a,e) in enumerate([(f"Empresa: {razao_emp}", f"CNPJ: {cnpj_emp}"),
        (f"Competência anterior: {comp_ant_txt}", f"Competência atual: {iu.get('dt_ini','')} a {iu.get('dt_fin','')}"),
        (f"Gerado em: {datetime.now().strftime('%d/%m/%Y %H:%M')}", "")], start=2):
        ws.merge_cells(start_row=i,start_column=1,end_row=i,end_column=4)
        c1=ws.cell(row=i,column=1,value=a); c1.font=Font(name="Arial",bold=(i==2),size=10)
        c1.alignment=Alignment(horizontal="left",vertical="center")
        ws.merge_cells(start_row=i,start_column=5,end_row=i,end_column=N)
        c2=ws.cell(row=i,column=5,value=e); c2.font=Font(name="Arial",size=10)
        c2.alignment=Alignment(horizontal="right",vertical="center")
        ws.row_dimensions[i].height=16

    def bloco_ok(r, texto):
        ws.merge_cells(start_row=r,start_column=1,end_row=r,end_column=N)
        c=ws.cell(row=r,column=1,value=f"✅  {texto}")
        c.font=Font(name="Arial",bold=True,size=10,color=C_VERDE_FG)
        c.fill=PatternFill("solid",start_color=C_VERDE_BG)
        c.alignment=Alignment(horizontal="left",vertical="center"); c.border=_brd()
        ws.row_dimensions[r].height=18; return r+1

    def bloco_err(r, texto):
        ws.merge_cells(start_row=r,start_column=1,end_row=r,end_column=N)
        c=ws.cell(row=r,column=1,value=f"❌  {texto}")
        c.font=Font(name="Arial",bold=True,size=10,color=C_VERM_FG)
        c.fill=PatternFill("solid",start_color=C_VERM_BG)
        c.alignment=Alignment(horizontal="left",vertical="center"); c.border=_brd()
        ws.row_dimensions[r].height=18; return r+1

    def detalhe(r, texto, cor_bg=C_VERM_BG, cor_fg=C_VERM_FG):
        ws.merge_cells(start_row=r,start_column=1,end_row=r,end_column=N)
        c=ws.cell(row=r,column=1,value=f"     ▸  {texto}")
        c.font=Font(name="Arial",size=10,color=cor_fg)
        c.fill=PatternFill("solid",start_color=cor_bg)
        c.alignment=Alignment(horizontal="left",vertical="center"); c.border=_brd()
        ws.row_dimensions[r].height=16; return r+1

    row=6

    # 1. CONFRONTO MENSAL
    _subtit(ws,row,"1.  CONFRONTO ENTRE MESES  –  Fechamento Anterior × Abertura Atual",N); row+=1
    sem_anterior = not conf_m["tanques"] and not conf_m["bicos"]
    if sem_anterior:
        row=bloco_ok(row, "SPED da competência anterior não enviado — confronto entre meses não realizado. "
                          "Conferência abaixo considera apenas a competência atual.")
    else:
        for i,h in enumerate(["Tipo","ID","Data Fech.","Valor Fech.","Data Aber.","Valor Aber.","Status"],1):
            _ch(ws,row,i,h,bg=C_CINZA,fg=C_AZUL_ESC)
        ws.row_dimensions[row].height=22; row+=1
        for tipo, lista in [("Tanque", conf_m["tanques"]), ("Bico", conf_m["bicos"])]:
            for x in lista:
                bg=_row_bg(x["status"])
                _dc(ws,row,1,tipo,bg=bg)
                _dc(ws,row,2,f"{tipo} {x['id']}",bg=bg)
                _dc(ws,row,3,x["dt_fech"].strftime("%d/%m/%Y") if x.get("dt_fech") else "",bg=bg)
                _dc(ws,row,4,x["fech"],NF,bg=bg)
                _dc(ws,row,5,x["dt_aber"].strftime("%d/%m/%Y") if x.get("dt_aber") else "",bg=bg)
                _dc(ws,row,6,x["aber"],NF,bg=bg); _sc(ws,row,7,x["status"])
                ws.row_dimensions[row].height=15; row+=1
    row+=1

    # 2. CONSISTÊNCIA DIÁRIA
    _subtit(ws,row,"2.  CONSISTÊNCIA DIÁRIA – COMPETÊNCIA ATUAL  (Fechamento dia N = Abertura dia N+1)",N); row+=1
    divs_t=[x for x in d_mai["tanques"] if x["status"]=="❌ DIVERGÊNCIA"]
    divs_b=[x for x in d_mai["bicos"]   if x["status"]=="❌ DIVERGÊNCIA"]

    if not divs_t:
        row=bloco_ok(row, f"Tanques: {len(d_mai['tanques'])} transições verificadas — nenhuma divergência encontrada")
    else:
        row=bloco_err(row, f"Tanques: {len(divs_t)} divergência(s) de {len(d_mai['tanques'])} transições")
        for x in divs_t:
            row=detalhe(row, f"Tanque {x['tanque']}  |  Fech. {x['dia_fech'].strftime('%d/%m/%Y')}: {x['fech']:,.3f}  →  Aber. {x['dia_aber'].strftime('%d/%m/%Y')}: {x['aber']:,.3f}  |  Dif.: {x['dif']:,.3f} L")

    if not divs_b:
        row=bloco_ok(row, f"Bicos: {len(d_mai['bicos'])} transições verificadas — nenhuma divergência encontrada")
    else:
        row=bloco_err(row, f"Bicos: {len(divs_b)} divergência(s) de {len(d_mai['bicos'])} transições")
        for x in divs_b:
            row=detalhe(row, f"Bico {x['bico']}  |  Fech. {x['dia_fech'].strftime('%d/%m/%Y')}: {x['fech']:,.3f}  →  Aber. {x['dia_aber'].strftime('%d/%m/%Y')}: {x['aber']:,.3f}  |  Dif.: {x['dif']:,.3f}")
    row+=1

    # 3. VALORES NEGATIVOS
    _subtit(ws,row,"3.  VALORES NEGATIVOS NOS REGISTROS 1310 / 1320  (Ambas as competências)",N); row+=1
    todos_neg_t = neg_abr["tanques"] + neg_mai["tanques"]
    todos_neg_b = neg_abr["bicos"]   + neg_mai["bicos"]

    if not todos_neg_t:
        row=bloco_ok(row, "Tanques (Reg. 1310): nenhum valor negativo detectado")
    else:
        row=bloco_err(row, f"Tanques: {len(todos_neg_t)} valor(es) negativo(s) encontrado(s)")
        for x in todos_neg_t:
            row=detalhe(row, f"Tanque {x['tanque']}  |  Data: {x['data'].strftime('%d/%m/%Y') if x.get('data') else 'N/D'}  |  Campo: {x['campo']}  |  Valor: {x['valor']:,.3f}  |  Linha SPED: {x['linha']}")
            diag=DIAGNOSTICO_CAMPO.get(x['campo'],"Valor negativo inesperado neste campo")
            row=detalhe(row, f"     ℹ️  {diag}", cor_bg="FCE4D6", cor_fg="7D4200")

    if not todos_neg_b:
        row=bloco_ok(row, "Bicos (Reg. 1320): nenhum valor negativo detectado")
    else:
        row=bloco_err(row, f"Bicos: {len(todos_neg_b)} valor(es) negativo(s) encontrado(s)")
        for x in todos_neg_b:
            row=detalhe(row, f"Bico {x['bico']}  |  Data: {x['data'].strftime('%d/%m/%Y') if x.get('data') else 'N/D'}  |  Campo: {x['campo']}  |  Valor: {x['valor']:,.3f}  |  Linha SPED: {x['linha']}")
            diag=DIAGNOSTICO_CAMPO.get(x['campo'],"Valor negativo inesperado neste campo")
            row=detalhe(row, f"     ℹ️  {diag}", cor_bg="FCE4D6", cor_fg="7D4200")
    row+=1

    # 4. VERSÃO E CAPACIDADE
    _subtit(ws,row,f"4.  VERSÃO DO SPED E CAPACIDADE DOS TANQUES  –  {vc_mai['periodo']}",N); row+=1
    versao_txt = (f"✅  Versão do SPED: {vc_mai['versao']} — correta (obrigatório: {VERSAO_OBRIGATORIA})"
                  if vc_mai["versao_ok"]
                  else f"❌  Versão do SPED: {vc_mai['versao']} — incorreta! Obrigatório: {VERSAO_OBRIGATORIA}. Pode causar rejeição no validador.")
    ws.merge_cells(start_row=row,start_column=1,end_row=row,end_column=N)
    c=ws.cell(row=row,column=1,value=versao_txt)
    bg=C_VERDE_BG if vc_mai["versao_ok"] else C_VERM_BG
    fg=C_VERDE_FG if vc_mai["versao_ok"] else C_VERM_FG
    c.font=Font(name="Arial",bold=True,size=10,color=fg)
    c.fill=PatternFill("solid",start_color=bg)
    c.alignment=Alignment(horizontal="left",vertical="center"); c.border=_brd()
    ws.row_dimensions[row].height=18; row+=1

    n_cap_ok  = sum(1 for t in vc_mai["tanques"] if t["status"]=="✅ OK")
    n_cap_err = sum(1 for t in vc_mai["tanques"] if t["status"]!="✅ OK")
    if n_cap_err==0:
        row=bloco_ok(row, f"Capacidade dos tanques: todos os {n_cap_ok} tanques declarados corretamente  |  "
                     + "  ".join(f"Tanque {t['tanque']}: {t['obs']}" for t in vc_mai["tanques"]))
    else:
        row=bloco_err(row, f"Capacidade: {n_cap_err} tanque(s) com problema")
        for t in vc_mai["tanques"]:
            bg2=_row_bg(t["status"]) or C_AMAR_BG
            fg2=C_VERM_FG if "❌" in t["status"] else C_AMAR_FG
            row=detalhe(row, f"Tanque {t['tanque']}  |  {t['status']}  |  {t['obs']}", cor_bg=bg2, cor_fg=fg2)

    ws.column_dimensions["A"].width=90
    for i in range(2,N+1):
        ws.column_dimensions[get_column_letter(i)].width=0.1

# ── ABA CONFRONTO MENSAL ───────────────────────────────────────────────────────
def aba_mensal(ws, conf_m, info_ant, info_atu):
    ws.sheet_view.showGridLines = False
    r = 1

    def fmt_comp(dt):
        meses = ['Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set','Out','Nov','Dez']
        try:
            mes = int(dt[2:4]); ano = dt[4:]
            return f"{meses[mes-1]}/{ano}"
        except:
            return dt

    comp_ant = fmt_comp(info_ant["info"].get("dt_fin", "")) or "—"
    comp_atu = fmt_comp(info_atu["info"].get("dt_fin", ""))

    if not conf_m["tanques"] and not conf_m["bicos"]:
        _titulo(ws, r, "CONFRONTO ENTRE MESES", 6, sz=13); r += 2
        cel = ws.cell(row=r, column=1,
            value="⚠️  SPED da competência anterior não foi enviado. "
                  "Não é possível confrontar o fechamento do mês anterior com a abertura "
                  "do mês atual. Consulte as demais abas para a conferência da competência atual.")
        cel.font = Font(name="Arial", size=11, bold=True, color=C_AMAR_FG)
        cel.fill = PatternFill("solid", start_color=C_AMAR_BG)
        cel.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        ws.merge_cells(start_row=r, start_column=1, end_row=r+2, end_column=6)
        ws.row_dimensions[r].height=60
        ws.column_dimensions["A"].width=90
        return

    for secao, lista, tipo, lbl_fech_ant, lbl_aber_atu, lbl_fech_atu in [
        ("TANQUES – Confronto Competência Anterior × Atual (Reg. 1310)",
         conf_m["tanques"], "Tanque",
         f"Est. Fechamento\n{comp_ant} (L)",
         f"Est. Abertura\n{comp_atu} (L)",
         f"Est. Fechamento\n{comp_atu} (L)"),
        ("BICOS – Confronto Competência Anterior × Atual (Reg. 1320)",
         conf_m["bicos"], "Bico",
         f"Enc. Fechamento\n{comp_ant}",
         f"Enc. Abertura\n{comp_atu}",
         f"Enc. Fechamento\n{comp_atu}"),
    ]:
        _titulo(ws, r, secao, 6); r += 1
        for i, h in enumerate([tipo, lbl_fech_ant, lbl_aber_atu, "Diferença", "Status", lbl_fech_atu], 1):
            _ch(ws, r, i, h)
        ws.row_dimensions[r].height = 34; r += 1

        for x in lista:
            bg = _row_bg(x["status"])
            _dc(ws, r, 1, f"{tipo} {x['id']}", bg=bg)
            _dc(ws, r, 2, x["fech"],     NF, bg=bg)
            _dc(ws, r, 3, x["aber"],     NF, bg=bg)
            _dc(ws, r, 4, x["dif"],      NF, bg=bg)
            _sc(ws, r, 5, x["status"])
            _dc(ws, r, 6, x.get("fech_atu"), NF, bg=bg)
            ws.row_dimensions[r].height = 15; r += 1
        r += 2

    for i, w in enumerate([14, 22, 22, 16, 22, 22], 1):
        ws.column_dimensions[get_column_letter(i)].width = w

def aba_diario(ws, d_mai):
    ws.sheet_view.showGridLines=False; r=1
    for secao,lista,id_f,cols in [
        ("TANQUES – Comparativo Diário Competência Atual (Reg. 1310)",
         d_mai["tanques"],"tanque",
         ["Tanque","Data Fech.","Est. Fechamento (L)","Data Aber.","Est. Abertura (L)","Diferença (L)","Status"]),
        ("BICOS – Comparativo Diário Competência Atual (Reg. 1320)",
         d_mai["bicos"],"bico",
         ["Bico","Data Fech.","Enc. Fechamento","Data Aber.","Enc. Abertura","Diferença","Status"]),
    ]:
        _titulo(ws,r,secao,8); r+=1
        for i,h in enumerate(cols,1): _ch(ws,r,i,h)
        ws.row_dimensions[r].height=30; r+=1
        tipo=cols[0]  # "Tanque" ou "Bico"
        for x in lista:
            bg=_row_bg(x["status"])
            _dc(ws,r,1,f"{tipo} {x[id_f]}",bg=bg)
            _dc(ws,r,2,x["dia_fech"].strftime("%d/%m/%Y"),bg=bg)
            _dc(ws,r,3,x["fech"],NF,bg=bg)
            _dc(ws,r,4,x["dia_aber"].strftime("%d/%m/%Y"),bg=bg)
            _dc(ws,r,5,x["aber"],NF,bg=bg)
            _dc(ws,r,6,x["dif"],NF,bg=bg)
            _sc(ws,r,7,x["status"])
            ws.row_dimensions[r].height=15; r+=1
        r+=2
    for i,w in enumerate([14,14,22,14,22,16,22],1):
        ws.column_dimensions[get_column_letter(i)].width=w

# ══════════════════════════════════════════════════════════════════════════════
# MAIN

def ler_sped_bytes(data):
    """Versão para receber bytes (upload web) ao invés de caminho de arquivo."""
    text = data.decode("latin-1", errors="replace")
    tanques={}; bicos={}; info={}
    data_atual=None; vals_1300={}

    for linha in text.splitlines():
        c=linha.strip().split("|")
        if len(c)<2: continue
        tp=c[1]

        if tp=="0000":
            info={"versao":c[2].strip() if len(c)>2 else "",
                  "razao": c[6].strip() if len(c)>6 else "",
                  "cnpj":  c[7].strip() if len(c)>7 else "",
                  "dt_ini":c[4].strip() if len(c)>4 else "",
                  "dt_fin":c[5].strip() if len(c)>5 else ""}

        elif tp=="1300":
            data_atual=_dt(c[3]) if len(c)>3 else None
            if not data_atual: continue
            vals_1300={
                "data":data_atual,
                "est_abert":_fl(c[4]),"entrada":_fl(c[5]),"saida":_fl(c[7]),
                "evap":_fl(c[9]) if len(c)>9 else 0.0,
                "ajuste":_fl(c[10]) if len(c)>10 else 0.0,
                "est_fech":_fl(c[11]) if len(c)>11 else 0.0,
            }

        elif tp=="1310":
            if not vals_1300: continue
            t=_nid(c[2])
            tem_campos_proprios = len(c) > 10 and c[3].strip() != ""
            if tem_campos_proprios:
                try:
                    est_abert_1310 = _fl(c[3])
                    entrada_1310   = _fl(c[4]) if len(c)>4 else 0.0
                    saida_1310     = _fl(c[6]) if len(c)>6 else 0.0
                    evap_1310      = _fl(c[8]) if len(c)>8 else 0.0
                    ajuste_1310    = _fl(c[9]) if len(c)>9 else 0.0
                    est_fech_1310  = _fl(c[10]) if len(c)>10 else 0.0
                    cap            = _fl(c[11]) if len(c)>11 and c[11].strip() else None
                except Exception:
                    tem_campos_proprios = False

            if tem_campos_proprios:
                key=(t,vals_1300["data"])
                tanques[key]={
                    "tanque":t,"data":vals_1300["data"],
                    "est_abert":est_abert_1310,"entrada":entrada_1310,
                    "saida":saida_1310,"evap":evap_1310,
                    "ajuste":ajuste_1310,"est_fech":est_fech_1310,
                    "capacidade":cap,
                }
            else:
                cap=_fl(c[11]) if len(c)>11 and c[11].strip() else None
                key=(t,vals_1300["data"])
                tanques[key]={
                    "tanque":t,"data":vals_1300["data"],
                    "est_abert":vals_1300["est_abert"],"entrada":vals_1300["entrada"],
                    "saida":vals_1300["saida"],"evap":vals_1300["evap"],
                    "ajuste":vals_1300["ajuste"],"est_fech":vals_1300["est_fech"],
                    "capacidade":cap,
                }

        elif tp=="1320":
            b=_nid(c[2])
            if not data_atual: continue
            bicos[(b,data_atual)]={
                "bico":b,"data":data_atual,
                "enc_abert":_fl(c[9]) if len(c)>9 else 0.0,
                "enc_fech": _fl(c[8]) if len(c)>8 else 0.0,
            }

    return {"info":info,"tanques":tanques,"bicos":bicos}


def verificar_negativos_bytes(data):
    """Versão para receber bytes ao invés de caminho de arquivo."""
    text = data.decode("latin-1", errors="replace")
    neg_t=[]; neg_b=[]; da=None
    for n, linha in enumerate(text.splitlines(), 1):
        c = linha.strip().split("|")
        if len(c)<2: continue
        tp=c[1]
        if tp=="1300":
            da=_dt(c[3]) if len(c)>3 else None
            tanque=_nid(c[2])
            for idx,nome in CAMPOS_1300.items():
                if idx>=len(c): continue
                try:
                    v=_fl(c[idx])
                    if v<0: neg_t.append({"tanque":tanque,"data":da,"campo":nome,"valor":v,"linha":n})
                except: pass
        elif tp=="1320":
            bico=_nid(c[2])
            for idx,nome in CAMPOS_1320.items():
                if idx>=len(c): continue
                try:
                    v=_fl(c[idx])
                    if v<0: neg_b.append({"bico":bico,"data":da,"campo":nome,"valor":v,"linha":n})
                except: pass
    return {"tanques": neg_t, "bicos": neg_b}

# Diagnóstico por campo negativo
DIAGNOSTICO_CAMPO = {
    "Est. Abertura":          "Estoque de abertura negativo — erro de lançamento ou fechamento anterior incorreto",
    "Entrada":                "Entrada de combustível negativa — possível estorno de nota fiscal ou lançamento incorreto",
    "Est. Aber. Pós Entrada": "Estoque após entrada negativo — inconsistência entre abertura e entrada lançadas",
    "Saída":                  "Saída negativa — possível estorno ou correção de venda lançada indevidamente",
    "Est. Fech. Pré Ajuste":  "Estoque pré-ajuste negativo — vendas ou perdas superiores ao estoque disponível",
    "Evaporação":             "Evaporação negativa — valor inválido; evaporação deve ser sempre positiva ou zero",
    "Ajuste":                 "Ajuste negativo — perda de inventário ou correção de estoque para menor",
    "Est. Fechamento Final":  "Estoque de fechamento negativo — saldo final abaixo de zero, impossível fisicamente",
    "Enc. Fechamento":        "Encerrante de fechamento negativo — erro de leitura do bico ou lançamento incorreto",
    "Enc. Abertura":          "Encerrante de abertura negativo — valor inválido; encerrante é sempre crescente",
    "Volume Vendido":         "Volume vendido negativo — possível estorno ou erro de registro no bico",
    "Diferença Encerrante":   "Diferença de encerrante negativa — bico apresentou recuo, possível adulteração ou falha",
}


def aba_dac(ws, conf_dac, info_atu):
    """Aba de confronto DAC × SPED da competência atual."""
    ws.sheet_view.showGridLines = False
    r = 1

    def fmt_comp(dt):
        meses = ['Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set','Out','Nov','Dez']
        try: return f"{meses[int(dt[2:4])-1]}/{dt[4:]}"
        except: return dt

    comp_atu = fmt_comp(info_atu["info"].get("dt_fin",""))
    comp_dac = conf_dac.get("competencia", comp_atu)

    _titulo(ws, r, f"CONFRONTO DAC × SPED  –  Competência {comp_atu}", 8, sz=13)
    ws.row_dimensions[r].height = 30; r += 1

    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=8)
    c = ws.cell(row=r, column=1,
        value=f"DAC competência: {comp_dac}  |  SPED competência: {comp_atu}  |  "
              f"Gerado em: {datetime.now().strftime('%d/%m/%Y %H:%M')}")
    c.font = Font(name="Arial", size=9, italic=True, color="595959")
    c.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[r].height = 16; r += 2

    # ── TANQUES ──────────────────────────────────────────────────────────────
    _subtit(ws, r, f"TANQUES (Reg. 1310)  –  Estoque Inicial e Final {comp_atu}", 8); r += 1
    for i, h in enumerate([
        "Tanque", "Produto",
        "Est. Inicial\nDAC (L)", "Est. Inicial\nSPED (L)", "Dif. Inicial",
        "Est. Final\nDAC (L)",   "Est. Final\nSPED (L)",   "Dif. Final",
    ], 1): _ch(ws, r, i, h, bg=C_CINZA, fg=C_AZUL_ESC)
    ws.row_dimensions[r].height = 30; r += 1

    for x in conf_dac.get("tanques", []):
        st_i = x["status_ini"]; st_f = x["status_fin"]
        bg = (C_VERM_BG if "DIVERGÊNCIA" in (st_i+st_f) else
              C_AMAR_BG if "AUSENTE"     in (st_i+st_f) else C_VERDE_BG)
        _dc(ws,r,1, f"Tanque {x['id']}", bg=bg)
        _dc(ws,r,2, x.get("produto",""), bg=bg)
        _dc(ws,r,3, x["ei_dac"],  NF, bg=bg)
        _dc(ws,r,4, x["ei_sped"], NF, bg=bg)
        _dc(ws,r,5, x["dif_ini"], NF, bg=bg)
        _dc(ws,r,6, x["ef_dac"],  NF, bg=bg)
        _dc(ws,r,7, x["ef_sped"], NF, bg=bg)
        _dc(ws,r,8, x["dif_fin"], NF, bg=bg)
        ws.row_dimensions[r].height = 15; r += 1
    r += 1

    # ── BICOS ─────────────────────────────────────────────────────────────────
    _subtit(ws, r, f"BICOS (Reg. 1320)  –  Encerrante Inicial e Final {comp_atu}", 7); r += 1
    for i, h in enumerate([
        "Bico",
        "Enc. Inicial\nDAC", "Enc. Inicial\nSPED", "Dif. Inicial",
        "Enc. Final\nDAC",   "Enc. Final\nSPED",   "Dif. Final",
    ], 1): _ch(ws, r, i, h, bg=C_CINZA, fg=C_AZUL_ESC)
    ws.row_dimensions[r].height = 30; r += 1

    for x in conf_dac.get("bicos", []):
        st_i = x["status_ini"]; st_f = x["status_fin"]
        bg = (C_VERM_BG if "DIVERGÊNCIA" in (st_i+st_f) else
              C_AMAR_BG if "AUSENTE"     in (st_i+st_f) else C_VERDE_BG)
        _dc(ws,r,1, f"Bico {x['id']}", bg=bg)
        _dc(ws,r,2, x["ei_dac"],  NF, bg=bg)
        _dc(ws,r,3, x["ei_sped"], NF, bg=bg)
        _dc(ws,r,4, x["dif_ini"], NF, bg=bg)
        _dc(ws,r,5, x["ef_dac"],  NF, bg=bg)
        _dc(ws,r,6, x["ef_sped"], NF, bg=bg)
        _dc(ws,r,7, x["dif_fin"], NF, bg=bg)
        ws.row_dimensions[r].height = 15; r += 1

    for i, w in enumerate([12, 22, 18, 18, 16, 18, 18, 16], 1):
        ws.column_dimensions[get_column_letter(i)].width = w


def aba_dac_sped(ws, d_atu, info_atu):
    """Gera um DAC de acompanhamento a partir dos dados do SPED da competência atual."""
    ws.sheet_view.showGridLines = False
    r = 1
    N = 9

    def fmt_comp(dt):
        meses=['Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set','Out','Nov','Dez']
        try: return f"{meses[int(dt[2:4])-1]}/{dt[4:]}"
        except: return dt

    info  = info_atu["info"]
    comp  = fmt_comp(info.get("dt_fin",""))
    razao = info.get("razao","")
    cnpj  = info.get("cnpj","")
    dt_ini= info.get("dt_ini","")
    dt_fin= info.get("dt_fin","")

    _titulo(ws, r, "DAC – DOCUMENTO DE ACOMPANHAMENTO DE COMBUSTÍVEIS (gerado pelo SPED)", N, sz=12)
    ws.row_dimensions[r].height = 28; r += 1

    for texto_a, texto_e in [
        (f"Empresa: {razao}", f"CNPJ: {cnpj}"),
        (f"Período: {dt_ini[:2]}/{dt_ini[2:4]}/{dt_ini[4:]} a {dt_fin[:2]}/{dt_fin[2:4]}/{dt_fin[4:]}", "Fonte: SPED Fiscal – Reg. 1300 / 1320"),
        (f"Gerado em: {datetime.now().strftime('%d/%m/%Y %H:%M')}", "⚠️ Valores baseados no escritural do SPED (sem medição física)"),
    ]:
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
        c1 = ws.cell(row=r, column=1, value=texto_a)
        c1.font = Font(name="Arial", size=10, bold=True)
        c1.alignment = Alignment(horizontal="left", vertical="center")
        ws.merge_cells(start_row=r, start_column=6, end_row=r, end_column=N)
        c2 = ws.cell(row=r, column=6, value=texto_e)
        c2.font = Font(name="Arial", size=10, italic=True, color="595959")
        c2.alignment = Alignment(horizontal="right", vertical="center")
        ws.row_dimensions[r].height = 16; r += 1
    r += 1

    # ── POSIÇÃO DOS TANQUES ───────────────────────────────────────────────────
    _subtit(ws, r, f"POSIÇÃO DOS TANQUES  –  {comp}", N); r += 1
    hdrs_t = ["Tanque","Produto","Est. Inicial (L)","Recebimento (L)","Venda (L)",
              "Evaporação (L)","Perda / Ganho (L)","Est. Final (L)","Variação (L)"]
    for i,h in enumerate(hdrs_t,1): _ch(ws,r,i,h)
    ws.row_dimensions[r].height = 28; r += 1

    for t in sorted(set(tt for (tt,_) in d_atu["tanques"]), key=_sk):
        dias = sorted(dt for (tt,dt) in d_atu["tanques"] if tt==t)
        est_ini  = d_atu["tanques"][(t,dias[0])]["est_abert"]
        est_fin  = d_atu["tanques"][(t,dias[-1])]["est_fech"]
        total_rec= round(sum(d_atu["tanques"][(t,dt)]["entrada"] for dt in dias),3)
        total_sai= round(sum(d_atu["tanques"][(t,dt)]["saida"]   for dt in dias),3)
        total_evap=round(sum(d_atu["tanques"][(t,dt)]["evap"]    for dt in dias),3)
        total_aj = round(sum(d_atu["tanques"][(t,dt)]["ajuste"]  for dt in dias),3)
        variacao = round(est_fin - est_ini, 3)

        cap = d_atu["tanques"].get((t,dias[0]),{}).get("capacidade")
        produto = f"{cap:,.0f} L" if cap else "—"

        if total_aj < 0:
            bg_aj = C_VERM_BG; fg_aj = C_VERM_FG
        elif total_aj > 0:
            bg_aj = C_VERDE_BG; fg_aj = C_VERDE_FG
        else:
            bg_aj = None; fg_aj = "000000"

        _dc(ws,r,1,f"Tanque {t}")
        _dc(ws,r,2,produto)
        _dc(ws,r,3,est_ini,   NF)
        _dc(ws,r,4,total_rec, NF, bg="D6E4F0" if total_rec>0 else None)
        _dc(ws,r,5,total_sai, NF)
        _dc(ws,r,6,total_evap,NF)
        cel_aj = ws.cell(row=r,column=7,value=total_aj)
        cel_aj.number_format = NF
        cel_aj.font = Font(name="Arial",size=10,bold=(total_aj!=0),color=fg_aj)
        cel_aj.alignment = Alignment(horizontal="center",vertical="center")
        cel_aj.border = _brd()
        if bg_aj: cel_aj.fill = PatternFill("solid",start_color=bg_aj)
        _dc(ws,r,8,est_fin,   NF)
        cel_v = ws.cell(row=r,column=9,value=variacao)
        cel_v.number_format = NF
        cel_v.font = Font(name="Arial",size=10,
                          color=C_VERDE_FG if variacao>=0 else C_VERM_FG,bold=True)
        cel_v.alignment = Alignment(horizontal="center",vertical="center")
        cel_v.border = _brd()
        cel_v.fill = PatternFill("solid", start_color=C_VERDE_BG if variacao>=0 else C_VERM_BG)
        ws.row_dimensions[r].height = 16; r += 1
    r += 2

    # ── POSIÇÃO DOS BICOS ─────────────────────────────────────────────────────
    _subtit(ws, r, f"POSIÇÃO DOS BICOS  –  {comp}", N); r += 1
    hdrs_b = ["Bico","Enc. Inicial","Enc. Final","Litros Vendidos"]
    for i,h in enumerate(hdrs_b,1): _ch(ws,r,i,h,bg=C_AZUL_ESC)
    ws.row_dimensions[r].height = 24; r += 1

    total_litros = 0
    for b in sorted(set(bb for (bb,_) in d_atu["bicos"]), key=_sk):
        dias = sorted(dt for (bb,dt) in d_atu["bicos"] if bb==b)
        enc_ini = d_atu["bicos"][(b,dias[0])]["enc_abert"]
        enc_fin = d_atu["bicos"][(b,dias[-1])]["enc_fech"]
        litros  = round(enc_fin - enc_ini, 3)
        total_litros += litros

        bg = C_VERDE_BG if litros>0 else C_AMAR_BG
        _dc(ws,r,1,f"Bico {b}")
        _dc(ws,r,2,enc_ini,NF)
        _dc(ws,r,3,enc_fin,NF)
        _dc(ws,r,4,litros, NF, bg=bg)
        ws.row_dimensions[r].height = 15; r += 1

    for i,v in enumerate(["TOTAL","","",round(total_litros,3)],1):
        c=ws.cell(row=r,column=i,value=v)
        c.font=Font(name="Arial",bold=True,size=10)
        c.fill=PatternFill("solid",start_color=C_CINZA)
        c.alignment=Alignment(horizontal="center",vertical="center")
        c.border=_brd()
        if i==4: c.number_format=NF
    ws.row_dimensions[r].height=18; r+=1

    for i,w in enumerate([12,14,18,18,18,16,14,18,16],1):
        ws.column_dimensions[get_column_letter(i)].width=w
