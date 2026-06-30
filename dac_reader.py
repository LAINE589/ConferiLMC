"""
Módulo de leitura do DAC — sem dependência de API externa.
Suporta:
  - Excel (.xlsx, .xls): leitura direta via openpyxl/xlrd
  - PDF com texto extraível: parse automático (formato AutoSystem PRO e similares)
  - PDF escaneado / imagem: retorna None (sistema gera sem confronto DAC)
"""
import io, re, json

def _fl(v):
    """Converte para float. Aceita number nativo (int/float, ex: do openpyxl)
    OU string no formato brasileiro: '1.234,56' -> 1234.56"""
    if v is None: return None
    if isinstance(v, (int, float)):
        return float(v)
    v = str(v).strip().replace(".", "").replace(",", ".")
    try: return float(v)
    except: return None

def _nid(v):
    """Normaliza ID: '01' → '1', '003' → '3'"""
    try: return str(int(str(v).strip()))
    except: return str(v).strip()


# ── EXCEL ─────────────────────────────────────────────────────────────────────
def _ler_excel(arquivo_bytes, filename):
    import openpyxl
    wb = openpyxl.load_workbook(io.BytesIO(arquivo_bytes), data_only=True)
    # Concatenar todas as abas em texto tabular
    linhas = []
    for ws in wb.worksheets:
        for row in ws.iter_rows(values_only=True):
            vals = [str(v).strip() if v is not None else "" for v in row]
            linha = " | ".join(vals).strip(" |")
            if linha:
                linhas.append(linha)
    texto = "\n".join(linhas)
    return _parse_texto(texto)


# ── PDF ───────────────────────────────────────────────────────────────────────
def _ler_pdf(arquivo_bytes):
    try:
        import pypdf
        reader = pypdf.PdfReader(io.BytesIO(arquivo_bytes))
        texto = ""
        for page in reader.pages:
            texto += (page.extract_text() or "") + "\n"
        texto = texto.strip()
        if len(texto) < 50:
            return None  # PDF escaneado/imagem — sem texto extraível

        # Tentar parser "Resumo DAC" primeiro (formato com 1 linha por bico/tanque)
        if re.search(r'Resumo\s*DAC', texto, re.I):
            resultado = _parse_resumo_dac(texto)
            if resultado["tanques"] or resultado["bicos"]:
                return resultado

        # Fallback: parser AutoSystem PRO (multi-linha por campo)
        resultado = _parse_texto(texto)
        if resultado["tanques"] or resultado["bicos"]:
            return resultado

        return None
    except Exception:
        return None


# ── PARSER FORMATO "RESUMO DAC" (uma linha por bico/tanque) ──────────────────
def _parse_resumo_dac(texto):
    """
    Parser para o formato 'Resumo DAC' onde cada bico/tanque está em UMA linha:
      Bomba Bico Combustível Enc.Inicial Enc.Final Vol.semInt. Aferição Vol.comInt.
      Tanque Combustível Est.Abertura Est.Fechamento Vol.Recebido Faltas/Sobras
    """
    tanques = []
    bicos   = []
    competencia = ""

    # Competência
    m = re.search(r'[Pp]er[íi]odo\s+de\s+apura[çc][ãa]o[:\s]+(\d{2}/\d{2}/\d{4})\s+at[ée]\s+(\d{2}/\d{2}/\d{4})', texto)
    if m:
        dt_fin = m.group(2)
        meses = ['Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set','Out','Nov','Dez']
        try:
            mes = int(dt_fin[3:5]); ano = dt_fin[6:]
            competencia = f"{meses[mes-1]}/{ano}"
        except:
            competencia = dt_fin

    num = r'-?[\d\.]+,\d+'  # número brasileiro: 1.234,567 ou -123,456

    # ── BICOS: "001 001 GASOLINA COMUM 1.655.772,200 1.681.591,360 0,000 25.819,160 1,211"
    # Padrão: bomba(3) bico(3) PRODUTO(palavras) num num num num num
    bico_pat = re.compile(
        r'^(\d{2,3})\s+(\d{2,3})\s+([A-ZÀ-Ú0-9 ]+?)\s+(' + num + r')\s+(' + num + r')\s+(' + num + r')\s+(' + num + r')(?:\s+(' + num + r'))?\s*$'
    )
    for linha in texto.splitlines():
        l = linha.strip()
        m = bico_pat.match(l)
        if m:
            bico_id = _nid(m.group(2))
            enc_ini = _fl(m.group(4))
            enc_fin = _fl(m.group(5))
            if enc_ini is not None and enc_fin is not None and enc_fin >= enc_ini:
                # Evita capturar linhas de bombas (modelo/série) por engano:
                # só aceita se já vimos ENCERRANTE crescente (bico real)
                ids_existentes = [b['id'] for b in bicos]
                if bico_id not in ids_existentes:
                    bicos.append({
                        "id": bico_id,
                        "encerrante_inicial": enc_ini,
                        "encerrante_final":   enc_fin,
                    })

    # ── TANQUES: "001 GASOLINA COMUM 10.840,862 5.618,234 149.000,000 -31.095,102"
    # Padrão real extraído: "001 10.840,862GASOLINA COMUM 5.618,234 149.000,000 -31.095,102"
    # (PyPDF às vezes gruda o número com o texto seguinte sem espaço)
    tanque_pat = re.compile(
        r'^(\d{3})\s+(' + num + r')\s*([A-ZÀ-Ú][A-ZÀ-Ú0-9 ]*?)\s+(' + num + r')\s+(' + num + r')\s+(' + num + r')\s*$'
    )
    em_estoque = False
    for linha in texto.splitlines():
        l = linha.strip()
        if re.search(r'Estoque\s*F[íi]sico', l, re.I):
            em_estoque = True; continue
        if re.search(r'^Resumo\s*DAC\s*-', l, re.I):
            em_estoque = False; continue
        if not em_estoque:
            continue
        m = tanque_pat.match(l)
        if m:
            tid     = _nid(m.group(1))
            ini     = _fl(m.group(2))
            produto = m.group(3).strip()
            fin     = _fl(m.group(4))
            ids_existentes = [t['id'] for t in tanques]
            if tid not in ids_existentes:
                tanques.append({
                    "id": tid,
                    "produto": produto,
                    "estoque_inicial": ini,
                    "estoque_final":   fin,
                })

    return {"competencia": competencia, "tanques": tanques, "bicos": bicos}


# ── PARSER UNIVERSAL ──────────────────────────────────────────────────────────
def _parse_texto(texto):
    """
    Parser para o formato AutoSystem PRO / Linx e similares.
    Extrai tanques e bicos da seção POSIÇÃO DOS TANQUES e BICO E ENCERRANTES.
    """
    tanques = []
    bicos   = []
    competencia = ""

    linhas = texto.splitlines()

    # Detectar período/competência
    for l in linhas:
        m = re.search(r'[Pp]er[íi]odo[:\s]+(\d{2}/\d{2}/\d{4})\s+a\s+(\d{2}/\d{2}/\d{4})', l)
        if m:
            dt_fin = m.group(2)  # "30/04/2026"
            meses = ['Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set','Out','Nov','Dez']
            try:
                mes = int(dt_fin[3:5]); ano = dt_fin[6:]
                competencia = f"{meses[mes-1]}/{ano}"
            except:
                competencia = dt_fin
            break

    # ── Seção BICO E ENCERRANTES ─────────────────────────────────────────────
    # Formato: "01 GC\n91.387,190\n108.958,710\n..." (uma linha por campo)
    # Ou: "01 GC | 91.387,190 | 108.958,710"
    em_bicos = False
    bico_buffer = []  # acumula tokens de um bico

    for i, l in enumerate(linhas):
        l = l.strip()
        if re.search(r'BICO\s*E\s*ENCERRANTE', l, re.I):
            em_bicos = True; continue
        if re.search(r'POSI[ÇC][ÃA]O\s*DOS\s*(COMBUST|TANQUE)', l, re.I):
            em_bicos = False; continue
        if not em_bicos: continue
        if re.search(r'^(Bico|Inicial|Final|Litros|Afer|Venda|Pre[çc]|Desc|Valor|Total)', l, re.I):
            continue

        # Linha de bico: começa com número seguido de letras (ex: "01 GC", "02 OGC")
        m = re.match(r'^(\d{1,2})\s+[A-Z]{1,4}$', l)
        if m:
            bico_buffer = [m.group(1)]  # novo bico
            continue

        # Acumular valores numéricos para o bico atual
        if bico_buffer and re.match(r'^[\d.,]+$', l.replace(".", "").replace(",", "")):
            bico_buffer.append(l)
            if len(bico_buffer) >= 3:  # id + inicial + final
                bid = _nid(bico_buffer[0])
                enc_ini = _fl(bico_buffer[1])
                enc_fin = _fl(bico_buffer[2])
                if enc_ini is not None and enc_fin is not None:
                    bicos.append({
                        "id": bid,
                        "encerrante_inicial": enc_ini,
                        "encerrante_final":   enc_fin,
                    })
                bico_buffer = []

    # ── Seção POSIÇÃO DOS TANQUES ─────────────────────────────────────────────
    # Formato: "TANQUE 001 /// GAS\n7.980,83\n79.000,00\n...\n6.851,10\n..."
    em_tanques = False
    tanq_buffer = []

    for l in linhas:
        l = l.strip()
        if re.search(r'POSI[ÇC][ÃA]O\s*DOS\s*TANQUE', l, re.I):
            em_tanques = True; continue
        if re.search(r'^\*Entrada', l, re.I):
            em_tanques = False; continue
        if not em_tanques: continue
        if re.search(r'^(Produto|Início|Entrada|Venda|Afer|Final|Medi|Difer|Total)', l, re.I):
            continue

        # Linha de tanque: "TANQUE 001 /// GAS" ou "TANQUE001///GAS"
        m = re.match(r'TANQUE\s*(\d+)\s*[/\\|]+\s*(\w+)', l, re.I)
        if m:
            tanq_buffer = [m.group(1), m.group(2)]
            continue

        if tanq_buffer and re.match(r'^-?[\d.,]+$', l.replace(".", "").replace(",","")):
            tanq_buffer.append(l)
            # Precisamos de pelo menos: id, produto, inicio, entrada, venda, afer, final
            if len(tanq_buffer) >= 7:
                tid     = _nid(tanq_buffer[0])
                produto = tanq_buffer[1]
                ini     = _fl(tanq_buffer[2])
                fin_raw = _fl(tanq_buffer[6])  # índice 6 = Final (antes de Medição)
                if ini is not None and fin_raw is not None:
                    # Verificar se já existe esse tanque (evitar duplicatas)
                    ids_existentes = [t['id'] for t in tanques]
                    if tid not in ids_existentes:
                        tanques.append({
                            "id":      tid,
                            "produto": produto,
                            "estoque_inicial": ini,
                            "estoque_final":   fin_raw,
                        })
                tanq_buffer = []

    # Fallback: se POSIÇÃO DOS TANQUES não funcionou, tentar POSIÇÃO DOS COMBUSTÍVEIS
    if not tanques:
        em_comb = False
        for l in linhas:
            l = l.strip()
            if re.search(r'POSI[ÇC][ÃA]O\s*DOS\s*COMBUST', l, re.I):
                em_comb = True; continue
            if re.search(r'POSI[ÇC][ÃA]O\s*DOS\s*TANQUE', l, re.I):
                em_comb = False; continue
            if not em_comb: continue
            if re.search(r'^(Produto|Início|Entrada|Venda|Afer|Final|Total|\*)', l, re.I):
                continue
            # Linha de combustível com nome + valores
            partes = [p.strip() for p in re.split(r'\s{2,}|\t', l) if p.strip()]
            if len(partes) >= 7 and not partes[0].isdigit():
                produto = partes[0]
                ini = _fl(partes[1])
                fin = _fl(partes[5]) if len(partes)>5 else None
                if ini and fin:
                    tanques.append({
                        "id": str(len(tanques)+1),
                        "produto": produto,
                        "estoque_inicial": ini,
                        "estoque_final": fin,
                    })

    return {
        "competencia": competencia,
        "tanques": tanques,
        "bicos": bicos,
    }


# ── EXCEL ESTRUTURADO (formato BP Combustíveis e similares) ──────────────────
def _parse_excel_estruturado(arquivo_bytes):
    """Parser para Excel com seções ESTOQUE e MOVIMENTAÇÃO POR BICO."""
    import openpyxl
    wb = openpyxl.load_workbook(io.BytesIO(arquivo_bytes), data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))

    tanques=[]; bicos=[]; competencia=""

    # Detectar período
    for row in rows:
        for cell in row:
            if cell and re.search(r'Periodo|Período', str(cell), re.I):
                # Tentar encontrar datas na mesma linha ou próximas
                for c in row:
                    m = re.search(r'(\d{2}/\d{2}/\d{4})', str(c) if c else "")
                    if m:
                        dt = m.group(1)
                        meses=['Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set','Out','Nov','Dez']
                        try: competencia=f"{meses[int(dt[3:5])-1]}/{dt[6:]}"
                        except: pass

    em_bicos=False; em_tanques=False
    bicos_vistos=set()

    for row in rows:
        c0 = str(row[0]).strip() if row[0] is not None else ""
        c1 = str(row[1]).strip() if len(row)>1 and row[1] is not None else ""

        # Detectar seções
        if re.search(r'MOVIMENTA[ÇC][ÃA]O', c0, re.I):
            em_bicos=True; em_tanques=False; continue
        if re.search(r'^ESTOQUE$', c0, re.I):
            em_tanques=True; em_bicos=False; continue
        if re.search(r'^(Serie|Série|Tanque|Bico|Produto)', c0, re.I):
            continue

        if em_bicos and len(row)>=6:
            # [Serie, Bico, Produto, '', Abertura, Fechamento]
            bid_raw = str(row[1]).strip() if row[1] is not None else ""
            if bid_raw.isdigit() and bid_raw not in bicos_vistos:
                ini = _fl(row[4]) if len(row)>4 and row[4] is not None else None
                fin = _fl(row[5]) if len(row)>5 and row[5] is not None else None
                if ini is not None and fin is not None:
                    bicos.append({"id":_nid(bid_raw),"encerrante_inicial":ini,"encerrante_final":fin})
                    bicos_vistos.add(bid_raw)

        if em_tanques and len(row)>=8:
            # [Tanque, Combustivel, '', '', Est.Abert, '', '', Est.Fech]
            tid_raw = c0
            if tid_raw.isdigit():
                produto = c1
                ini = _fl(row[4]) if row[4] is not None else None
                fin = _fl(row[7]) if len(row)>7 and row[7] is not None else None
                if ini is not None and fin is not None:
                    tanques.append({"id":_nid(tid_raw),"produto":produto,
                                    "estoque_inicial":ini,"estoque_final":fin})

    return {"competencia":competencia,"tanques":tanques,"bicos":bicos}


# ── FUNÇÃO PRINCIPAL ──────────────────────────────────────────────────────────
def ler_dac(arquivo_bytes, filename):
    """
    Lê o DAC em qualquer formato suportado.
    Retorna dict {"competencia", "tanques", "bicos"} ou None se não suportado.
    """
    ext = filename.lower().rsplit(".", 1)[-1] if "." in filename else ""

    try:
        if ext in ("xlsx", "xlsm"):
            # Tentar parser estruturado primeiro, depois texto
            resultado = _parse_excel_estruturado(arquivo_bytes)
            if not resultado["tanques"] and not resultado["bicos"]:
                resultado = _ler_excel(arquivo_bytes, filename)
            return resultado if (resultado["tanques"] or resultado["bicos"]) else None

        elif ext == "xls":
            return _ler_excel(arquivo_bytes, filename)

        elif ext == "pdf":
            return _ler_pdf(arquivo_bytes)  # retorna None se escaneado

        else:
            return None  # formato não suportado

    except Exception as e:
        print(f"Erro ao ler DAC ({filename}): {e}")
        return None


# ── CONFRONTO DAC × SPED ──────────────────────────────────────────────────────
def confrontar_dac_sped(dac, d_atu):
    from app import _sk
    resultado = {"tanques": [], "bicos": [], "competencia": dac.get("competencia", "")}

    tanq_pri={}; tanq_ult={}
    for (t,d) in d_atu["tanques"]:
        if t not in tanq_pri or d<tanq_pri[t]: tanq_pri[t]=d
        if t not in tanq_ult or d>tanq_ult[t]: tanq_ult[t]=d

    bico_pri={}; bico_ult={}
    for (b,d) in d_atu["bicos"]:
        if b not in bico_pri or d<bico_pri[b]: bico_pri[b]=d
        if b not in bico_ult or d>bico_ult[b]: bico_ult[b]=d

    def _st(dif):
        if dif is None: return "⚠️ AUSENTE"
        return "✅ OK" if abs(dif) < 0.01 else "❌ DIVERGÊNCIA"

    # Tanques
    for item in sorted(dac.get("tanques",[]), key=lambda x: _sk(str(x.get("id","")))):
        tid     = str(item.get("id","")).strip()
        produto = item.get("produto","")
        ei_dac  = item.get("estoque_inicial")
        ef_dac  = item.get("estoque_final")
        pri_d   = tanq_pri.get(tid); ult_d = tanq_ult.get(tid)
        r_pri   = d_atu["tanques"].get((tid,pri_d),{}) if pri_d else {}
        r_ult   = d_atu["tanques"].get((tid,ult_d),{}) if ult_d else {}
        ei_sped = r_pri.get("est_abert"); ef_sped = r_ult.get("est_fech")
        dif_ini = round(ei_dac-ei_sped,3) if ei_dac is not None and ei_sped is not None else None
        dif_fin = round(ef_dac-ef_sped,3) if ef_dac is not None and ef_sped is not None else None

        # Diagnóstico
        dias_t = sorted(d for (tt,d) in d_atu["tanques"] if tt==tid)
        total_rec  = round(sum(d_atu["tanques"][(tid,d)].get("entrada",0) for d in dias_t),3)
        total_sai  = round(sum(d_atu["tanques"][(tid,d)].get("saida",0)   for d in dias_t),3)
        total_evap = round(sum(d_atu["tanques"][(tid,d)].get("evap",0)    for d in dias_t),3)
        total_aj   = round(sum(d_atu["tanques"][(tid,d)].get("ajuste",0)  for d in dias_t),3)

        diagnostico = ""
        if dif_fin is not None and abs(dif_fin) >= 0.01:
            pct_est = abs(dif_fin)/ef_sped*100 if ef_sped else 0
            pct_ven = abs(dif_fin)/total_sai*100 if total_sai else 0
            if dif_fin > 0:
                causas = []
                if total_aj>0: causas.append(f"ajuste positivo já lançado ({total_aj:,.3f} L)")
                causas += ["entrada não lançada no sistema","venda registrada a maior que a real","ganho por temperatura/dilatação"]
                diagnostico=(f"SOBRA FÍSICA de {dif_fin:,.3f} L ({pct_est:.1f}% do est. final). "
                             f"Possíveis causas: {'; '.join(causas[:3])}.")
            else:
                causas = []
                if total_aj<0: causas.append(f"ajuste negativo já lançado ({total_aj:,.3f} L)")
                causas += [f"venda maior que vol. escritural ({pct_ven:.1f}% das saídas)",
                           f"evaporação não contabilizada (lançado: {total_evap:,.3f} L)",
                           "vazamento ou perda não registrada"]
                diagnostico=(f"FALTA FÍSICA de {abs(dif_fin):,.3f} L ({pct_est:.1f}% do est. final). "
                             f"Possíveis causas: {'; '.join(causas[:3])}.")

        resultado["tanques"].append({
            "id":tid,"produto":produto,
            "ei_dac":ei_dac,"ei_sped":ei_sped,"dif_ini":dif_ini,"status_ini":_st(dif_ini),
            "ef_dac":ef_dac,"ef_sped":ef_sped,"dif_fin":dif_fin,"status_fin":_st(dif_fin),
            "diagnostico":diagnostico,
            "total_entrada":total_rec,"total_saida":total_sai,
            "total_evap":total_evap,"total_ajuste":total_aj,
        })

    # Bicos
    for item in sorted(dac.get("bicos",[]), key=lambda x: _sk(str(x.get("id","")))):
        bid    = str(item.get("id","")).strip()
        ei_dac = item.get("encerrante_inicial"); ef_dac = item.get("encerrante_final")
        pri_d  = bico_pri.get(bid); ult_d = bico_ult.get(bid)
        r_pri  = d_atu["bicos"].get((bid,pri_d),{}) if pri_d else {}
        r_ult  = d_atu["bicos"].get((bid,ult_d),{}) if ult_d else {}
        ei_sped= r_pri.get("enc_abert"); ef_sped = r_ult.get("enc_fech")
        dif_ini= round(ei_dac-ei_sped,3) if ei_dac is not None and ei_sped is not None else None
        dif_fin= round(ef_dac-ef_sped,3) if ef_dac is not None and ef_sped is not None else None
        resultado["bicos"].append({
            "id":bid,
            "ei_dac":ei_dac,"ei_sped":ei_sped,"dif_ini":dif_ini,"status_ini":_st(dif_ini),
            "ef_dac":ef_dac,"ef_sped":ef_sped,"dif_fin":dif_fin,"status_fin":_st(dif_fin),
        })

    return resultado
