<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Sistema LMC – Conferência SPED Fiscal</title>
<style>
  :root{
    --azul:#1F3864;--azul-m:#2E75B6;--azul-c:#D6E4F0;
    --verde:#375623;--verde-bg:#C6EFCE;
    --verm:#9C0006;--verm-bg:#FFC7CE;
    --cinza:#F4F6F9;--borda:#D0D8E4;--texto:#1a2340;
    --fonte:'Segoe UI',Arial,sans-serif;
  }
  *{box-sizing:border-box;margin:0;padding:0}
  body{font-family:var(--fonte);background:var(--cinza);color:var(--texto);min-height:100vh;}

  /* TOPO */
  header{
    background:linear-gradient(135deg,var(--azul) 0%,var(--azul-m) 100%);
    color:#fff;padding:0 40px;
    display:flex;align-items:center;justify-content:space-between;
    height:64px;box-shadow:0 4px 18px rgba(31,56,100,.25);
  }
  .header-left{display:flex;align-items:center;gap:14px;}
  header h1{font-size:1.15rem;font-weight:700;letter-spacing:-.2px;}
  header p{font-size:.76rem;opacity:.78;margin-top:1px;}
  .header-user{display:flex;align-items:center;gap:12px;font-size:.82rem;}
  .header-user span{opacity:.85;}
  .btn-logout{
    background:rgba(255,255,255,.15);border:1px solid rgba(255,255,255,.3);
    color:#fff;padding:6px 14px;border-radius:6px;font-size:.78rem;
    font-weight:600;cursor:pointer;text-decoration:none;transition:background .2s;
  }
  .btn-logout:hover{background:rgba(255,255,255,.25);}

  /* MAIN */
  main{max-width:820px;margin:36px auto;padding:0 24px 60px;}

  /* CARD */
  .card{
    background:#fff;border-radius:12px;padding:32px;
    box-shadow:0 2px 12px rgba(31,56,100,.08);
    border:1px solid var(--borda);margin-bottom:22px;
  }
  .card h2{
    font-size:.98rem;font-weight:700;color:var(--azul);
    margin-bottom:18px;display:flex;align-items:center;gap:9px;
  }
  .num{
    background:var(--azul);color:#fff;width:26px;height:26px;border-radius:50%;
    display:inline-flex;align-items:center;justify-content:center;
    font-size:.75rem;font-weight:800;flex-shrink:0;
  }

  /* UPLOAD */
  .uploads{display:grid;grid-template-columns:1fr 1fr;gap:16px;}
  @media(max-width:580px){.uploads{grid-template-columns:1fr;}}
  .upload-box{
    border:2px dashed var(--borda);border-radius:10px;
    padding:22px 18px;text-align:center;cursor:pointer;
    transition:border-color .2s,background .2s;position:relative;
  }
  .upload-box:hover,.upload-box.drag{border-color:var(--azul-m);background:var(--azul-c);}
  .upload-box.ok{border-color:#375623;background:var(--verde-bg);border-style:solid;}
  .upload-box input{position:absolute;inset:0;opacity:0;cursor:pointer;width:100%;}
  .upload-icon{font-size:2rem;margin-bottom:8px;color:var(--azul-m);}
  .upload-box.ok .upload-icon{color:var(--verde);}
  .upload-label{font-size:.82rem;font-weight:600;color:var(--azul);}
  .upload-sub{font-size:.74rem;color:#6b7a99;margin-top:3px;}
  .upload-name{font-size:.76rem;color:var(--verde);font-weight:600;margin-top:6px;word-break:break-all;}

  /* TAGS */
  .mes-tags{display:flex;gap:10px;flex-wrap:wrap;margin-bottom:6px;}
  .mes-tag{display:flex;align-items:center;gap:6px;background:var(--azul-c);border-radius:6px;padding:4px 11px;font-size:.76rem;font-weight:600;color:var(--azul);}
  .dot{width:8px;height:8px;border-radius:50%;}
  .dot-ant{background:var(--azul-m);}
  .dot-atu{background:#E35D3B;}

  /* BOTÃO GERAR */
  .btn-gerar{
    width:100%;padding:14px;border:none;border-radius:10px;
    background:linear-gradient(135deg,var(--azul),var(--azul-m));
    color:#fff;font-size:1rem;font-weight:700;cursor:pointer;
    display:flex;align-items:center;justify-content:center;gap:10px;
    transition:opacity .2s,transform .1s;
    box-shadow:0 4px 14px rgba(31,56,100,.25);
  }
  .btn-gerar:disabled{opacity:.5;cursor:not-allowed;transform:none!important;}
  .btn-gerar:not(:disabled):hover{opacity:.92;transform:translateY(-1px);}
  .btn-gerar:not(:disabled):active{transform:translateY(0);}

  /* ALERTAS FLASK */
  .alert{border-radius:8px;padding:12px 16px;margin-bottom:16px;font-size:.85rem;font-weight:600;}
  .alert-danger{background:var(--verm-bg);color:var(--verm);border:1px solid var(--verm);}
  .alert-success{background:var(--verde-bg);color:var(--verde);border:1px solid var(--verde);}
  .alert-warning{background:#FFEB9C;color:#7D6608;border:1px solid #7D6608;}

  /* INFO */
  .info-list{list-style:none;}
  .info-list li{display:flex;align-items:flex-start;gap:10px;font-size:.81rem;color:#556;padding:7px 0;border-bottom:1px solid var(--borda);}
  .info-list li:last-child{border-bottom:none;}
  .tag{background:var(--azul-c);color:var(--azul);border-radius:4px;padding:1px 7px;font-size:.71rem;font-weight:700;white-space:nowrap;flex-shrink:0;margin-top:1px;}

  /* LOADING */
  .loading-overlay{display:none;position:fixed;inset:0;background:rgba(31,56,100,.55);z-index:999;align-items:center;justify-content:center;}
  .loading-overlay.ativo{display:flex;}
  .loading-box{background:#fff;border-radius:14px;padding:36px 44px;text-align:center;box-shadow:0 8px 40px rgba(0,0,0,.25);}
  .spinner{width:44px;height:44px;border:4px solid var(--azul-c);border-top-color:var(--azul-m);border-radius:50%;animation:spin 1s linear infinite;margin:0 auto 16px;}
  .loading-box p{font-size:.9rem;font-weight:600;color:var(--azul);}
  .loading-box small{font-size:.76rem;color:#667;margin-top:4px;display:block;}
  @keyframes spin{to{transform:rotate(360deg)}}
</style>
</head>
<body>

<header>
  <div class="header-left">
    <svg width="36" height="36" viewBox="0 0 36 36" fill="none">
      <rect width="36" height="36" rx="8" fill="rgba(255,255,255,.15)"/>
      <rect x="7" y="9" width="22" height="3" rx="1.5" fill="white"/>
      <rect x="7" y="15" width="15" height="3" rx="1.5" fill="white"/>
      <rect x="7" y="21" width="18" height="3" rx="1.5" fill="white"/>
      <circle cx="27" cy="27" r="6" fill="#4CA3E8"/>
      <path d="M24.5 27l2 2 3.5-3.5" stroke="white" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"/>
    </svg>
    <div>
      <h1>Sistema LMC – Conferência SPED Fiscal</h1>
      <p>Registros 1310 (Tanques) e 1320 (Bicos) · Postos de Combustíveis</p>
    </div>
  </div>
  <div class="header-user">
    <span>👤 {{ nome }}</span>
    <a href="/logout" class="btn-logout">Sair</a>
  </div>
</header>

<main>

  {% with msgs = get_flashed_messages(with_categories=true) %}
    {% for cat, msg in msgs %}
      <div class="alert alert-{{ cat }}">{{ msg }}</div>
    {% endfor %}
  {% endwith %}

  <!-- CARD 1: UPLOAD -->
  <div class="card">
    <h2><span class="num">1</span> Selecione os arquivos SPED Fiscal (.txt)</h2>
    <div class="mes-tags">
      <div class="mes-tag"><span class="dot dot-ant"></span>Competência Anterior (opcional)</div>
      <div class="mes-tag"><span class="dot dot-atu"></span>Competência Atual (obrigatória)</div>
    </div>
    <p style="font-size:.76rem;color:#6b7a99;margin-bottom:16px;margin-top:4px;">
      Versão e Capacidade dos Tanques são verificadas apenas na competência atual.<br>
      Se a competência anterior não for enviada, o sistema confronta apenas o mês atual
      (consistência diária, negativos, ANP e DAC × SPED) — sem o confronto entre meses.
    </p>

    <form id="form-lmc" action="/processar" method="POST" enctype="multipart/form-data">
      <div class="uploads">
        <div class="upload-box" id="box-ant" ondragover="drag(event,'ant')" ondragleave="undrag('ant')" ondrop="drop(event,'ant')">
          <input type="file" name="ant" accept=".txt" onchange="pick(event,'ant')">
          <div class="upload-icon" id="icon-ant">📂</div>
          <div class="upload-label">Competência Anterior</div>
          <div class="upload-sub">Ex.: SPED abril/2026 (opcional)</div>
          <div class="upload-name" id="name-ant"></div>
        </div>
        <div class="upload-box" id="box-atu" ondragover="drag(event,'atu')" ondragleave="undrag('atu')" ondrop="drop(event,'atu')">
          <input type="file" name="atu" accept=".txt" onchange="pick(event,'atu')" required>
          <div class="upload-icon" id="icon-atu">📂</div>
          <div class="upload-label">Competência Atual</div>
          <div class="upload-sub">Ex.: SPED maio/2026</div>
          <div class="upload-name" id="name-atu"></div>
        </div>
      </div>

      <!-- DAC opcional -->
      <div style="margin-top:16px;">
        <p style="font-size:.78rem;font-weight:600;color:var(--azul);margin-bottom:8px;">
          📋 DAC (opcional) — Excel, PDF ou imagem
        </p>
        <p style="font-size:.74rem;color:#6b7a99;margin-bottom:10px;">
          Apenas para empresas que enviam o relatório DAC. Quando enviado, o sistema confronta os valores do DAC com o SPED atual e gera uma aba extra no relatório.
        </p>
        <div class="upload-box" id="box-dac" style="border-style:dashed;border-color:#b0bec5;"
             ondragover="drag(event,'dac')" ondragleave="undrag('dac')" ondrop="drop(event,'dac')">
          <input type="file" name="dac" accept=".xlsx,.xls,.pdf,.jpg,.jpeg,.png,.webp" onchange="pick(event,'dac')">
          <div class="upload-icon" id="icon-dac" style="color:#78909c;">📄</div>
          <div class="upload-label" style="color:#546e7a;">Relatório DAC</div>
          <div class="upload-sub">Excel, PDF ou imagem · opcional</div>
          <div class="upload-name" id="name-dac"></div>
        </div>
      </div>

      <!-- CARD 2: BOTÃO -->
      <div style="margin-top:22px;">
        <button type="submit" class="btn-gerar" id="btn-gerar" disabled onclick="iniciarProcessamento()">
          <span id="btn-icon">⚙️</span>
          <span id="btn-txt">Selecione os dois arquivos para continuar</span>
        </button>
      </div>
    </form>
  </div>

  <!-- CARD 3: INFO -->
  <div class="card">
    <h2><span class="num">2</span> O que é verificado</h2>
    <ul class="info-list">
      <li><span class="tag">CONFRONTO</span> Fechamento da competência anterior × Abertura da competência atual — tanques e bicos</li>
      <li><span class="tag">DIÁRIO</span> Fechamento do dia N = Abertura do dia N+1, para todos os tanques e bicos da competência atual</li>
      <li><span class="tag">NEGATIVOS</span> Detecção de valores negativos nos registros 1310 e 1320 (ambas as competências)</li>
      <li><span class="tag">VERSÃO</span> Confirmação de que o SPED está na versão 020 (obrigatória)</li>
      <li><span class="tag">CAPACIDADE</span> Verificação da capacidade declarada em cada tanque — obrigatória desde jan/2026</li>
    </ul>
  </div>

</main>

<!-- LOADING OVERLAY -->
<div class="loading-overlay" id="loading">
  <div class="loading-box">
    <div class="spinner"></div>
    <p>Processando os arquivos SPED…</p>
    <small>Gerando relatório Excel, aguarde.</small>
  </div>
</div>

<script>
const selecionados = {ant: false, atu: false, dac: false};

function pick(e, key) {
  const f = e.target.files[0];
  if (f) marcarArquivo(key, f.name);
}
function drag(e, key) { e.preventDefault(); document.getElementById('box-'+key).classList.add('drag'); }
function undrag(key)  { document.getElementById('box-'+key).classList.remove('drag'); }
function drop(e, key) {
  e.preventDefault(); undrag(key);
  const f = e.dataTransfer.files[0];
  if (!f) return;
  // Injetar no input file
  const input = document.querySelector(`#box-${key} input`);
  const dt = new DataTransfer();
  dt.items.add(f);
  input.files = dt.files;
  marcarArquivo(key, f.name);
}
function marcarArquivo(key, nome) {
  selecionados[key] = true;
  document.getElementById('box-'+key).classList.add('ok');
  document.getElementById('icon-'+key).textContent = '✅';
  document.getElementById('name-'+key).textContent = nome;
  atualizarBtn();
}
function atualizarBtn() {
  const btn = document.getElementById('btn-gerar');
  const txt = document.getElementById('btn-txt');
  // Apenas a competência atual é obrigatória; a anterior é opcional.
  const ok  = selecionados.atu;
  btn.disabled = !ok;
  if (selecionados.ant && selecionados.atu)
                       txt.textContent = 'Gerar Relatório Excel';
  else if (selecionados.atu)
                       txt.textContent = 'Gerar Relatório (sem confronto entre meses)';
  else                 txt.textContent = 'Selecione ao menos a competência atual';
}
function iniciarProcessamento() {
  document.getElementById('loading').classList.add('ativo');
}
</script>

</body>
</html>
