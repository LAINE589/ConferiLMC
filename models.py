{% extends "base.html" %}
{% block title %}Enviar SPED{% endblock %}
{% block sidebar %}
<div class="sidebar">
  <div class="sidebar-sec">Meu Posto</div>
  <a href="/posto">🏠 Painel</a>
  <a href="/posto/conferir" class="ativo">📤 Enviar SPED</a>
  <div class="sidebar-sec">Conta</div>
  <a href="/perfil">🔑 Alterar Senha</a>
</div>
{% endblock %}
{% block content %}
<div style="margin-bottom:16px">
  <a href="/posto" style="color:var(--azul2);font-size:.85rem">← Voltar ao Painel</a>
</div>
<h2 style="margin-bottom:20px;color:var(--azul)">Verificar SPED Fiscal</h2>

<div class="card">
  <div class="alert alert-info" style="margin-bottom:20px">
    <strong>ℹ️ Como funciona:</strong> envie os arquivos SPED do mês anterior (opcional) e do mês atual.
    O sistema gera automaticamente um relatório com todas as divergências encontradas.
    O arquivo do SPED não é armazenado — apenas o relatório Excel resultante.
  </div>

  <form method="POST" action="/posto/processar" enctype="multipart/form-data" id="form-posto">
    <!-- SPED anterior -->
    <div class="form-group">
      <label>📁 SPED Competência Anterior <span style="color:#999;font-weight:400">(opcional — necessário para confronto entre meses)</span></label>
      <input type="file" name="ant" accept=".txt" onchange="marcarArquivo(this,'label-ant')">
      <div id="label-ant" style="font-size:.78rem;color:#999;margin-top:4px">Nenhum arquivo selecionado</div>
    </div>

    <!-- SPED atual -->
    <div class="form-group">
      <label>📁 SPED Competência Atual <span style="color:var(--verm)">*</span></label>
      <input type="file" name="atu" accept=".txt" required onchange="marcarArquivo(this,'label-atu')">
      <div id="label-atu" style="font-size:.78rem;color:#999;margin-top:4px">Nenhum arquivo selecionado</div>
    </div>

    <!-- DAC opcional -->
    <div class="form-group">
      <label>📁 DAC (Documento de Apuração de Combustíveis) <span style="color:#999;font-weight:400">(opcional — PDF ou Excel)</span></label>
      <input type="file" name="dac" accept=".pdf,.xlsx,.xls" onchange="marcarArquivo(this,'label-dac')">
      <div id="label-dac" style="font-size:.78rem;color:#999;margin-top:4px">Nenhum arquivo selecionado</div>
    </div>

    <div style="margin-top:20px;display:flex;gap:12px;align-items:center">
      <button type="submit" class="btn btn-primary" id="btn-enviar" style="font-size:.95rem;padding:11px 28px">
        ⚙️ Gerar Relatório
      </button>
      <div id="loading" style="display:none;color:#666;font-size:.85rem">
        ⏳ Processando... aguarde.
      </div>
    </div>
  </form>
</div>

<div class="card" style="background:#f8faff">
  <div class="card-title">O que é verificado</div>
  <div style="display:grid;grid-template-columns:1fr 1fr;gap:12px;font-size:.84rem">
    <div>✅ Consistência diária (fechamento = abertura)</div>
    <div>✅ Confronto entre meses (se enviado SPED anterior)</div>
    <div>✅ Valores negativos em tanques e bicos</div>
    <div>✅ Versão do SPED (obrigatório versão 020)</div>
    <div>✅ Capacidade declarada dos tanques</div>
    <div>✅ Limite ANP de 0,6% sobre recebimentos</div>
    <div>✅ DAC × SPED (se enviado arquivo DAC)</div>
    <div>✅ Relatório de divergências para envio ao cliente</div>
  </div>
</div>

<script>
function marcarArquivo(input, labelId) {
  const label = document.getElementById(labelId);
  if (input.files.length > 0) {
    label.textContent = '✅ ' + input.files[0].name;
    label.style.color = 'var(--verde)';
  } else {
    label.textContent = 'Nenhum arquivo selecionado';
    label.style.color = '#999';
  }
}

document.getElementById('form-posto').onsubmit = function() {
  document.getElementById('btn-enviar').disabled = true;
  document.getElementById('loading').style.display = 'inline';
};
</script>
{% endblock %}
