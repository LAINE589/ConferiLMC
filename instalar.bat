{% extends "base.html" %}
{% block title %}Meu Painel – {{ posto.razao_social }}{% endblock %}
{% block sidebar %}
<div class="sidebar">
  <div class="sidebar-sec">Meu Posto</div>
  <a href="/posto" class="ativo">🏠 Painel</a>
  <a href="/posto/conferir">📤 Enviar SPED</a>
  <div class="sidebar-sec">Conta</div>
  <a href="/perfil">🔑 Alterar Senha</a>
</div>
{% endblock %}
{% block content %}
<!-- Cabeçalho do posto -->
<div class="card" style="margin-bottom:20px;border-left:4px solid var(--azul2)">
  <div style="display:flex;justify-content:space-between;align-items:center">
    <div>
      <h2 style="color:var(--azul)">{{ posto.razao_social }}</h2>
      {% if posto.nome_fantasia %}<div style="color:#666;font-size:.9rem">{{ posto.nome_fantasia }}</div>{% endif %}
      <div style="font-size:.82rem;color:#888;margin-top:6px">
        CNPJ: {{ posto.cnpj }}
        {% if posto.cidade %}&nbsp;|&nbsp; {{ posto.cidade }}/{{ posto.estado }}{% endif %}
      </div>
    </div>
    <div style="text-align:right">
      <span class="badge {{ 'badge-ok' if posto.licenca_ativa else 'badge-inativo' }}" style="font-size:.85rem;padding:5px 14px">
        {{ '✅ Licença Ativa' if posto.licenca_ativa else '❌ Licença Inativa' }}
      </span>
      <div style="font-size:.75rem;color:#999;margin-top:6px">Plano: {{ posto.plano|upper }}</div>
    </div>
  </div>
</div>

<!-- Ação principal -->
{% if posto.licenca_ativa %}
<div class="card" style="text-align:center;padding:40px;margin-bottom:20px">
  <div style="font-size:3rem;margin-bottom:12px">📤</div>
  <h3 style="color:var(--azul);margin-bottom:8px">Verificar SPED Fiscal</h3>
  <p style="color:#666;font-size:.9rem;margin-bottom:20px">
    Envie os arquivos do SPED para conferência antes de transmitir para a contabilidade.<br>
    O sistema verifica consistência, valores negativos, limite ANP e gera o relatório automaticamente.
  </p>
  <a href="/posto/conferir" class="btn btn-primary" style="font-size:1rem;padding:12px 32px">
    Enviar Arquivo SPED
  </a>
</div>
{% else %}
<div class="alert alert-danger" style="text-align:center;padding:24px">
  <strong>Licença inativa.</strong> Entre em contato com a Cleodon Contabilidade para reativar.
</div>
{% endif %}

<!-- Histórico -->
<div class="card">
  <div class="card-title">Meus Relatórios</div>
  {% if relatorios %}
  <table>
    <tr><th>Competência</th><th>Data</th><th>DAC</th><th>Divergências</th><th>Status</th><th>Download</th></tr>
    {% for r in relatorios %}
    <tr>
      <td>{{ r.competencia_ant or '—' }} → <strong>{{ r.competencia_atu }}</strong></td>
      <td>{{ r.gerado_em.strftime('%d/%m/%Y %H:%M') }}</td>
      <td style="text-align:center">{{ '✅' if r.tem_dac else '—' }}</td>
      <td style="text-align:center">
        {% if r.total_divergencias == 0 %}
          <span style="color:var(--verde)">Nenhuma</span>
        {% else %}
          <span style="color:{{ '#9C0006' if r.status_geral=='critico' else '#7D6608' }}">{{ r.total_divergencias }}</span>
        {% endif %}
      </td>
      <td><span class="badge badge-{{ r.status_geral }}">{{ r.status_icone }} {{ r.status_geral|upper }}</span></td>
      <td>
        {% if r.arquivo_nome %}
        <a href="/relatorio/{{ r.id }}/download" class="btn btn-sm btn-outline">⬇️ Baixar</a>
        {% else %}—{% endif %}
      </td>
    </tr>
    {% endfor %}
  </table>
  {% else %}
  <div style="text-align:center;padding:32px;color:#999">
    <div style="font-size:2.5rem;margin-bottom:8px">📋</div>
    <div>Nenhum relatório gerado ainda.</div>
    <div style="font-size:.82rem;margin-top:4px">Envie seu primeiro SPED para começar.</div>
  </div>
  {% endif %}
</div>
{% endblock %}
