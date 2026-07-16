<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Login – Sistema LMC</title>
<style>
  *{box-sizing:border-box;margin:0;padding:0}
  body{font-family:'Segoe UI',Arial,sans-serif;background:linear-gradient(135deg,#1F3864 0%,#2E75B6 100%);min-height:100vh;display:flex;align-items:center;justify-content:center;}
  .card{background:#fff;border-radius:16px;padding:48px 40px 40px;width:100%;max-width:400px;box-shadow:0 20px 60px rgba(0,0,0,.25);}
  .logo{text-align:center;margin-bottom:32px;}
  .logo svg{margin-bottom:12px;}
  .logo h1{font-size:1.3rem;font-weight:700;color:#1F3864;line-height:1.2;}
  .logo p{font-size:.8rem;color:#6b7a99;margin-top:4px;}
  .field{margin-bottom:18px;}
  label{display:block;font-size:.82rem;font-weight:600;color:#334;margin-bottom:6px;}
  input{width:100%;padding:12px 14px;border:1.5px solid #D0D8E4;border-radius:8px;font-size:.95rem;color:#1a2340;transition:border-color .2s,box-shadow .2s;outline:none;}
  input:focus{border-color:#2E75B6;box-shadow:0 0 0 3px rgba(46,117,182,.15);}
  .btn{width:100%;padding:13px;background:linear-gradient(135deg,#1F3864,#2E75B6);color:#fff;border:none;border-radius:8px;font-size:1rem;font-weight:700;cursor:pointer;transition:opacity .2s,transform .1s;margin-top:6px;}
  .btn:hover{opacity:.92;transform:translateY(-1px);}
  .btn:active{transform:translateY(0);}
  .erro{background:#FFC7CE;color:#9C0006;border:1px solid #9C0006;border-radius:8px;padding:10px 14px;font-size:.85rem;font-weight:600;margin-bottom:18px;}
  .rodape{text-align:center;margin-top:24px;font-size:.75rem;color:#9aabbf;}
</style>
</head>
<body>
<div class="card">
  <div class="logo">
    <svg width="52" height="52" viewBox="0 0 52 52" fill="none">
      <rect width="52" height="52" rx="13" fill="#1F3864"/>
      <rect x="10" y="13" width="32" height="4" rx="2" fill="white"/>
      <rect x="10" y="21" width="22" height="4" rx="2" fill="white"/>
      <rect x="10" y="29" width="27" height="4" rx="2" fill="white"/>
      <circle cx="40" cy="38" r="9" fill="#4CA3E8"/>
      <path d="M37 38l2.5 2.5 5-5" stroke="white" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/>
    </svg>
    <h1>Sistema LMC</h1>
    <p>Conferência SPED Fiscal · Postos de Combustíveis</p>
  </div>

  {% if erro %}
  <div class="erro">⚠️ {{ erro }}</div>
  {% endif %}

  <form method="POST">
    <div class="field">
      <label for="usuario">Usuário</label>
      <input type="text" id="usuario" name="usuario" placeholder="Digite seu usuário" autocomplete="username" required>
    </div>
    <div class="field">
      <label for="senha">Senha</label>
      <input type="password" id="senha" name="senha" placeholder="Digite sua senha" autocomplete="current-password" required>
    </div>
    <button type="submit" class="btn">Entrar</button>
  </form>

  <div class="rodape">Sistema LMC · Registros 1310 e 1320 · v2.0</div>
</div>
</body>
</html>
