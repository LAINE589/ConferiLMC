# Sistema LMC – Conferência SPED Fiscal
## Cleodon Contabilidade

---

## Como publicar no Render.com (passo a passo)

### Passo 1 – Criar conta no GitHub
1. Acesse **github.com** e clique em **Sign up**
2. Crie sua conta gratuitamente

### Passo 2 – Criar repositório e subir os arquivos
1. Clique em **New repository**
2. Nome: `sistema-lmc` → clique em **Create repository**
3. Clique em **uploading an existing file**
4. Extraia o ZIP e arraste **todos os arquivos** desta pasta
5. Clique em **Commit changes**

### Passo 3 – Publicar no Render.com
1. Acesse **render.com** e crie uma conta gratuita
2. Clique em **New +** → **Web Service**
3. Conecte sua conta GitHub → selecione o repositório `sistema-lmc`
4. Preencha:
   - **Name:** sistema-lmc
   - **Runtime:** Python 3
   - **Build Command:** `pip install -r requirements.txt`
   - **Start Command:** `gunicorn app:app --workers 4 --timeout 120`
5. Clique em **Create Web Service**
6. Aguarde ~2 minutos → o sistema estará no ar com uma URL pública

### Passo 4 – Configurar a chave da API (para leitura automática do DAC)
No painel do Render → **Environment** → adicionar:

| Variável | Valor |
|---|---|
| `ANTHROPIC_API_KEY` | Sua chave em console.anthropic.com |

> Sem essa chave o sistema funciona normalmente para os SPEDs — só não lê o DAC automaticamente.

---

## Usuários do sistema

| Login | Senha |
|---|---|
| fiscal@cleodoncontabilidade.com.br | Cld@123 |
| lucroreal@cleodoncontabilidade.com.br | Cld@123 |

---

## Como atualizar o sistema após mudanças
1. Substitua os arquivos modificados no repositório GitHub
2. O Render faz o redeploy automaticamente em ~1 minuto

## Observação sobre o plano gratuito do Render
O sistema pode "dormir" após 15 minutos sem acesso — na primeira visita pode demorar ~30s para acordar. Para acesso contínuo sem espera, considere o plano pago (US$ 7/mês).
