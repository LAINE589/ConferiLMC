# Sistema LMC – Cleodon Contabilidade
## Conferência SPED Fiscal – Postos de Combustíveis

---

## OPÇÃO 1: Usar no computador (GRATUITO)

### Requisitos
- Windows 7, 8, 10 ou 11
- Conexão com internet (só na primeira instalação)

### Instalação (apenas uma vez)
1. Baixe e extraia o arquivo ZIP numa pasta (ex: `C:\SistemaLMC`)
2. Certifique-se de que o **Python** está instalado:
   - Acesse: https://www.python.org/downloads/
   - Baixe a versão mais recente e instale
   - ⚠️ **IMPORTANTE:** marque a opção **"Add Python to PATH"**
3. Clique duas vezes em **`instalar.bat`**
4. Aguarde a instalação terminar

### Como usar
1. Clique duas vezes em **`iniciar.bat`**
2. O navegador abre automaticamente em `http://localhost:5000`
3. Faça login com suas credenciais
4. Para encerrar, feche a janela preta do terminal

### Usuários
| Login | Senha |
|---|---|
| fiscal@cleodoncontabilidade.com.br | Cld@123 |
| lucroreal@cleodoncontabilidade.com.br | Cld@123 |

---

## OPÇÃO 2: Usar na web (Render.com – US$ 7/mês)

### Passo 1 – GitHub
1. Crie conta em **github.com**
2. Clique em **New repository** → nome: `sistema-lmc` → **Create**
3. Clique em **uploading an existing file**
4. Arraste todos os arquivos desta pasta → **Commit changes**

### Passo 2 – Render.com
1. Crie conta em **render.com**
2. Clique em **New + → Web Service**
3. Conecte o GitHub → selecione `sistema-lmc`
4. Configure:
   - **Build Command:** `pip install -r requirements.txt`
   - **Start Command:** `gunicorn app:app --workers 4 --timeout 120`
5. Selecione o plano **Starter (US$ 7/mês)**
6. Clique em **Create Web Service**

---

## Sobre o DAC
O sistema lê automaticamente:
- ✅ Excel (.xlsx) com seções ESTOQUE e MOVIMENTAÇÃO
- ✅ PDF com texto extraível (AutoSystem PRO, Linx)
- ❌ PDF escaneado ou imagem → gera relatório sem confronto DAC

---

## Suporte
Em caso de dúvidas, entre em contato com o desenvolvedor.
