# 🎓 Gerador de Plano de Aula Web

Aplicação web moderna para gerar **Planos de Aula** semanais de cursos técnicos a partir de planilhas (Excel / CSV) e modelos Word (`.docx`).

Pronto para deploy gratuito na **Vercel** ou execução local com Python/Flask.

---

## 🚀 Como fazer Deploy na Vercel

1. Acesse o painel da [Vercel](https://vercel.com).
2. Clique em **"Add New..."** ➔ **"Project"**.
3. Importe o repositório `meuplanodeaula`.
4. Deixe as configurações padrão e clique em **"Deploy"**.
5. Em poucos segundos seu site estará no ar com link público!

---

## 💻 Como rodar Localmente no computador

```bash
# 1. Clone o repositório
git clone https://github.com/thamyoliveira-create/meuplanodeaula.git
cd meuplanodeaula

# 2. Crie e ative o ambiente virtual
python -m venv .venv
source .venv/bin/activate   # macOS / Linux
# Windows: .venv\Scripts\activate

# 3. Instale as dependências
pip install -r requirements.txt

# 4. Inicie o servidor web
python app.py
```

Abra o navegador em: **http://127.0.0.1:5000**

---

## ✨ Funcionalidades

- 📤 **Upload de Planilha:** Suporte a arquivos `.xlsx`, `.xls` e `.csv`.
- 🔍 **Detecção Automática:** Identifica automaticamente as semanas disponíveis na planilha.
- 👁️ **Pré-visualização Online:** Veja os dados do plano antes de baixar.
- 📥 **Download Word (.docx):** Documento formatado pronto para edição ou impressão.
- 🖨️ **Impressão / Salvar em PDF:** Botão de impressão com layout otimizado.
- 🎨 **Modelo Customizável:** Permite usar o modelo padrão ou enviar seu próprio `.docx` com tags `{{NomeDaColuna}}`.

---

## 📄 Estrutura de Arquivos

```
├── api/
│   └── index.py            # Servidor Flask e rotas serverless da Vercel
├── templates/
│   └── index.html          # Interface visual web (Bootstrap 5)
├── lesson_plan_generator/
│   ├── plan_generator.py   # Processamento de planilhas e Word
│   └── template/
│       └── template.docx   # Modelo padrão do plano
├── app.py                  # Executável para testes locais
├── vercel.json             # Configuração de build e rotas da Vercel
└── requirements.txt        # Dependências Python
```
