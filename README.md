# 🔐 Key Manager

**Key Manager** é uma aplicação desktop robusta para organizar, armazenar e gerenciar **chaves de ativação de produtos** de forma segura, eficiente e personalizável.  
Ideal para **vendedores de software**, **técnicos de informática**, **empresas** ou qualquer pessoa que lida com grandes volumes de **licenças digitais**.

---

## 🚀 Funcionalidades Principais

- 📋 Cadastro, edição e exclusão de chaves de ativação  
- 🔍 Filtros por produto, status, canal de venda ou categoria  
- 📁 Organização por categorias com layout e instruções personalizados  
- 🧾 Geração de PDFs com layout customizado e logo da categoria  
- 🌙 Tema escuro completo  
- 🌐 Suporte multilíngue (PT, EN, ES)  
- 📧 Envio de chaves por e-mail com opção de anexo PDF  
- 📦 Exportação de dados em JSON e Excel  
- 🔄 Funcionalidade de desfazer/refazer alterações  
- ⏳ Backup automático do banco de dados  
- 📊 Dashboard de vendas com relatórios  
- 📥 Importação de chaves diretamente de arquivos `.XLS/.XLSX`  

---

## 🛠 Tecnologias Utilizadas

- **Python** + **Tkinter** para a interface gráfica  
- **SQLite** como banco de dados local  
- **ReportLab** para geração de PDFs personalizados  
- **pandas** para importação de arquivos Excel  
- **smtplib** para envio de e-mails com chave e PDF  
- **threading** para operações assíncronas como envio de e-mail  

---

## 📷 Captura de Tela

<img width="800" alt="Captura de Tela do Key Manager" src="https://github.com/user-attachments/assets/93a82e66-24d8-404e-9f8a-10687ac50a04" />

---

## 📦 Instalação

1. Clone o repositório:
   ```bash
   git clone https://github.com/VininLeao/key-manager.git
   cd key-manager
   ```

2. Instale as dependências necessárias:
   ```bash
   pip install reportlab pandas openpyxl pyperclip
   ou
   pip install -r requirements.txt
   ```

3. Execute a aplicação:
   ```bash
   python main.py
   ```

---

## ✅ Requisitos

- Python 3.8 ou superior  
- Dependências: `reportlab`, `pandas`, `openpyxl`, `pyperclip`

---

## 💬 Idiomas Suportados

- 🇧🇷 Português  
- 🇺🇸 Inglês  
- 🇪🇸 Espanhol  

---

## 🧪 Testado em

- Windows 10 e 11  
- Resoluções Full HD e superiores  

---

## 📌 Observações

- Todas as instruções, logos e informações das categorias podem ser configuradas individualmente.  
- O sistema já detecta e migra automaticamente dados antigos em `.json` para o banco SQLite.

---

## 🤝 Contribuições

Contribuições são bem-vindas!  
Sinta-se à vontade para abrir *issues*, enviar *pull requests* ou sugerir novas funcionalidades.

---

Desenvolvido por **Vinícius Leão**
