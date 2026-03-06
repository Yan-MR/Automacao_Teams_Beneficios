# 🤖 Suite de Automação RPA - Benefícios corporativos (Teams & Outlook)

![Python](https://img.shields.io/badge/Python-3.11+-blue?style=for-the-badge&logo=python&logoColor=white)
![Playwright](https://img.shields.io/badge/Playwright-Automated_Testing-2EAD33?style=for-the-badge&logo=playwright&logoColor=white)
![Pandas](https://img.shields.io/badge/Pandas-Data_Manipulation-150458?style=for-the-badge&logo=pandas&logoColor=white)
![Windows](https://img.shields.io/badge/Windows-win32com-0078D6?style=for-the-badge&logo=windows&logoColor=white)

Este repositório contém uma suíte de automação (RPA) desenvolvida para otimizar e escalar a comunicação interna do setor de Benefícios. O projeto é composto por dois robôs independentes com Interface Gráfica (GUI) que realizam disparos em massa de mensagens personalizadas para os colaboradores via **Microsoft Teams**.

## ✨ Funcionalidades

* **Disparo em Massa via Microsoft Teams:** Utiliza automação web para pesquisar colaboradores, abrir o chat e enviar comunicados formatados (com tabelas HTML), lidando ativamente com latência do navegador e bloqueios de área de transferência.
* **Disparo Automático via Microsoft Outlook:** Integração nativa com o Windows via protocolo COM (`win32com`), enviando e-mails personalizados e salvando o recibo de envio diretamente na pasta de "Itens Enviados" de uma caixa de e-mail compartilhada.
* **Leitura e Escrita Dinâmica em Excel:** Consulta bases de dados em planilhas `.xlsx` para personalizar as mensagens (Nome, Matrícula, Cargo, Código de Rastreio) e atualiza o status de cada envio em tempo real ("Enviado" ou "Não Encontrado").
* **Gerador de Template Blindado:** Script auxiliar que recria a planilha base do zero com fórmulas `PROCX` avançadas já injetadas, prevenindo corrompimento de arquivos.
* **Sistema de Backup Automático:** O robô gera uma cópia de segurança da base de dados antes de iniciar os disparos, garantindo a integridade da informação original da área.
* **Interface Gráfica (GUI):** Desenvolvida com `ttkbootstrap`, oferecendo uma experiência amigável para o usuário final, com pré-visualização de mensagens, logs de status em tempo real e botões de manutenção (ex: Limpar Status).

## 🛠️ Tecnologias Utilizadas

* **Python 3:** Linguagem base da aplicação.
* **Playwright:** Automação de navegador para controle do Microsoft Teams Web.
* **pywin32 (win32com.client):** Comunicação direta com a API do Microsoft Outlook local.
* **Pandas & Openpyxl:** Manipulação avançada e segura de dados estruturados em planilhas Excel.
* **Tkinter & TTKBootstrap:** Construção da interface de usuário moderna e responsiva.
* **Shutil & OS:** Gerenciamento de rotas e criação de backups em nível de sistema operacional.

## 🚀 Como instalar e executar

1. Clone este repositório:

```bash
git clone https://github.com/SEU_USUARIO/SEU_REPOSITORIO.git
```

2. Crie e ative um ambiente virtual:

```bash
python -m venv .venv
# No Windows:
.venv\Scripts\activate
```

3. Instale as dependências:

```bash
pip install pandas openpyxl ttkbootstrap playwright pywin32 pillow
playwright install chromium
```

4. Para executar o Gerador de Planilha Base:

```bash
python gerador_planilha.py
```

5. Para executar os robôs:

```bash
python app_teams.py
# ou
python app_emails.py
```
📦 Gerando o Executável (.exe)
O projeto foi estruturado para ser compilado como um aplicativo desktop autônomo (Standalone) utilizando o PyInstaller, embutindo os ícones dinamicamente:

```bash
pip install pyinstaller
pyinstaller --noconsole --onefile --add-data "logo.png;." app_teams.py
```
## 🧠 Desafios e Soluções de Engenharia

Durante o desenvolvimento, implementei soluções robustas para lidar com comportamentos inesperados do sistema operacional e aplicações de terceiros:

"Hack" de Foco Global no Teams: Como o MS Teams web possui renderização pesada para tabelas HTML, implementamos eventos rígidos de foco de teclado (.focus(), .fill()) e buscas ativas de seletores na DOM, evitando que envios ficassem presos como rascunhos.

Redirecionamento de Caixa Compartilhada: Manipulação do objeto de sessão do MS Exchange para que e-mails disparados usando as permissões do usuário logado fossem arquivados nos "Itens Enviados" do grupo (ex: adm.beneficios).

## 👨‍💻 Autor

Desenvolvido por Yan
Jovem Aprendiz em Benefícios e Bem-Estar no Grupo Casas Bahia | Programador Fullstack formado pelo Instituto PROA.

Apaixonado por tecnologia, desenvolvimento front-end (React/Vite) e por encontrar soluções de automação em Python para problemas corporativos reais do dia a dia.
