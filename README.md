# Chatbot-automation-selenium
Automação de Envio de Mensagens com Interface Gráfica (Tkinter + Selenium + Google Sheets)
Este projeto é uma automação em Python com interface gráfica (Tkinter) que realiza o envio automatizado de mensagens para contatos obtidos de uma planilha do Google Sheets.
A automação integra autenticação via API do Google, controle de status de mensagens e envio automático utilizando Selenium WebDriver.

# O sistema foi estruturado em módulos separados para melhor manutenção e organização:
-functionsA.py: núcleo da automação e integração com o Google Sheets.
-gui.py: interface gráfica e controle de execução.
-main.py: ponto de entrada principal do programa.

🧩 Funcionalidades
-Interface intuitiva criada com Tkinter.
-Envio automatizado de mensagens com base em dados da planilha.
-Integração completa com o Google Sheets (via API).
-Suporte a controle de status e registro de data de envio.
-Uso de Selenium para simular interação humana no navegador.
-Logs em tempo real diretamente na interface.
-Efeitos sonoros e interações visuais com pygame.

⚙️ Tecnologias utilizadas
-Python 3
-Tkinter (interface gráfica)
-Selenium WebDriver
-gspread e Google Sheets API
-oauth2client
-pygame (efeitos sonoros)
-webdriver_manager
-requests


🧰 Objetivo do projeto
Este projeto foi desenvolvido com o propósito de automatizar atendimentos repetitivos em uma plataforma de mensagens, conectando dados da planilha à execução prática via browser automatizado.
Também serve como um exemplo de integração entre automação web e interface desktop em Python.

⚠️ Observação
As credenciais, URLs e dados de login devem ser ocultadas para segurança.
Antes de publicar, substitua essas informações por placeholders e adicione o arquivo .env ao .gitignore.

📄 Licença
Este projeto é de uso livre para fins educacionais e demonstração de automação Python com interface gráfica.
