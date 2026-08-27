# 📦 Sistema Integrado de Gestão de Estoque (Serverless)

![JavaScript](https://img.shields.io/badge/JavaScript-F7DF1E?style=for-the-badge&logo=javascript&logoColor=black)
![HTML5](https://img.shields.io/badge/HTML5-E34F26?style=for-the-badge&logo=html5&logoColor=white)
![Google Apps Script](https://img.shields.io/badge/Google_Apps_Script-4285F4?style=for-the-badge&logo=google&logoColor=white)
![Chart.js](https://img.shields.io/badge/Chart.js-FF6384?style=for-the-badge&logo=chartdotjs&logoColor=white)

## 📌 O Problema de Negócio
Ambientes de infraestrutura de TI lidam diariamente com um alto volume de movimentação de hardwares e insumos (cabos, conectores, peças de reposição). O controle manual e descentralizado gera gargalos críticos: indisponibilidade de materiais no momento de manutenções urgentes e falhas na comunicação com o setor de compras.

## 💡 A Solução
Desenvolvi uma aplicação web *Full-Cycle* e *Serverless* para resolver essa dor na base da operação. O sistema gerencia todo o ciclo de vida dos insumos: desde o cadastro e autenticação de operadores, até o registro de entradas/saídas em bancada e o disparo automatizado de alertas de reposição.

### Principais Funcionalidades:
*   **Controle de Acesso (RBAC):** Sistema de login com diferenciação de permissões (Operador vs. Administrador) e criptografia de senhas utilizando Hash SHA-256.
*   **Gestão de Compras:** Fluxo completo de pedidos com status dinâmicos (Pendente, Recebido Parcial, Concluído ou Cancelado).
*   **Dashboard Analítico:** Renderização de métricas de consumo e custo médio diário utilizando `Chart.js`, auxiliando na tomada de decisão financeira e previsão de demandas.
*   **Alertas Automatizados:** Integração com a API do Gmail para notificar a equipe proativamente quando um insumo atinge níveis críticos ou baixos de estoque.

## 🛠️ Arquitetura do Sistema
O projeto foi construído utilizando o ecossistema do Google Workspace como backend e banco de dados, reduzindo custos de infraestrutura a zero.

*   **Frontend (Client-Side):** `HTML5`, `CSS3`, `JavaScript Vanilla` (Foco em performance e ausência de dependências pesadas).
*   **Backend (Server-Side):** `Google Apps Script` processando as requisições, executando a lógica de negócio e as rotas.
*   **Banco de Dados:** `Google Sheets` atuando de forma estruturada (tabelas relacionais lógicas: Resumo, Cadastro, RH, Configurações e Históricos).
*   **Mensageria:** `MailApp API` para automação de e-mails de alerta.

---

## 🚀 Como testar este projeto (Ambiente de Avaliação)

Preparei um ambiente com dados *mockados* para facilitar a avaliação deste projeto. Siga os passos abaixo para instanciar o seu próprio servidor de testes em menos de 2 minutos:

1. **Crie sua base de dados:**
   Clique no link abaixo para gerar uma cópia automática da estrutura do banco de dados no seu Google Drive (não se preocupe, isso não afetará o projeto original):
   > [👉 Gerar Cópia da Planilha Template](https://docs.google.com/spreadsheets/d/1f7zPEqAKPOpdERt5GRH44ey7TGx079Og6EtE-hyb3co/copy)

2. **Ative o Backend:**
   * Na sua nova planilha, vá no menu superior e clique em `Extensões > Apps Script`.
   * Note que os arquivos `Code.gs` e `Index.html` já estarão lá.
   * No canto superior direito, clique em **Implantar (Deploy) > Nova Implantação**.
   * Selecione o tipo **App da Web (Web App)**.
   * Em "Executar como", deixe sua conta. Em "Quem pode acessar", selecione "Qualquer pessoa".
   * Clique em Implantar e autorize as permissões do Google (é necessário para o disparo dos e-mails de teste e manipular a planilha).

3. **Acesse o Sistema:**
   Abra a URL gerada pelo Google. Para entrar e testar as funcionalidades administrativas e os gráficos, utilize as credenciais de teste abaixo:
   * **Login:** `admin`
   * **Senha:** `123456`

---

## 👨‍💻 Autor
**Roger Schiavon**
*Estudante de Análise e Desenvolvimento de Sistemas*
