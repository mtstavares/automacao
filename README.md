# 🔐 Automação de Validação de Credenciais e Coleta de Dados

Este projeto automatiza a **validação de credenciais de usuários** e a **coleta de dados funcionais**, utilizando uma planilha Excel como entrada e gerando resultados atualizados com logs detalhados. Ele é usado para validar credenciais coletadas por meio de OSINT e suas respectivas validações em sistemas institucionais.

🔒 **Observação de Segurança**
- Este projeto interage com sistemas institucionais internos e utiliza APIs privadas. Seu uso é restrito a ambientes autorizados, seguindo políticas rígidas de privacidade e segurança da informação. URLs dos sistemas e APIs foram mascaradas neste projeto público para preservar a confidencialidade.

## 🚀 Funcionalidades

- **Consulta de dados funcionais** (CPF, Nome Completo e E-mail Institucional) a partir de:
  - Registro de Empregado (RE)
  - E-mail institucional
- **Testes automáticos de login**:
  - Sistemas que utilizam senhas do Módulo de Segurança 
  - Sistemas que utilizam senhas do Active Directory
- **Atualização em lote da planilha Excel**:
  - Marca visual (verde/vermelho) para indicar sucesso e falha.
  - Mensagens detalhadas para cada tentativa de autenticação.
- **Geração de logs** com contadores de sucesso/falha e detalhes de erros.
- **Limpeza automática** de arquivos intermediários após a execução.

## ⚙️ Tecnologias Utilizadas

- **Python 3**
- requests: Integração com APIs
- subprocess: Consultas LDAP com dsquery
- selenium: Testes automatizados de login
- openpyxl: Manipulação de arquivos Excel
- urllib3: Gerenciamento de conexões HTTP

## 🗂️ Estrutura dos Arquivos

- `Credenciais.xlsx`: Planilha original com as credenciais.
- `Logs_BuscarNome.txt`: Log da coleta de dados (CPF, Nome, E-mail).
- `Logs_testarCredenciais.txt`: Log dos testes de login.
- `Credenciais_<data>.xlsx`: Resultado final com marcações e status.

## 📈 Fluxo de Funcionamento

1️⃣ **Coleta de Dados**
   - Lê a planilha `Credenciais.xlsx`.
   - Consulta a API usando o RE ou o e-mail institucional.
   - Preenche as colunas CPF, Nome Completo e E-mail Funcional.

2️⃣ **Testes de Login**
   - Utiliza Selenium para automatizar o login nos sistemas MS e AD.
   - Atualiza a planilha com:
     - Status de sucesso/falha.
     - Mensagens detalhadas sobre o resultado do login.

3️⃣ **Finalização**
   - Salva a planilha atualizada com data/hora no nome.
   - Registra logs completos da execução.
   - Remove arquivos temporários.

