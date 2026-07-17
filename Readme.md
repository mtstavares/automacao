# CredenciaisFullV3

Sistema em Python para processar planilhas de credenciais, resolver identidades, testar acessos em sistemas internos e gerar relatórios consolidados.

## Para que este projeto foi feito

Este projeto foi criado para automatizar a análise de credenciais obtidas através de ferramentas de OSINT. O objetivo é aumentar a celeridade do processo de identificação de credenciais com acesso ao ambiente corporativo que foram vazadas, e bloquear acessos ao ambiente interno, reduzindo possibilidades de ataques.


## Problema que o projeto resolve

Em análises manuais de credenciais, normalmente é necessário:

- conferir CPF, RE ou e-mail;
- consultar sistemas internos para descobrir nome e e-mail funcional;
- verificar se a senha ainda concede acesso;
- separar acessos válidos de acessos inválidos;
- evitar duplicidade em relatórios mensais;
- manter histórico de identificações;
- produzir logs para auditoria.

Esse processo é repetitivo, sujeito a erros e pode consumir muito tempo quando há muitas linhas na planilha.

O `ConcultaCredenciais` automatiza esse fluxo, padroniza os resultados e gera arquivos prontos para consulta e acompanhamento.

## Principais funcionalidades

- Leitura da planilha com dados brutos colhidos via OSINT.
- Normalização de CPF e RE.
- Resolução de identidade por CPF, RE ou e-mail.
- Consulta a APIs internas para buscar CPF, nome, e-mail e situação legal.
- Busca complementar no Active Directory via ferramentas de domínio.
- Teste automatizado de login nos sistemas MS e AD com Selenium.
- Retentativa automática em caso de erro técnico ou resultado inconclusivo.
- Cache de autenticação por `CPF + senha + sistema` durante a execução.
- Cache de situação legal por CPF durante a execução.
- Remoção de linhas sem dados testáveis.
- Geração de planilha individual de resultado.
- Atualização de planilha mensal consolidada.
- Preservação de campos manuais na planilha mensal.
- Logs separados para resumo de execução e auditoria técnica.

## Tecnologias usadas

- Python
- OpenPyXL
- Selenium WebDriver
- Requests
- PyInstaller
- Active Directory / dsquery
- APIs REST internas
- Microsoft Excel

## Resumo

O `ConsultaCredenciais` automatiza o processo de identificação e validação de credenciais, reduzindo esforço manual, padronizando resultados, aumentando a celeridade em processos de cibersegurança e gerando relatórios mensais úteis para acompanhamento e auditoria.

