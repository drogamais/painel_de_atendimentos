# Dashboard Painel_de_acessos

# Projeto Power BI - Conexão com Banco de Dados

Este projeto utiliza conexão via **ODBC**.

## Pré-requisitos
1. Instalar o driver ODBC apropriado (ex.: SQL Server, Oracle, MySQL).
2. Criar um **DSN de sistema** com o nome esperado pelo projeto.

## Configuração do DSN
- **Nome do DSN:** `dbSults` e `drogamais`(Aqui é arbitrário o nome)
- **Driver:** MySQL ODBC 9.4 Unicode Driver
- **Banco de dados:** `dbSults` e `drogamais`

⚠️ **Atenção:**  
- O **host, usuário e senha** **não estão neste repositório** por motivos de segurança.  
- Encontre essas informações na planilha READ ME.  
- Configure no Power BI Desktop a primeira vez que abrir o projeto — as credenciais ficam salvas localmente.  

## Primeira execução
1. Abra o Power BI Desktop.
2. Certifique-se que o DSN está configurado.
3. Ao abrir o relatório, forneça as credenciais.
4. O Power BI salvará as credenciais localmente (não são versionadas).
