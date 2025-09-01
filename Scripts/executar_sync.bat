@echo off
REM Define o título da janela do console para algo descritivo
TITLE Sincronizador de Atendimentos

REM Muda para o diretório onde o script Python está localizado.
REM IMPORTANTE: Verifique se este caminho está 100% correto.
cd /d "T:\APF\Inteligência de Mercado\Códigos Python\Projeto Tarefas Focais"

echo.
echo ========================================================
echo   Iniciando o script de sincronizacao de atendimentos...
echo ========================================================
echo.

REM Executa o script Python.
REM O Python precisa estar instalado e configurado no PATH do sistema.
python sync_atendimentos.py
