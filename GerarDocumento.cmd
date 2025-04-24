@echo off

REM Definir o diretório base do projeto
set BASE_DIR=%~dp0

REM Ativar o ambiente Python (caso necessário)
REM activate <nome_do_seu_ambiente>

REM Executar o script Python para gerar os certificados
python "%BASE_DIR%gerar_documentos.py"

REM Mensagem de conclusão
echo Certificados gerados com sucesso!
pause
