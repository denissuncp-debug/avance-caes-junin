@echo off
echo INICIANDO ACTUALIZACION DEL DASHBOARD CAES...
cd /d "C:\Users\denis\Documentos\GitHub\avance-caes-junin"

echo 1. Ejecutando Robot en SIGO...
C:\Users\denis\AppData\Local\Python\pythoncore-3.14-64\python.exe automatizador_sigo_pae_junin.py

echo PROCESO DE DESCARGA COMPLETADO! 
echo Por favor, abre GitHub Desktop para hacer el Commit y Push.
pause