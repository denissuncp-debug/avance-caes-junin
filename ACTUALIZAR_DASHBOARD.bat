@echo off
echo INICIANDO ACTUALIZACION DEL DASHBOARD CAES...
cd /d "C:\Users\denis\OneDrive - continental.edu.pe\Escritorio\Dashboard-CAES"

echo 1. Ejecutando Robot en SIGO...
C:\Users\denis\AppData\Local\Python\pythoncore-3.14-64\python.exe automatizador_sigo_pae_junin.py

echo 2. Subiendo a Internet...
git add index.html
git commit -m "Actualizacion manual desde PC"
git push origin main

echo PROCESO COMPLETADO!
pause