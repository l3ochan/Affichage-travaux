echo Cette fenêtre est en train de mettre a jour le document excel du planning des travaux.
echo Ne pas toucher a l'ordinateur tant qu'excel n'est pas ouvert et que cette fenêtre n'est pas fermee. 
title Mise a jour du document Excel
taskkill -f -im "WINWORD.EXE" 
timeout 1
del "%userprofile%\Downloads\test.docx"
timeout 1 
powershell -Command "Invoke-WebRequest -uri "https://cloud.nekocorp.fr/index.php/s/eiyK2Eg2MXGR5Az/download/test.docx" -OutFile "$env:userprofile\Downloads\test.docx" 
timeout 3
echo Fichier de maj par Léonard :) 
start "" "%userprofile%\Downloads\test.docx" 
