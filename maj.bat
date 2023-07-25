echo Cette fenêtre est en train de mettre a jour le document excel du planning des travaux.
echo Ne pas toucher a l'ordinateur tant qu'excel n'est pas ouvert et que cette fenêtre n'est pas fermee. 
title Mise a jour du document Excel
chcp 863
taskkill -f -im "EXCEL.EXE" 
timeout 1
del "%userprofile%\Downloads\Test planning.xlsm"
timeout 1 
start "" "https://totalworkplace.sharepoint.com/sites/Runionpermistravaux/Documents partages/Reunion de coordination/Test Planning.xlsm"
timeout 20
color a
echo Fichier de maj par Léonard :) 
start "" "%userprofile%\Downloads\Test planning.xlsm" 
