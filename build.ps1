$exclude = @("venv", "Contoso_Faturas.zip")
$files = Get-ChildItem -Path . -Exclude $exclude
Compress-Archive -Path $files -DestinationPath "Contoso_Faturas.zip" -Force