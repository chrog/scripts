Import-Module ActiveDirectory

$pfad = 'D:\Export' # ohne Backslash am Ende!
$date = Get-Date -UFormat "%d.%m.%Y"

# Daten aus AD holen
$lehrer = Get-ADUser -Filter {Name -like "*" -and SamAccountName -ne 'proflehrer'} -SearchBase "OU=Lehrer,OU=Benutzer,DC=musterschule,DC=schule,DC=paedml" -Properties Company |
	      Select-Object @{Name="Schulart";Expression={$_.Company}},@{Name="Nachname";Expression={$_.Surname}},@{Name="Vorname";Expression={$_.GivenName}},@{Name="Benutzername";Expression={$_.SamAccountName}} |
          Sort-Object Schulart,Nachname,Vorname,Benutzername
$schueler = Get-ADUser -Filter {Name -like "*" -and SamAccountName -ne 'profschueler'} -SearchBase "OU=Schueler,OU=Benutzer,DC=musterschule,DC=schule,DC=paedml" -Properties Company,Department |
	        Select-Object @{Name="Schulart";Expression={$_.Company}},@{Name="Klasse";Expression={$_.Department}},@{Name="Nachname";Expression={$_.Surname}},@{Name="Vorname";Expression={$_.GivenName}},@{Name="Benutzername";Expression={$_.SamAccountName}} |
            Sort-Object Schulart,Klasse,Nachname,Vorname,Benutzername

# CSV-Export: Lehrer und Schueler
$lehrer | ConvertTo-Csv -Delimiter ';' -NoTypeInformation | % { $_ -replace '"', ''} | Set-Content -Encoding UTF8  "$pfad\Export_Lehrer.txt"
$schueler | ConvertTo-Csv -Delimiter ';' -NoTypeInformation | % { $_ -replace '"', ''} | Set-Content -Encoding UTF8  "$pfad\Export_Schueler.txt"

# HTML-Export: Lehrer und Schueler
$lehrer | ConvertTo-Html -Property Schulart,Nachname,Vorname,Benutzername -Title 'Lehrerliste' -PreContent "<h1>Lehrerliste ($date)</h1>" |
          ForEach-Object -Process { $_.Replace('<table>','<table border>') } |
          Out-File -FilePath "$pfad\Export_Lehrer.html"
$schueler | ConvertTo-Html -Property Schulart,Klasse,Nachname,Vorname,Benutzername -Title 'Schuelerliste' -PreContent "<h1>Schuelerliste ($date)</h1>" |
            ForEach-Object -Process { $_.Replace('<table>','<table border>') } |
            Out-File -FilePath "$pfad\Export_Schueler.html"

# HTML-Export: Klassen
$klassen = $schueler | Select-Object Klasse -Unique
foreach ($klasse in $klassen)
{
    $tmp = $klasse.Klasse
    $schueler | Where-Object -FilterScript { $_.Klasse -eq $tmp } |
                ConvertTo-Html -Property Schulart,Klasse,Nachname,Vorname,Benutzername -Title 'Schuelerliste' -PreContent "<h1>Klassenliste $tmp ($date)</h1>" |
                ForEach-Object -Process { $_.Replace('<table>','<table border>') } |
                Out-File -FilePath "$pfad\Export_Klassenliste_$tmp.html"            
}
