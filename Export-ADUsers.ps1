Import-Module ActiveDirectory

$path = 'D:\Export'
$date = Get-Date -UFormat "%d.%m.%Y"

# Get data from Active Directory
$teachers = Get-ADUser -Filter {Name -like "*" -and sAMAccountName -ne 'proflehrer'} -SearchBase "OU=Lehrer,OU=Benutzer,DC=musterschule,DC=schule,DC=paedml" -Properties Company |
				Select-Object @{Name="Schulart";Expression={$_.Company}},@{Name="Nachname";Expression={$_.Surname}},@{Name="Vorname";Expression={$_.GivenName}},@{Name="Benutzername";Expression={$_.sAMAccountName}} |
					Sort-Object Schulart,Nachname,Vorname,Benutzername
$students = Get-ADUser -Filter {Name -like "*" -and sAMAccountName -ne 'profschueler'} -SearchBase "OU=Schueler,OU=Benutzer,DC=musterschule,DC=schule,DC=paedml" -Properties Company,Department |
				Select-Object @{Name="Schulart";Expression={$_.Company}},@{Name="Klasse";Expression={$_.Department}},@{Name="Nachname";Expression={$_.Surname}},@{Name="Vorname";Expression={$_.GivenName}},@{Name="Benutzername";Expression={$_.sAMAccountName}} |
					Sort-Object Schulart,Klasse,Nachname,Vorname,Benutzername

# Export all teachers/students to csv files
$teachers | ConvertTo-Csv -Delimiter ';' -NoTypeInformation | % { $_ -replace '"', ''} | Set-Content -Encoding UTF8  "$path\Export_Lehrer.txt"
$students | ConvertTo-Csv -Delimiter ';' -NoTypeInformation | % { $_ -replace '"', ''} | Set-Content -Encoding UTF8  "$path\Export_Schueler.txt"

# Export all teachers/students to html files
$teachers | ConvertTo-Html -Property Schulart,Nachname,Vorname,Benutzername -Title 'Lehrerliste' -PreContent "<h1>Lehrerliste ($date)</h1>" |
				ForEach-Object -Process { $_.Replace('<table>','<table border>') } |
					Out-File -FilePath "$path\Export_Lehrer.html"
$students | ConvertTo-Html -Property Schulart,Klasse,Nachname,Vorname,Benutzername -Title 'Schuelerliste' -PreContent "<h1>Schuelerliste ($date)</h1>" |
				ForEach-Object -Process { $_.Replace('<table>','<table border>') } |
					Out-File -FilePath "$path\Export_Schueler.html"

# Export all classes to html files
$classes = $students | Select-Object Klasse -Unique
foreach ($class in $classes)
{
    $tmp = $class.Klasse
    $students | Where-Object -FilterScript { $_.Klasse -eq $tmp } |
					ConvertTo-Html -Property Schulart,Klasse,Nachname,Vorname,Benutzername -Title 'Schuelerliste' -PreContent "<h1>Klassenliste $tmp ($date)</h1>" |
						ForEach-Object -Process { $_.Replace('<table>','<table border>') } |
							Out-File -FilePath "$path\Export_Klassenliste_$tmp.html"            
}