Import-Module ActiveDirectory

$students = Get-ADUser -Filter * -SearchBase 'OU=Schueler,OU=Benutzer,DC=musterschule,DC=schule,DC=paedml' -Properties BusinessCategory,Department,DepartmentNumber,memberOf |
				Where-Object { $_.sAMAccountName -notmatch '^ka\.' -and $_.sAMAccountName -notmatch '^profschueler' }

$groupStudents = 'G_Schueler'
$groupOcto = 'OCTO_'

$students | ForEach-Object {
		$groupSchool = 'G_Schueler_' + $_.BusinessCategory[0]
		$classGroup = $groupSchool + '_' + $_.Department + '_' + $_.DepartmentNumber[0]
		$classOctoGroup = 'OCTO_' + $_.BusinessCategory[0] + '_' + $_.Department
		Add-ADGroupMember -Identity $groupPupils -Members $_
		Add-ADGroupMember -Identity $groupSchool -Members $_
		Add-ADGroupMember -Identity $classGroup -Members $_
		Add-ADGroupMember -Identity $classOctoGroup -Members $_
}