Import-Module ActiveDirectory

# Get data from Active Directory
$students = Get-ADUser -Filter * -SearchBase 'OU=Schueler,OU=Benutzer,DC=musterschule,DC=schule,DC=paedml' -Properties BusinessCategory,Department,DepartmentNumber,memberOf |
				Where-Object { $_.sAMAccountName -notmatch '^ka\.' -and $_.sAMAccountName -notmatch '^profschueler' }
$groupStudents = 'G_Schueler'

# Fix the group memberships
$students |
	ForEach-Object {
		$groupSchool = 'G_Schueler_' + $_.BusinessCategory[0]
		$groupClass = $groupSchool + '_' + $_.Department + '_' + $_.DepartmentNumber[0]
		$groupOctoClass = 'OCTO_' + $_.BusinessCategory[0] + '_' + $_.Department
		Add-ADGroupMember -Identity $groupStudents -Members $_
		Add-ADGroupMember -Identity $groupSchool -Members $_
		Add-ADGroupMember -Identity $groupClass -Members $_
		Add-ADGroupMember -Identity $groupOctoClass -Members $_
	}