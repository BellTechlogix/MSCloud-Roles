#
# O365-Roles.ps1
#

Connect-MsolService

$AllUSerRolls = Get-MsolRole | %{$role = $_.name; Get-MsolRoleMember -RoleObjectId $_.objectid} | select @{Name="Role"; Expression = {$role}}, DisplayName, EmailAddress
$Users = $AllUSerRolls|select DisplayName,EmailAddress -Unique
$Roles = $AllUSerRolls|select Role -Unique
$RoleCounts = $Roles.count

    FOREACH($Role in $Roles)
    {
        $Users|Add-Member -NotePropertyName ($Role.Role).ToString() -NotePropertyValue "" 
    }

ForEach($User in $Users)
{
   $USERRoles = $AllUSerRolls|where{$_.DisplayName -eq $User.DisplayName}
   FOREACH($URole in $USERRoles)
   {
    $User.($URole.Role) = "X"
   }
   $User

}

$Users|export-csv c:\projects\gtil\MSO365Roles.csv -NoTypeInformation




