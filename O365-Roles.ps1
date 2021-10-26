<#
	O365-Roles.ps1
	Created By - Kristopher Roy
	Created On - October 26 2021
	Modified On - 

	This Script gathers information on idividual User Roles in O365 and then generates a report
#>

#Connect to O365
Connect-MsolService

#Gather all Users and attached Roles
$AllUSerRolls = Get-MsolRole | %{$role = $_.name; Get-MsolRoleMember -RoleObjectId $_.objectid} | select @{Name="Role"; Expression = {$role}}, DisplayName, EmailAddress

#Create a Unique list of just the Users
$Users = $AllUSerRolls|select DisplayName,EmailAddress -Unique

#Create a Unique List of just the Roles
$Roles = $AllUSerRolls|select Role -Unique
$RoleCounts = $Roles.count

#Add a field to the Users list for each Role
    FOREACH($Role in $Roles)
    {
        $Users|Add-Member -NotePropertyName ($Role.Role).ToString() -NotePropertyValue "" 
    }

#Disignate a value to indicate which role each individual user has
ForEach($User in $Users)
{
   $USERRoles = $AllUSerRolls|where{$_.DisplayName -eq $User.DisplayName}
   FOREACH($URole in $USERRoles)
   {
    $User.($URole.Role) = "X"
   }
   $User

}

#Export the list to a CSV
$Users|export-csv c:\projects\gtil\MSO365Roles.csv -NoTypeInformation




