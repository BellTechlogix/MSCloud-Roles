<#
	AzureAD-Roles.ps1
	Created By - Kristopher Roy
	Created On - Nov 10 2021
	Modified On - 

	This Script gathers information on idividual User Roles in AzureAD Tenets and then generates a report
#>

#Timestamp
$runtime = Get-Date -Format "yyyyMM"

#folder to store completed reports
$rptfolder = "C:\Reports\GTIL\"

#Get AzureAD Creds
$AzureAdCred = Get-Credential

#List of Directories to pull from
$Directories =  @('1f05524b-c860-46e3-a2e8-4072506b3e4a','b009e798-8069-46d1-97b5-ad8a2010e6fc','5cfa8de4-c4b0-4db7-bd10-acd213d42419')

FOREACH($Dir in $Directories)
{
    #Connect to Directory
    Connect-AzureAD -Credential $AzureAdCred -TenantId $Dir
    
    $membersarray =@()
    
    $Tenet = Get-AzureADTenantDetail
    $roles = Get-AzureADDirectoryRole
    FOREACH($role in $Roles)
    {
        $members = Get-AzureADDirectoryRoleMember -ObjectId $role.ObjectID|select DisplayName,Mail,Role,Tenet,TenetID,TenetDomain
        #Get-AzureADDirectoryRoleMember -ObjectId $role.ObjectID|select *
        FOREACH($member in $members)
        {
            $member.Role = $role.DisplayName
            $member.Tenet = $Tenet.DisplayName
            $member.TenetID = $Tenet.ObjectId
            $member.TenetDomain = $Tenet.VerifiedDomains.Name[1]
            $member
            $membersarray += $member
        }
    }
    
    #Gather all Users and attached Roles
    $AllUSerRolls = $membersarray

    #Create a Unique list of just the Users
    $Users = $AllUSerRolls|select DisplayName,Mail -Unique

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

    $filename = $Tenet.DisplayName
	$Users|export-csv $rptFolder$runtime-$filename-MSO365Roles.csv -NoTypeInformation

    $membersarray = $null
    $Users = $null
    $AllUSerRolls = $null
    $Roles = $null

}
