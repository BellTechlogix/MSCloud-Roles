<#
	AzureAD-Roles.ps1
	Created By - Kristopher Roy
	Created On - Nov 10 2021
	Modified On - 

	This Script gathers information on idividual User Roles in AzureAD Tenets and then generates a report
#>

#Get AzureAD Creds
$AzureAdCred = Get-Credential

#List of Directories to pull from
$Directories =  @('1f05524b-c860-46e3-a2e8-4072506b3e4a','b009e798-8069-46d1-97b5-ad8a2010e6fc','5cfa8de4-c4b0-4db7-bd10-acd213d42419')

#Create Array for storing each member
$membersarray =@()

#Go through each Tenet/Directory and pull roles and members
FOREACH($Dir in $Directories)
{
    
    #Connect to Directory
    Connect-AzureAD -Credential $AzureAdCred -TenantId $Dir

    $Tenet = Get-AzureADTenantDetail
    $roles = Get-AzureADDirectoryRole
    FOREACH($role in $Roles)
    {
        $members = Get-AzureADDirectoryRoleMember -ObjectId $role.ObjectID|select DisplayName,Role,Tenet,TenetID,TenetDomain
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
}
