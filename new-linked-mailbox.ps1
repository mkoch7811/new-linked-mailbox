<#
.SYNOPSIS
    NewLinkedMailbox - Creates a linked mailbox for accounts with MailboxPlease in the info attribute
    Created by Mike Koch on October 29, 2019
.DESCRIPTION
    Queries USERDOMAIN for user accounts that have 'MailboxPlease' in the 'info' attribute
        - this attribute is populated by the regional IT teams, which have been delegated rights to create user accounts in their own OUs
    Creates a matching account in the RESOURCEDOMAIN, in the appropriate region OU
    Sets the Company attribute on the mailbox account, to ensure that the proper email address policy is applied
    Creates the linked mailbox, then copies the primary smtp address back to the user's account in USERDOMAIN
    Clears 'MailboxPlease' from the info attribute
.NOTES
    Assumes the account running this script has sufficient rights to do the following:
        1. Read user properties in USERDOMAIN
        2. Create user accounts in RESOURCEDOMAIN
        3. Create linked mailboxes in the on-premises Exchange org
        4. Write user properties in USERDOMAIN (to write the primary smtp address back to the user's account)
#>
[CmdletBinding()]
Param()

$USERDOMAINDC = "dc1.USERDOMAIN.local"
$RESOURCEDOMAINDC = "dc1.RESOURCEDOMAIN.local"

###
# Query USERDOMAIN for enabled accounts containing "MailboxPlease" in the info attribute
###
$mbxrequests = @(Get-ADUser -Filter { enabled -eq $true -AND info -like "*MailboxPlease*" } -Server $USERDOMAINDC -SearchBase "ou=All Users,dc=USERDOMAIN,dc=local" -Properties givenname, sn, title, description, department, office, company, manager)

if (!$mbxrequests) {
    Write-Verbose "No mailbox requests detected."
}
else {
    foreach ($mbx in $mbxrequests) {
        Write-Verbose "Attempting mailbox creation for $($mbx.Name)..."
        $resAccountParms = @{ }
        $resAccountParms.Add("displayName", $mbx.Name)
        $resAccountParms.Add("userprincipalname", "$($mbx.SamAccountName)@COMPANYNAME.com")
        switch -Wildcard ($mbx.DistinguishedName) {
            "*Australia Associates*" { 
                $path = "ou=Australia,ou=LinkedMailboxAccounts,dc=RSEOURCEDOMAIN,dc=local"
                $resAccountParms.Add("company", "COMPANYNAME Australia") # required to trigger custom email address policy
            }
            "*Canada Associates*" { 
                $path = "ou=Canada,ou=LinkedMailboxAccounts,dc=RESOURCEDOMAIN,dc=local"
                $resAccountParms.Add("company", "COMPANYNAME Canada") # required to trigger custom email address policy
            }
            "*France Associates*" { 
                $path = "ou=France,ou=LinkedMailboxAccounts,dc=RESOURCEDOMAIN,dc=local"
                $resAccountParms.Add("company", "COMPANYNAME France") # required to trigger custom email address policy
            }
            "*Germany Associates*" { 
                $path = "ou=Germany,ou=LinkedMailboxAccounts,dc=RESOURCEDOMAIN,dc=local"
                if ($mbx.company) { $resAccountParms.Add("company", $mbx.company)}
            }
            "*Ireland Associates*" { 
                $path = "ou=Ireland,ou=LinkedMailboxAccounts,dc=RESOURCEDOMAIN,dc=local"
                if ($mbx.company) { $resAccountParms.Add("company", $mbx.company)}
            }
            "*Italy Associates*" {
                $path = "ou=Italy,ou=LinkedMailboxAccounts,dc=RESOURCEDOMAIN,dc=local"
                if ($mbx.company) { $resAccountParms.Add("company", $mbx.company)}
            }
            Default {
                $path = "ou=LinkedMailboxAccounts,dc=RESOURCEDOMAIN,dc=local"
                if ($mbx.company) { $resAccountParms.Add("company", $mbx.company)}
            }
        }
        if ($mbx.givenname) { $resAccountParms.Add("givenName", $mbx.GivenName) }
        if ($mbx.sn) { $resAccountParms.Add("sn", $mbx.sn) }
        if ($mbx.Department) { $resAccountParms.Add("department", $mbx.Department) }
        if ($mbx.description) { $resAccountParms.Add("description", $mbx.description) }
        if ($mbx.Title) { $resAccountParms.Add("title", $mbx.Title) }
        if ($mbx.office) { $resAccountParms.Add("physicalDeliveryOfficeName", $mbx.office) }
        $resAccountParms.Add("extensionAttribute1", "migrate.me")  # triggers migration script

        ### Let's see if we can locate this user's manager's mailbox account in RESOURCEDOMAIN
        if ($mbx.Manager) {
            $mgrsam = (Get-ADUser $mbx.Manager -Server $USERDOMAINDC).SamAccountName
            $resAccountParms.Add("Manager", (Get-ADUser $mgrsam -Server $RESOURCEDOMAINDC).DistinguishedName)
        }

        ### Make sure this samaccountname doesn't exist in RESOURCEDOMAIN
        if (!(Get-ADUser -Filter "samaccountname -eq '$($mbx.SamAccountName)'" -Server $RESOURCEDOMAINDC)) {
            Write-Verbose "Creating RESOURCEDOMAIN account for $($mbx.Name)..."
            New-ADUser -Name $mbx.Name -SamAccountName $mbx.SamAccountName -Enabled $FALSE -Path $path -Server $RESOURCEDOMAINDC -OtherAttributes $resAccountParms
            $resacct = Get-ADUser $mbx.SamAccountName -Server $RESOURCEDOMAINDC
            if ($resacct) {
                Write-Verbose "Creating linked mailbox for $($mbx.Name)..."
                $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://OnPremExchangeServer/powershell/ -Authentication Kerberos -AllowRedirection
                Import-PSSession $Session -CommandName Enable-Mailbox
                Enable-Mailbox -Identity $resacct.DistinguishedName -DomainController $RESOURCEDOMAINDC -Alias $resacct.SamAccountName -LinkedMasterAccount $mbx.DistinguishedName -LinkedDomainController $USERDOMAINDC
                Get-PSSession | Remove-PSSession

                ### copy email address to the USERDOMAIN account
                $email = (Get-ADUser $mbx.SamAccountName -Server $RESOURCEDOMAINDC -Properties mail).mail
                Set-ADUser $mbx.SamAccountName -Server $USERDOMAINDC -EmailAddress $email -Clear info
            } else {
                Write-Verbose "Creation of RESOURCEDOMAIN account failed $($mbx.SamAccountName)"
            }
        }
        else {
            Write-Verbose "Samaccountname already exists in RESOURCEDOMAIN ($($mbx.SamAccountName))."
        }
    }
}
Write-Verbose "Finished."
