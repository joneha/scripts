# ///////////////////////////////////////////////////////
# Logging in to client 365 tenant and modify junk-config to (hopefully) disable spam-filtering
# Last revised on 26/11/2020.
# ///////////////////////////////////////////////////////

# Connect to client tenant
Connect-ExchangeOnline -UserPrincipalName john.doe@example.com -DelegatedOrganization contoso.onmicrosoft.com
# Verify current ruleset
Get-MailboxJunkEmailConfiguration -Identity "bob@contoso.com"
# Disable junk-email rule
Set-MailboxJunkEmailConfiguration -Identity "bob@contoso.com" -Enabled $false
# Disconnect
Disconnect-ExchangeOnline

# ///////////////////////////////////////////////////////

# Connect to client tenant
Connect-ExchangeOnline -UserPrincipalName john.doe@example.com -DelegatedOrganization contoso.onmicrosoft.com
# Get current config (doubles as verification that there are no typos)
Get-MailboxJunkEmailConfiguration -Identity "bob@contoso.com"
# Enable junk-email rule
Set-MailboxJunkEmailConfiguration -Identity "bob@contoso.com" -Enabled $true
# Disconnect
Disconnect-ExchangeOnline
