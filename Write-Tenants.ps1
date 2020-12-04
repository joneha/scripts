# ///////////////////////////////////////////////////////
# Simple script to fetch and list all tenants connected to a Partner-center account
# 
# ///////////////////////////////////////////////////////

Connect-MsolService
$customers = Get-MsolPartnerContract -All

foreach ($customer in $customers) {
Write-Host "Tenant ID: $($customer.TenantID) `r`nKunde: $($customer.Name)`r`n"
}