cls

$scriptPath = split-path -parent $MyInvocation.MyCommand.Definition
. $scriptPath\config.ps1
Install-Module ImportExcel -scope CurrentUser
$global:authHeaders = ""


function CreateAuthorizationHeader()
{
    $pair = "$($zendesk_user_name):$($zendesk_password)"
    $encodedCreds = [System.Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes($pair))
    $basicAuthValue = "Basic $encodedCreds"
    $global:authHeaders = @{
        Authorization = $basicAuthValue
    }
}


function GetOrganisations()
{
    $organizations = @()
    $list_organizations = Invoke-RestMethod -Uri "$zendesk_address/api/v2/organizations.json?per_page=100" -Headers $authHeaders
    $organizations += $list_organizations.organizations
    while ($list_organizations.next_page -ne $null)
    {
        $list_organizations = Invoke-RestMethod -Uri $list_organizations.next_page -Headers $authHeaders        
        $organizations += $list_organizations.organizations
    }
    Write-Host "Found" $organizations.count "organizations" -ForegroundColor Yellow
    return $organizations
}


function GetTickets()
{
    $tickets = @()
    $tickets_list = Invoke-RestMethod -Uri "$zendesk_address/api/v2/tickets.json?per_page=100" -Headers $authHeaders
    $tickets += $tickets_list.tickets
    while ($tickets_list.next_page -ne $null)
    {
        $tickets_list = Invoke-RestMethod -Uri $tickets_list.next_page -Headers $authHeaders        
        $tickets += $tickets_list.tickets
    }
    Write-Host "Found" $tickets.count "tickets" -ForegroundColor Yellow
    return $tickets
}


function GetAudits()
{
    $audits = @()
    $audits_list = Invoke-RestMethod -Uri "$zendesk_address/api/v2/ticket_audits.json?per_page=100" -Headers $authHeaders
    $audits += $audits_list.audits
    while ($audits_list.before_url -ne $null)
    {
        $audits_list = Invoke-RestMethod -Uri $audits_list.before_url -Headers $authHeaders        
        $audits += $audits_list.audits
    }
    Write-Host "Found" $audits.count "audits" -ForegroundColor Yellow
    return $audits
}

function GetOrganisationFields()
{
    $organization_fields = @()
    $list_organization_fields = Invoke-RestMethod -Uri "$zendesk_address/api/v2/organization_fields.json?per_page=100" -Headers $authHeaders
    $organization_fields += $list_organization_fields.organization_fields
    while ($list_organization_fields.next_page -ne $null)
    {
        $list_organization_fields = Invoke-RestMethod -Uri $list_organization_fields.next_page -Headers $authHeaders        
        $organization_fields += $list_organizations.organization_fields
    }
    Write-Host "Found" $organization_fields.count "organization fields" -ForegroundColor Yellow
    return $organization_fields
}




function execute(){
    [Net.ServicePointManager]::SecurityProtocol = "tls12, tls11, tls"
    CreateAuthorizationHeader
    
    GetOrganisations | Export-Excel -Path $scriptPath\Organisations.xlsx
    GetOrganisationFields | Export-Excel -Path $scriptPath\OrganisationFields.xlsx
    GetTickets | Export-Excel -Path $scriptPath\Tickets.xlsx
    GetAudits | Export-Excel -Path $scriptPath\Audits.xlsx
}

execute