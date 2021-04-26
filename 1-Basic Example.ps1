# related Blog Post https://www.techguy.at/control-pipedrive-crm-with-powershell-start-and-connect/


$PipeDrive_APIKey="3e3c33c3a3d3333dcd333ab3c333333c33b33fc"
$Pipedrive_domain="company"

$Pipedrive_BaseURL="https://$Pipedrive_domain.pipedrive.com/v1/"




# Get All Persons
#https://developers.pipedrive.com/docs/api/v1/#!/Persons/getPersons
#https://company.pipedrive.com/v1/persons?start=0&api_token=3e3c33c3a3d3333dcd333ab3c333333c33b33fc
$Url=$Pipedrive_BaseURL+"persons?start=0&api_token=$PipeDriveAPI"

$Result=Invoke-RestMethod -uri $URL -Method GET 

$Result.data.count


