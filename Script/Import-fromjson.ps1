
$Policy = get-content -Path "C:\Users\GBATTEUX\Desktop\Conditionalaccess\json\f42ba384-96bc-4169-b508-139ff7f1cebc.json"
$apiUrl = 'https://graph.microsoft.com/beta/conditionalAccess/policies'
Post-MsGraph -AccessToken $AccessToken -Uri $apiUrl -Body $Policy