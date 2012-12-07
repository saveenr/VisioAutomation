$app = Get-WmiObject -Class Win32_Product -Filter "Name = 'Visio PowerShell Module'"
if ($app -eq $null)
{
    Write-Host "VisioPS module is not installed"
}
else
{
    Write-Host "VisioPS module is installed"
    Write-Host "Uninstalling now"
    $app.Uninstall()
    Write-Host "Finished Uninstalling"
    
}