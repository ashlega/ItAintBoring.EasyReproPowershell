Import-Module .\EasyRepro\EasyRepro.psm1 -Force

#Load settings - comment out if using in pipelines
.\Settings.ps1

$testResults = @{}
$testingFailed = $false

function Invoke-CreateAccount
{
    $testName = "Create Account"
    
	try{
		Open-XrmApp -appName "Sales Hub"
		Open-XrmAppSubArea "Sales" "Accounts"
		Submit-XrmCommand "New"
		Set-XrmValue "123name123" "Account #1234567"
		Submit-XrmForm
		
		$testResults.Add($testName, "Succeded")
	}
	catch
	{
	    $script:testingFailed = $true
	    $testResults.Add($testName, "Failed")
	}
	
}

function Invoke-XrmTests
{
    Invoke-CreateAccount
}

try{
   Start-XrmTesting
   Invoke-XrmTests
}
finally{
   Clear-XrmObjects
}

write-host ""
write-host "Testing Summary:"
write-host ""
foreach($key in $testResults.Keys)
{
   write-host "${key}: $($testResults[$key])"
}

write-host $testingFailed
if(($testingFailed -eq $true) -and ($env:ThrowTestError -ne $null)) { throw "Testing failed!" }
#$null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')

