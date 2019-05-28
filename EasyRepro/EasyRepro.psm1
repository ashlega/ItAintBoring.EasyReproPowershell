$script:XrmBrowser = $null
$script:XrmApp = $null
$script:WebClient = $null



function Start-XrmTesting{
	Register-EasyRepro
	
	$script:XrmUserName = $env:XrmUserName
    $script:XrmPassword = $env:XrmPassword
    $script:XrmUrl = $env:XrmUrl


	if($script:XrmPassword -eq $null)
	{
	    $script:XrmPassword = Read-Host -assecurestring "Please enter your password"
	}
	else
	{
	    $script:XrmPassword = Get-SecureString $script:XrmPassword
	}
	
    $script:XrmUserName = Get-SecureString $script:XrmUserName
    Initialize-XrmClient -url $script:XrmUrl -userName $script:XrmUserName -pwd $script:XrmPassword
}

function Clear-XrmObjects {
    if($script:XrmBrowser -ne $null){
      $script:XrmBrowser.Dispose()
	}
	if($script:XrmApp -ne $null){
      $script:XrmApp.Dispose()
	}
	if($script:WebClient -ne $null){
      #$script:WebClient.Dispose()
	}
}

function Open-XrmSubArea{
    Param(
      [Parameter(Mandatory = $true)]
      [String]
      $area,
	  [Parameter(Mandatory = $true)]
      [String]
	  $subArea,
	  [Parameter(Mandatory = $true)]
      [Int]
      $waitTime
	)
    $script:XrmBrowser.Navigation.OpenSubArea($area, $subArea)
	$script:XrmBrowser.ThinkTime($waitTime)
}

function Wait-ThinkTime{
    Param(
      [Parameter(Mandatory = $true)]
      [Int]
      $value
	)
    $script:XrmBrowser.ThinkTime($value)
}

function Register-EasyRepro{
  try
  {
	  Add-Type -Path EasyRepro\Newtonsoft.Json.dll
	  Add-Type -Path EasyRepro\WebDriver.dll
	  Add-Type -Path EasyRepro\WebDriver.Support.dll
	  Add-Type -Path EasyRepro\Microsoft.ApplicationInsights.dll
	  Add-Type -Path EasyRepro\Microsoft.Dynamics365.UIAutomation.Browser.dll
	  Add-Type -Path EasyRepro\Microsoft.Dynamics365.UIAutomation.Api.UCI.dll
	  Add-Type -Path EasyRepro\Microsoft.Dynamics365.UIAutomation.Api.dll
  }
  catch
  {
  }
}
 

function Get-SecureString{
    Param(
      [Parameter(Mandatory = $true)]
      [String]
      $value
    )
    [Microsoft.Dynamics365.UIAutomation.Browser.StringExtensions]::ToSecureString($value)
}

function Initialize-XrmBrowser {
    Param(
      [Parameter(Mandatory = $true)]
      [String]
      $url,
	  [Parameter(Mandatory = $true)]
      [PSObject]
      $userName,
	  [Parameter(Mandatory = $true)]
      [PSObject]
      $pwd
    )
	
	
    $options = [Microsoft.Dynamics365.UIAutomation.Browser.BrowserOptions]::New()
    $options.BrowserType = [Enum]::Parse([Microsoft.Dynamics365.UIAutomation.Browser.BrowserType], "Chrome")
    $options.PrivateMode = $true
    $options.FireEvents = $false
    $options.Headless = $false
    $options.UserAgent = $false
    $options.DefaultThinkTime = 2000
    $script:XrmBrowser = [Microsoft.Dynamics365.UIAutomation.Api.Browser]::New($options)
	$result = $script:XrmBrowser.LoginPage.Login([Uri]::New($url), $userName, $pwd)
    $result = $script:XrmBrowser.GuidedHelp.CloseGuidedHelp()
}

function Initialize-XrmClient {
    Param(
      [Parameter(Mandatory = $true)]
      [String]
      $url,
	  [Parameter(Mandatory = $true)]
      [PSObject]
      $userName,
	  [Parameter(Mandatory = $true)]
      [PSObject]
      $pwd
    )
	
	
    $options = [Microsoft.Dynamics365.UIAutomation.Browser.BrowserOptions]::New()
    $options.BrowserType = [Enum]::Parse([Microsoft.Dynamics365.UIAutomation.Browser.BrowserType], "Chrome")
    $options.PrivateMode = $true
    $options.FireEvents = $false
    $options.Headless = $false
    $options.UserAgent = $false
    $options.DefaultThinkTime = 2000

    $script:WebClient = [Microsoft.Dynamics365.UIAutomation.Api.UCI.WebClient]::New($options)
	$script:XrmApp = [Microsoft.Dynamics365.UIAutomation.Api.UCI.XrmApp]::New($script:WebClient)
	$result = $script:XrmApp.OnlineLogin.Login([Uri]::New($url), $userName, $pwd)
	$result = $script:XrmApp.Dialogs.CloseWarningDialog()
}

function Open-XrmApp{
    Param(
      [Parameter(Mandatory = $true)]
      [String]
      $appName
	)
    $script:XrmApp.Navigation.OpenApp($appName)
}

function Open-XrmAppSubArea{
    Param(
      [Parameter(Mandatory = $true)]
      [String]
      $area,
	  [Parameter(Mandatory = $true)]
      [String]
      $subArea
	)
    $script:XrmApp.Navigation.OpenSubArea($area,$subArea)
}

function Submit-XrmCommand{
    Param(
      [Parameter(Mandatory = $true)]
      [String]
      $command
	)
    $script:XrmApp.CommandBar.ClickCommand($command)
}

function Set-XrmValue{
    Param(
      [Parameter(Mandatory = $true)]
      [String]
      $field,
	  [Parameter(Mandatory = $true)]
      [String]
      $value
	)
	$script:XrmApp.Entity.SetValue($field, $value)
}

function Submit-XrmForm{
    $script:XrmApp.Entity.Save()
}




 

 

     