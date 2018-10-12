<#
Summer 2018
Joshua Williamson
Functions to ease use of Selenium Webdriver to be used inside of Windows PowerShell for Chrome browser automations.

.NOTE
    Dynamic Params: If you have a custom type as a param, the dynamic params will not show up visually with tab completion / Intellisense after that one has been set in the function.
    FIX: Make the custom object a dynamic parm that is always created

#>

#Load needed modules
[System.Reflection.Assembly]::LoadFrom("$PSScriptRoot\WebDriver.dll") | Out-Null
[System.Reflection.Assembly]::LoadFrom("$PSScriptRoot\WebDriver.Support.dll") | Out-Null
[System.Reflection.Assembly]::LoadFrom("$PSScriptRoot\Selenium.WebDriverBackedSelenium.dll") | Out-Null
Import-Module "$PSScriptRoot\GetDynamicParam.ps1"

function New-SeChrome {
    <#
    .SYNOPSIS
    Returns a Selenium Chromedriver object
    .PARAMETER CustomOptions
    Any additional parameters you would like, I included the switches for common ones.
    
    If you wish to use a profile you can use a custom option of "--user-data-dir=$ProfilePath"
    Please note that if you try to open a browser with a profile that is already open then it will freeze Selenium. You must close all other instances of chrome first.
    Also you must leave the "\default" off of the path as ChromeDriver adds it itself. 
    Example : "$env:LOCALAPPDATA\Google\Chrome\User Data" would stat chrome with your default profile.
    WARNING: I found using a profile to be buggy and cause random unintended issues.

    .EXAMPLE
    $driver = New-SeChrome -Maximized
    .EXAMPLE
    $driver = New-SeChrome -Headless -MuteAudio
    #>
    [cmdletbinding()]
    [OutputType([OpenQA.Selenium.Chrome.Chromedriver])]
    param(
        [switch]$Headless,
        [switch]$MuteAudio,
        [switch]$Maximized,
        [switch]$Incognito,
        [string[]]$CustomOptions
    )
    process {
        [OpenQA.Selenium.Chrome.ChromeOptions]$chrome_options = New-Object OpenQA.Selenium.Chrome.ChromeOptions
        if ($Headless) {
            $chrome_options.AddArgument("--headless")
            $chrome_options.AddArgument("--disable-gpu")
        }
        if ($MuteAudio) {
            $chrome_options.AddArgument("--mute-audio")
        }
        if ($Maximized) {
            $chrome_options.AddArgument("--start-maximized")
        }
        if ($incognito) {
            $chrome_options.AddArgument("--incognito")
        }
        if ($NULL -ne $CustomOptions) {
            foreach ($option in $CustomOptions) {
                $chrome_options.AddArgument($option)
            }
        }
        if ($chrome_options.Arguments -eq 0) {
            New-Object -TypeName OpenQA.Selenium.Chrome.ChromeDriver
        }
        else {
            New-Object -TypeName OpenQA.Selenium.Chrome.ChromeDriver($chrome_options)
        }
    }
}

function New-SeFireFox {
    <#
    .SYNOPSIS
    Returns a Selenium FirefoxDriver object
    .PARAMETER CustomOptions
    Any additional parameters you would like, I included the switches for common ones.

    You can start FireFox with a profile / the profile manager by using "-P" with a custom path, or a specific profile with '-P "ProfileName"'. As a warning I have found using a profile can give uninteded errors and issues.
    .EXAMPLE
    $driver = New-SeFireFox -Maximized
    .EXAMPLE
    $driver = New-SeFireFox -Headless -MuteAudio
    #>
    [cmdletbinding()]
    [OutputType()]
    param(
        [switch]$Headless,
        [switch]$MuteAudio,
        [switch]$Maximized,
        [switch]$Private,
        [string[]]$CustomOptions
    )
    process {
        [OpenQA.Selenium.Firefox.FirefoxOptions]$firefox_options = New-Object OpenQA.Selenium.Firefox.FirefoxOptions
        if ($Headless) { $firefox_options.AddArgument("--headless")}
        if ($Maximized) { $firefox_options.AddArgument("--start-maximized")}
        if ($MuteAudio) { $firefox_options.SetPreference("media.volume_scale", "0.0")}
        if ($Private) { $firefox_options.AddArgument("-private")}
        if ($NULL -ne $CustomOptions) {
            foreach ($option in $CustomOptions) {
                $firefox_options.AddArgument($option)
            }
        }
        if ($firefox_options.Arguments -eq 0) {
            New-Object -TypeName OpenQA.Selenium.Firefox.FirefoxDriver
        }
        else {
            New-Object -TypeName OpenQA.Selenium.Firefox.FirefoxDriver($firefox_options)
        }
    }
}

function Get-SeDriverStatus {
    <#
    .SYNOPSIS
    Returns various items from one or more drivers
    .DESCRIPTION
    Takes in one or more WebDrivers and returns a custom object that contains the Url, Title, SessionID, CurrentWindowHandle (browser tab with focus), any other window handles (all other tabs), and then any capabilities of the driver. All of this data can be accessed from the driver object itself though so this may have limited usefullness
    
    .PARAMETER driver
    The Selenium WebDriver object. This can be an array of drivers as well. 
    
    .EXAMPLE
    $cdriver = New-SeChrome -Headless
    $fdriver = New-SeFireFox -Headless
    Get-SeDriverStatus -Driver $cdriver,$fdriver
    
    .NOTES
    General notes
    #>
    [cmdletbinding()]
    param(
        [Parameter(Mandatory = $true, ValueFromPipeLine = $true)] 
        [OpenQA.Selenium.Remote.RemoteWebDriver[]]$DriverList
    )
    process {
        #Need to make this a foreach loop#
        [PSCustomObject[]]$ReturnList = [System.Collections.ArrayList]::new()
        Foreach ($driver in $DriverList) {
            $ReturnList += [PSCustomObject] @{
                DriverType          = $driver.ToString()
                Url                 = $driver.url
                Title               = $driver.title
                SessionId           = $driver.sessionId
                CurrentWindowHandle = $driver.CurrentWindowHandle
                WindowHandles       = $driver.WindowHandles
                Capabilities        = $driver.Capabilities
            }
        }
        $ReturnList
    }
}

function Invoke-SeJavaScript {
    <#
    .SYNOPSIS
    Run JavaScript on browser window that has focus
    
    .DESCRIPTION
    Runs one or more JavaScript scripts that are given as input on the browser. It will attempt to execute the scripts (they can be Async or not depending on if you set the switch value) and if it fails it will print an error message that includes the script that fails (if you have the verbose switch set, otherwise it should fail silently).
    
    .PARAMETER DriverList
    Selenium webdriver(s) that you will execute the JavaScript on. 
    
    .PARAMETER scripts
    One or more scripts to be run, stored as an array of strings. 
    
    .PARAMETER Arguments
    An array of objects to be used as arguments for input scripts. To be clear these arguments are will be the arguments for every script that is input. Your script should use the variable arguments[0] to reference the first argument, arguments[1] for the 2nd and so on. (there is no error if you input more arguments then what is used in the script, but it will error if you try to call or use more arguments then you have).

    .PARAMETER Async
    If set the scripts will be run as AsyncScripts
    
    .EXAMPLE
    New-SeChrome | Invoke-Javascript -Script "while(!document.ready()); window.url = "https://gmail.com; return document.title;"
    
    .EXAMPLE
    #Create new tab
    Invoke-JavaScript -DriverList $d -Script "window.open()"

    .NOTES
    It was complicated to get the arguments to work correctly with powershell. It is suppose to allow you to pass in an array such as $driver.ExecuteScript("return (arguments[0] + arguments[1];", $ArgumentArray) but this kept throwing an error, I believe because it does not recognize the PowerShell format for arrays. It does work if you put each element of the array as parameters. With this I make a string for the command and input all of the arguments into a string (this allows me to put the commas between all of the arguments without knowing how many there are) then I turn the string into a scriptblock and invoke the command.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, ValueFromPipeLine = $true)]
        [OpenQA.Selenium.Remote.RemoteWebDriver[]]$DriverList,
        [Parameter(Mandatory = $true)]
        [string[]]$Scripts,
        [System.Object[]]$Arguments,
        [switch]$Async
    )
    process {
        ForEach ($Driver in $DriverList) {
            if ($Async) {
                ForEach ($script in $scripts) {
                    try {
                        if ($NULL -ne $Arguments) {
                            [string]$command = '$driver.ExecuteAsyncScript($script,'
                            for ($i = 0; $i -lt $arguments.Count; $i++) {
                                $command += "`$Arguments[$i],"
                            }
                            $command = $command.TrimEnd(',') + ')'
                            [scriptblock]::Create($command).Invoke()
                        }
                        else {
                            $driver.ExecuteAsyncScript($script);
                        }
                    }
                    catch {
                        Write-Verbose "ERROR: Invalid JavaScript used:`n$script"
                    }
                }
            }
            else {
                foreach ($script in $scripts) {
                    try {
                        if ($NULL -ne $Arguments) {
                            [string]$command = '$driver.ExecuteScript($script,'
                            for ($i = 0; $i -lt $arguments.Count; $i++) {
                                $command += "`$Arguments[$i],"
                            }
                            $command = $command.TrimEnd(',') + ')'
                            [scriptblock]::Create($command).Invoke()
                        }
                        else {
                            $driver.ExecuteScript($script);
                        }
                    }
                    catch {
                        Write-Verbose "ERROR: Invalid JavaScript used:`n$script"
                    }
                }
            }
        }
    }
}

function Invoke-SeFindElements {
    <#
    .SYNOPSIS
    Search the page for the given element(s)
    
    .DESCRIPTION
    Allows you to search your browser for the element and returns an object representing that element that you can then manipulate. There are various search options you can choose that all require different type of locators and will attempt to run and fail if you do not have the correct format for the search you wish to use. You must input either a WebDriver object or an IWebElement. If nothing is found it will return nothing. 
    
    .PARAMETER driver
    Webdriver to search
    
    .PARAMETER element
    You can also run these search functions with a different element as your base instead of the webdriver.
    
    .PARAMETER by
    Your search type, the values you can choose are predefined and you can tab autocomplete through them. 
    
    .PARAMETER locator
    This will be a string containing how to find the element. Must match the corresponding by type in order to not throw an error. 
    
    .EXAMPLE
    $PasswordElement = Invoke-SeFindElement -Driver $driver -by CssSelector -Locator input[type="password"]
    
    .NOTES
    This should only be used if you know the page has alreadyoaded completely as it will immediately search for the given element. See Invoke-SeWaitUntil to allow for the browser to wait for certain conditions to be filled. 
    #>
    [CmdletBinding()]
    [OutputType([OpenQA.Selenium.IWebElement])]
    param(
        [Parameter(Mandatory = $true, ParameterSetName = "Driver", ValueFromPipeline = $true)]
        [OpenQA.Selenium.Remote.RemoteWebDriver[]]$DriverList,
        [Parameter(Mandatory = $true, ParameterSetName = "Element", ValueFromPipeline = $true)]
        [OpenQA.Selenium.IWebElement[]]$ElementList,
        [Parameter(Mandatory = $true)]
        [ValidateSet("TagName", "ClassName", "ID", "LinkText", "PartialLinkText", "CssSelector", "XPath", "Name")]
        $By,
        [Parameter(Mandatory = $true, HelpMessage = "Input the selector that will match your choosen search method")]
        $Locator
    )
    process {
        if ($PSCmdlet.ParameterSetName -eq "Driver") {
            $itemList = $DriverList
        }
        else {
            $itemList = $ElementList
        }
        foreach ($item in $itemList) {
            $item.FindElements([OpenQA.Selenium.By]::$by($locator))
        }
    }
}

function Invoke-SeWaitUntil {
    <#
    .SYNOPSIS
    Cause the driver to wait for specific conditions to be filled. 
    
    .DESCRIPTION
    Use this function to allow a script to wait until the browser is able to fulfill specific required conditions. There are various predefined conditions that you can use that are defined already, allowing for tab auto completion. Each of these conditions can require different parameters and once you choose a condition the specific params for that condition will become available and must be filled. Sometimes their will be different options you can choose and as soon as you choose one the others will no longer be available (such as searching either with a by condition + locator or inputting an IWebElement object).
    Certain conditions may only return true or false while others will return one or more IWebElement objects (similar to Invoke-SeFindElements).
    
    .PARAMETER Condition
    The condition to wait for. This will define various other dynamic parameters needed once it has been set. Here is a list of dynamic parameters, if the condition is not listed it does not have any(Condition on left, dynamic params on right. All caps is mandatory):
        -AlertState : [bool]STATE
        -ElementExists : [string]BY [string]LOCATOR
        -ElementIsVisible : [string]BY [string]LOCATOR
        -ElementSelectionStateToBe : [bool]SELECTED ([string]BY [string]LOCATOR) -OR ([IWebElement]ELEMENT)
        -ElementToBeClickable : ([string]BY [string]LOCATOR) -OR ([IWebElement]ELEMENT)
        -ElementToBeSelected : ([string]BY [string]LOCATOR) -OR ([IwebElement]ELEMENT [bool]Selected)
        -FrameToBeAvailableAndSwitchToIt : ([string]FRAME_LOCATOR_ID_OR_NAME) -OR ([string]BY [string]LOCATOR)
        -InvisibilityOfElementLocated : [string]BY [string]LOCATOR
        -InvisibilityOfElementWithText : [string]BY [string]LOCATOR [string]TEXT
        -PresenceOfAllElementsLocatedBy : [string]BY [string]LOCATOR
        -StalenessOf : [IWebElement]ELEMENT
        -TextToBePresentInElement : [IWebElement]ELEMENT [string]TEXT
        -TextToBePresentInElementLocated : [string]BY [string]LOCATOR [string]TEXT
        -TextToBePresentInElementValue : [string]TEXT ([string]BY [string]LOCATOR) -OR ([IWebElement]Element)
        -TitleContains : [string]TITLE
        -TitleIs : [string]TITLE
        -UrlContains : [string]UrlFraction
        -UrlMatches : [string]REGEX
        -UrlToBe : [string]URL
        -VisibilityOfAllElementsLocatedBy : ([string]BY [string]LOCATOR) -OR ([System.Collections.ObjectModel.ReadOnlyCollection[OpenQA.Selenium.IWebElement]]ELEMENT_LIST)

    .PARAMETER WaitTime
    The time to wait for the condition to be filled before throwing a TimeOut error. Default is 5 seconds.
    
    .PARAMETER PollingInterval
    Interval at which the condition will be polled to see if it has been fulfilled. Default is 500ms
    
    .PARAMETER TimeOutMessage
    An additional message that can be added to the TimeOut error. 

    .PARAMETER Driver
    Selenium WebDriver Object

    .EXAMPLE
    Invoke-SeWaitUntil -Driver $driver -Condition AlertState -State $true
    
    .NOTES
    There are a lot of dynamic parameters for this function that depends on the condition wanted. I had an issue where the dynamic params were not visually showing up once the driver param was set, it is a bug with PowerShell tab autocomplete and custom objects. To fix that I moved the driver param to be a dynamic param.
    #>
    [cmdletbinding()]
    [OutputType([boolean], [OpenQA.Selenium.IWebElement[]])]
    param(
        [ValidateSet("AlertIsPresent", "AlertState", "ElementExists", "ElementIsVisible", "ElementSelectionStateToBe", "ElementToBeClickable", "ElementToBeSelected", "FrameToBeAvailableAndSwitchToIt", "InvisibilityOfElementLocated", "InvisibilityOfElementWithText", "PresenceOfAllElementsLocatedBy", "StalenessOf", "TextToBePresentInElement", "TextToBePresentInElementLocated", "TextToBePresentInElementValue", "TitleContains", "TitleIs", "UrlContains", "UrlMatches", "UrlToBe", "VisibilityOfAllElementsLocatedBy")]
        $Condition,
        [timespan]$WaitTime = (New-TimeSpan -Seconds 5),
        [timespan]$PollingInterval = [timespan]::FromMilliseconds(500),
        [string]$TimeOutMessage
    )
    DynamicParam {
        $RuntimeParameterDictionary = New-Object System.Management.Automation.RuntimeDefinedParameterDictionary
        $RuntimeParameterDictionary.Add('DriverList', (Get-DynamicParam -Name "DriverList" -type "OpenQA.Selenium.Remote.RemoteWebDriver[]" -mandatory -FromPipeline ))
        
        switch ($Condition) {
            "AlertState" {  
                $RuntimeParameterDictionary.Add('State', (Get-DynamicParam -name "State" -type bool -mandatory))
                break
            }"ElementExists" {
                $RuntimeParameterDictionary.Add("By", (Get-DynamicParam -name "By" -type string -mandatory -ValidateSet "TagName", "ClassName", "ID", "LinkText", "PartialLinkText", "CssSelector", "XPath", "Name" -SetName "By"))
                $RuntimeParameterDictionary.Add("Locator", (Get-DynamicParam -name "Locator" -type string -mandatory -SetName "By"))
                break
            }"ElementIsVisible" {
                $RuntimeParameterDictionary.Add("By", (Get-DynamicParam -name "By" -type string -mandatory -ValidateSet "TagName", "ClassName", "ID", "LinkText", "PartialLinkText", "CssSelector", "XPath", "Name" -SetName "By"))
                $RuntimeParameterDictionary.add("Locator", (Get-DynamicParam -name "Locator" -type string -mandatory -SetName "By"))
                break
            }"ElementSelectionStateToBe" {
                $RuntimeParameterDictionary.Add("By", (Get-DynamicParam -name "By" -type string -mandatory -ValidateSet "TagName", "ClassName", "ID", "LinkText", "PartialLinkText", "CssSelector", "XPath", "Name" -SetName "By"))
                $RuntimeParameterDictionary.Add("Locator", (Get-DynamicParam -name "Locator" -type string -mandatory -SetName "By"))
                $RuntimeParameterDictionary.Add("Selected", (Get-DynamicParam -name "Selected" -type bool -mandatory))
                $RuntimeParameterDictionary.Add("Element", (Get-DynamicParam -name "Element" -type OpenQA.Selenium.IWebElement -mandatory -SetName "ByElement"))
                break
            }"ElementToBeClickable" {
                $RuntimeParameterDictionary.Add("By", (Get-DynamicParam -name "By" -type string -mandatory -ValidateSet "TagName", "ClassName", "ID", "LinkText", "PartialLinkText", "CssSelector", "XPath", "Name" -SetName "By"))
                $RuntimeParameterDictionary.Add("Locator", (Get-DynamicParam -name "Locator" -type string -mandatory -SetName "By"))
                $RuntimeParameterDictionary.Add("Element", (Get-DynamicParam -name "Element" -type OpenQA.Selenium.IWebElement -mandatory -SetName "ByElement"))
                break
            }"ElementToBeSelected" {
                $RuntimeParameterDictionary.Add("By", (Get-DynamicParam -name "By" -type string -mandatory -ValidateSet "TagName", "ClassName", "ID", "LinkText", "PartialLinkText", "CssSelector", "XPath", "Name" -SetName "By"))
                $RuntimeParameterDictionary.Add("Locator", (Get-DynamicParam -name "Locator" -type string -mandatory -SetName "By"))
                $RuntimeParameterDictionary.Add("Element", (Get-DynamicParam -name "Element" -type OpenQA.Selenium.IWebElement -mandatory -SetName "ByElement"))
                $RuntimeParameterDictionary.Add("Selected", (Get-DynamicParam -name "Selected" -type bool -SetName "ByElement"))
                break
            }"FrameToBeAvailableAndSwitchToIt" {
                $RuntimeParameterDictionary.Add("FrameLocatorIdorName", (Get-DynamicParam -name "FrameLocatorIdorName" -type string -mandatory -SetName "StringLocator"))
                $RuntimeParameterDictionary.Add("By", (Get-DynamicParam -name "By" -type string -mandatory -ValidateSet "TagName", "ClassName", "ID", "LinkText", "PartialLinkText", "CssSelector", "XPath", "Name" -SetName "By"))
                $RuntimeParameterDictionary.Add("Locator", (Get-DynamicParam -name "Locator" -type string -mandatory -SetName "By"))
                break
            }"InvisibilityOfElementLocated" {
                $RuntimeParameterDictionary.Add("By", (Get-DynamicParam -name "By" -type string -mandatory -ValidateSet "TagName", "ClassName", "ID", "LinkText", "PartialLinkText", "CssSelector", "XPath", "Name" -SetName "By"))
                $RuntimeParameterDictionary.Add("Locator", (Get-DynamicParam -name "Locator" -type string -mandatory -SetName "By"))
                break
            }"InvisibilityOfElementWithText" {
                $RuntimeParameterDictionary.Add("By", (Get-DynamicParam -name "By" -type string -mandatory -ValidateSet "TagName", "ClassName", "ID", "LinkText", "PartialLinkText", "CssSelector", "XPath", "Name" -SetName "By"))
                $RuntimeParameterDictionary.Add("Locator", (Get-DynamicParam -name "Locator" -type string -mandatory -SetName "By"))
                $RuntimeParameterDictionary.Add("Text", (Get-DynamicParam -name "Text" -type string -mandatory))
                break
            }"PresenceOfAllElementsLocatedBy" {
                $RuntimeParameterDictionary.Add("By", (Get-DynamicParam -name "By" -type string -mandatory -ValidateSet "TagName", "ClassName", "ID", "LinkText", "PartialLinkText", "CssSelector", "XPath", "Name" -SetName "By"))
                $RuntimeParameterDictionary.Add("Locator", (Get-DynamicParam -name "Locator" -type string -mandatory -SetName "By"))
                break
            }"StalenessOf" {
                $RuntimeParameterDictionary.Add("Element", (Get-DynamicParam -name "Element" -type OpenQA.Selenium.IWebElement -mandatory -SetName "ByElement"))
                break
            }"TextToBePresentInElement" {
                $RuntimeParameterDictionary.Add("Element", (Get-DynamicParam -name "Element" -type OpenQA.Selenium.IWebElement -mandatory -SetName "ByElement"))
                $RuntimeParameterDictionary.Add("Text", (Get-DynamicParam -name "Text" -type string -mandatory))
                break
            }"TextToBePresentInElementLocated" {
                $RuntimeParameterDictionary.Add("By", (Get-DynamicParam -name "By" -type string -mandatory -ValidateSet "TagName", "ClassName", "ID", "LinkText", "PartialLinkText", "CssSelector", "XPath", "Name" -SetName "By"))
                $RuntimeParameterDictionary.Add("Locator", (Get-DynamicParam -name "Locator" -type string -mandatory -SetName "By"))
                $RuntimeParameterDictionary.Add("Text", (Get-DynamicParam -name "Text" -type string -mandatory))
                break
            }"TextToBePresentInElementValue" {
                $RuntimeParameterDictionary.Add("By", (Get-DynamicParam -name "By" -type string -mandatory -ValidateSet "TagName", "ClassName", "ID", "LinkText", "PartialLinkText", "CssSelector", "XPath", "Name" -SetName "By"))
                $RuntimeParameterDictionary.Add("Locator", (Get-DynamicParam -name "Locator" -type string -mandatory -SetName "By"))
                $RuntimeParameterDictionary.Add("Element", (Get-DynamicParam -name "Element" -type OpenQA.Selenium.IWebElement -mandatory -SetName "ByElement"))
                $RuntimeParameterDictionary.Add("Text", (Get-DynamicParam -name "Text" -type string -mandatory))
                break
            }"TitleContains" {
                $RuntimeParameterDictionary.Add("Title", (Get-DynamicParam -name "Title" -type string -mandatory))
                break
            }"TitleIs" {
                $RuntimeParameterDictionary.Add("Title", (Get-DynamicParam -name "Title" -type string -mandatory -help_message "Must be exact match of title"))
                break
            }"UrlContains" {
                $RuntimeParameterDictionary.Add("UrlFraction", (Get-DynamicParam -name "UrlFraction" -type string -mandatory))
                break
            }"UrlMatches" {
                $RuntimeParameterDictionary.Add("Regex", (Get-DynamicParam -name "Regex" -type string -mandatory -help_message "Regular expression to match URL"))
                break
            }"UrlToBe" {
                $RuntimeParameterDictionary.Add("Url", (Get-DynamicParam -name "Url" -type string -mandatory))
                break
            }"VisibilityOfAllElementsLocatedBy" {
                $RuntimeParameterDictionary.Add("By", (Get-DynamicParam -name "By" -type string -mandatory -ValidateSet "TagName", "ClassName", "ID", "LinkText", "PartialLinkText", "CssSelector", "XPath", "Name" -SetName "By"))
                $RuntimeParameterDictionary.Add("Locator", (Get-DynamicParam -name "Locator" -type string -mandatory -SetName "By"))
                $RuntimeParameterDictionary.Add("ElementList", (Get-DynamicParam -name "ElementList" -type 'System.Collections.ObjectModel.ReadOnlyCollection[OpenQA.Selenium.IWebElement]' -mandatory -SetName "ByElement"))
                break
            }
            Default {}
        }
        $RuntimeParameterDictionary
    } 
    Begin {
        #This standard block of code loops through bound parameters...
        #It is for the Dynamic Params to make sure they have a variable assigned to them.
        #If no corresponding variable exists, one is created
        #Get common parameters, pick out bound parameters not in that set
        Function _temp { [cmdletbinding()] param() }
        $BoundKeys = $PSBoundParameters.keys | Where-Object { (get-command _temp | Select-Object -ExpandProperty parameters).Keys -notcontains $_}
        foreach ($param in $BoundKeys) {
            if (-not ( Get-Variable -name $param -scope 0 -ErrorAction SilentlyContinue ) ) {
                New-Variable -Name $Param -Value $PSBoundParameters.$param
                Write-Verbose "Adding variable for dynamic parameter '$param' with value '$($PSBoundParameters.$param)'"
            }
        }
        #Appropriate variables should now be defined and accessible
        #Get-Variable -scope 0
    }
    process {
        foreach ($driver in $DriverList) {
            $wait = New-Object -TypeName OpenQA.Selenium.Support.UI.WebDriverWait($driver, $waitTime)
            $wait.PollingInterval = $PollingInterval
            $wait.Message = $TimeOutMessage
            switch ($Condition) {
                "AlertIsPresent" {
                    $wait.Until([OpenQA.Selenium.Support.UI.ExpectedConditions]::AlertIsPresent())
                }
                "AlertState" {  
                    $wait.Until([OpenQA.Selenium.Support.UI.ExpectedConditions]::AlertState($state))
                    break
                }"ElementExists" {
                    $wait.Until([OpenQA.Selenium.Support.UI.ExpectedConditions]::ElementExists([OpenQA.Selenium.By]::$by($locator)))
                    break
                }"ElementIsVisible" {
                    $wait.Until([OpenQA.Selenium.Support.UI.ExpectedConditions]::ElementIsVisible([OpenQA.Selenium.By]::$by($locator)))
                    break
                }"ElementSelectionStateToBe" {
                    if ($PSCmdlet.ParameterSetName -eq "By") {
                        $wait.Until([OpenQA.Selenium.Support.UI.ExpectedConditions]::ElementSelectionStateToBe([OpenQA.Selenium.By]::$by($locator), $selected))
                    }
                    else {
                        $wait.Until([OpenQA.Selenium.Support.UI.ExpectedConditions]::ElementSelectionStateToBe($Element, $selected))
                    }
                    break
                }"ElementToBeClickable" {
                    if ($PSCmdlet.ParameterSetName -eq "By") {
                        $wait.Until([OpenQA.Selenium.Support.UI.ExpectedConditions]::ElementToBeClickable([OpenQA.Selenium.By]::$by($locator)))
                    }
                    else {
                        $wait.Until([OpenQA.Selenium.Support.UI.ExpectedConditions]::ElementToBeClickable($Element))
                    }
                    break
                }"ElementToBeSelected" {
                    if ($PSCmdlet.ParameterSetName -eq "By") {
                        $wait.Until([OpenQA.Selenium.Support.UI.ExpectedConditions]::ElementToBeSelected([OpenAQ.Selenium.By]::$by($locator)))
                    }
                    else {
                        if ($NULL -ne $selected) {
                            $wait.Until([OpenQA.Selenium.Support.UI.ExpectedConditions]::ElementToBeSelected($Element, $selected))
                        }
                        else {
                            $wait.Until([OpenQA.Selenium.Support.UI.ExpectedConditions]::ElementToBeSelected($Element))
                        }
                    }
                    break
                }"FrameToBeAvailableAndSwitchToIt" {
                    if ($PSCmdlet.ParameterSetName -eq "StringLocator") {
                        $wait.Until([OpenQA.Selenium.Support.UI.ExpectedConditions]::FrameToBeAvailableAndSwitchToIt($FrameLocatorIdorName))
                    }
                    else {
                        $wait.Until([OpenQA.Selenium.Support.UI.ExpectedConditions]::FrameToBeAvailableAndSwitchToIt([OpenQA.Selenium.By]::$by($locator)))
                    }
                    break
                }"InvisibilityOfElementLocated" {
                    $wait.Until([OpenQA.Selenium.Support.UI.ExpectedConditions]::InvisibilityOfElementLocated([OpenQA.Selenium.By]::$by($locator)))
                    break
                }"InvisibilityOfElementWithText" {
                    $wait.Until([OpenQA.Selenium.Support.UI.ExpectedConditions]::InvisibilityOfElementWithText([OpenQA.Selenium.By]::$by($locator), $text))
                    break
                }"PresenceOfAllElementsLocatedBy" {
                    $wait.Until([OpenQA.Selenium.Support.UI.ExpectedConditions]::PresenceOfAllElementsLocatedBy([OpenQA.Selenium.By]::$by($locator)))
                    break
                }"StalenessOf" {
                    $wait.Until([OpenQA.Selenium.Support.UI.ExpectedConditions]::PresenceOfAllElementsLocatedBy($Element))
                    break
                }"TextToBePresentInElement" {
                    $wait.Until([OpenQA.Selenium.Support.UI.ExpectedConditions]::TextToBePresentInElement($Element, $text))
                    break
                }"TextToBePresentInElementLocated" {
                    $wait.Until([OpenQA.Selenium.Support.UI.ExpectedConditions]::TextToBePresentInElementLocated([OpenQA.Selenium.By]::$by($locator), $text))
                    break
                }"TextToBePresentInElementValue" {
                    if ($PSCmdlet.ParameterSetName -eq "By") {
                        $wait.Until([OpenQA.Selenium.Support.UI.ExpectedConditions]::TextToBePresentInElementValue([OpenQA.Selenium.By]::$by($locator)))
                    }
                    else {
                        $wait.Until([OpenQA.Selenium.Support.UI.ExpectedConditions]::TextToBePresentInElementValue($Element))
                    }
                    break
                }"TitleContains" {
                    $wait.Until([OpenQA.Selenium.Support.UI.ExpectedConditions]::TitleContains($title))
                    break
                }"TitleIs" {
                    $wait.Until([OpenQA.Selenium.Support.UI.ExpectedConditions]::TitleIs($title))
                    break
                }"UrlContains" {
                    $wait.Until([OpenQA.Selenium.Support.UI.ExpectedConditions]::UrlContains($UrlFraction))
                    break
                }"UrlMatches" {
                    $wait.Until([OpenQA.Selenium.Support.UI.ExpectedConditions]::UrlMatches($Regex))
                    break
                }"UrlToBe" {
                    $wait.Until([OpenQA.Selenium.Support.UI.ExpectedConditions]::UrlToBe($Regex))
                    break
                }"VisibilityOfAllElementsLocatedBy" {
                    if ($PSCmdlet.ParameterSetName -eq "By") {
                        $wait.Until([OpenQA.Selenium.Support.UI.ExpectedConditions]::VisibilityOfAllElementsLocatedBy([OpenQA.Selenium.By]::$by($Locator)))
                    }
                    else {
                        wait.until([OpenQA.Selenium.Support.UI.ExpectedConditions]::VisibilityOfAllElementsLocatedBy($ElementList))
                    }
                    break
                }
                Default {}
            }
        }
    }
}

function Set-SeUrl {
    <#
    .SYNOPSIS
    Navigates driver to given URL
    
    .DESCRIPTION
    Takes an array of webdrivers and sets all of them to the given URL
    
    .PARAMETER DriverList
    Array of WebDriver objects
    
    .PARAMETER Url
    Valid URL
    
    .EXAMPLE
    Set-SeUrl -DriverList $a, $b, $c -Url "https://gmail.com"
    
    .NOTES
    Can take input from pipeline
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [OpenQA.Selenium.Remote.RemoteWebDriver[]]$DriverList,
        [Parameter(Mandatory = $true)]
        [System.Uri]$Url
    )
    process {
        ForEach ($driver in $DriverList) {
            $driver.Url = $url
        }
    }
}

function Exit-SeDriver {
    <#
    .SYNOPSIS
    Closes the driver completely
    
    .DESCRIPTION
    Runs the .quit() dispose method on all browsers input into the function
    
    .PARAMETER DriverList
    Array of WebDriver objects to dispose
    
    .EXAMPLE
    Close-SeDriver -DriverList $a, $b, $c
    
    .NOTES
    Using $driver.Close() will only close the current tab
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [OpenQA.Selenium.Remote.RemoteWebDriver[]]$DriverList
    )
    end {
        ForEach ($driver in $DriverList) {
            $driver.quit()
        }
    }
}

function Send-SeKeys {
    <#
    .SYNOPSIS
    Sends keys to the driver or element
    
    .DESCRIPTION
    Takes in either a WebDriver object or an IWebElement object and sends keys to it. It can send special keys as well, and if you wish to send both special keys + other text as well there is a value for SpecialKeys called PositionNthTextHere that will allow you to choose where the text goes. This number of time you use this value must be equal to the number of text strings you have or it will throw an error.
    
    .PARAMETER DriverList
    Array of WebDriver objects, mandatory only if no Elements are input
    
    .PARAMETER ElementList
    Arrray of WebElements, mandatory if no DriverList is specified
    
    .PARAMETER Text
    Text to send to object, mandatory. Can be an empty string though I believe.
    
    .PARAMETER SpecialKeys
    List of special keys to send. This can be any size and has tab autocompletion for the available keys.
    
    .EXAMPLE
    #Will send Username then press enter to all inputs of type username in input driver
    Send-SeKeys -ElementList (Invoke-SeFindElment -Driver $driver -By CssSelector -Locator input[type*="username"]) -text "MyUsername" -SpecialKeys PositionNthTextHere, Enter
    
    .NOTES
    Will either take an array of elements from one driver or an array of drivers that will have the keys send to them.
    #>
    [cmdletbinding()]
    param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, ParameterSetName = "Driver")]
        [OpenQA.Selenium.Remote.RemoteWebDriver[]]$DriverList,
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, ParameterSetName = "Element")]
        [OpenQA.Selenium.IWebElement[]]$ElementList,
        [Parameter(Mandatory = $true)]
        [string[]]$Text,
        [ValidateSet("Add", "Alt", "ArrowDown", "ArrowLeft", "ArrowRight", "ArrowUp", "Backspace", "Cancel", "Clear", "Command", "Control", "Decimal", "Delete", "Divide", "Down", "End", "Enter", "Equal", "Escape", "F1", "F2", "F3", "F4", "F5", "F6", "F7", "F8", "F9", "F10", "F11", "F12", "Help", "Home", "Insert", "Left", "LeftAlt", "LeftControl", "LeftShift", "Meta", "Multiply", "Null", "NumberPAd0", "NumberPad1", "NumberPad2", "NumberPad3", "NumberPad4", "NumberPad5", "NumberPad6", "NumberPad7", "NumberPad8", "NumberPad9", "PageDown", "PageUp", "Pause", "Return", "Right", "Semicolon", "Separator", "Shift", "Space", "Subtract", "Tab", "Up", "PositionNthTextHere")]
        [string[]]$SpecialKeys
    )
    begin {
        if (($SpecialKeys -eq "PositionNthTextHere").Count -ne $text.Count) {
            throw "When using Special Keys your PositionNthTextHere must equal the number of text strings. This is to position your text in the correct order in relation to the special keys and goes in order from first string entered to last (the SpecialKeys array is treated the same)."
        }
        else {
            $true
        }
    }
    process {
        if ($PSCmdlet.ParameterSetName -eq "Driver") {
            $itemList = $DriverList
        }
        else {
            $itemList = $ElementList
        }
        [string]$SendKeys = $NULL
        if ($NULL -ne $SpecialKeys) {
            $TextStringsUsed = 0;
            foreach ($key in $SpecialKeys) {
                if ($key -eq "PositionNthTextHere") {
                    $key = $key -replace $key, $text[$TextStringsUsed]
                    $TextStringsUsed++
                    $SendKeys += $key
                }
                else {
                    $SendKeys += [OpenQA.Selenium.Keys]::$key
                }
            }
        }
        else {
            $SendKeys = $text -join ""
        }
        foreach ($item in $itemList) {
            if ($PSCmdlet.ParameterSetName -eq "Driver") {
                $item.Keyboard.SendKeys($SendKeys)
            }
            else {
                $item.SendKeys($SendKeys)
            }
        }
        $itemList
    }
}

function Set-SeTabFocus {
    <#
    .SYNOPSIS
    Switches driver tab to new tab
    
    .DESCRIPTION
    Changes tab to the given input, will throw an error if the tab does not exist
    
    .PARAMETER DriverList
    Array of drivers to switch tabs
    
    .PARAMETER TabNumber
    Tab to switch to
    
    .PARAMETER UrlOrTitle
    A regex that will match with a URL or Title of a tab. Will stop at the first tab that matches the expression or go all the way to the last tab and stop their.

    .EXAMPLE
    Set-SeTabFoucs -DriverList $a, $b -TabNumber 2 | Invoke-SeJavaScript -Script "return document.title;"
    
    .NOTES
   This does return the Driver object so you can pass the output of this function into another one that accepts an array of WebDrivers. Using the Regex to match will be much slower then switching to a numbered tab.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [OpenQA.Selenium.Remote.RemoteWebDriver[]]$DriverList,
        [Parameter(Mandatory = $true, ParameterSetName = "Number")]
        [int]$TabNumber,
        [Parameter(Mandatory = $true, ParameterSetName = "Regex")]
        [regex]$UrlOrTitle
    )
    process {
        Foreach ($driver in $DriverList) {
            if ($PSCmdlet.ParameterSetName -eq "Number") {
                if ($TabNumber -gt $driver.WindowHandles.Count) {
                    throw "Tab number can't be greater then the number of tabs"
                }
                $driver.SwitchTo().Window($driver.WindowHandles[$TabNumber])
            }
            else {
                $WindowHandles = $driver.WindowHandles
                For ($i = 0; $i -lt $WindowHandles.Count; $i++) {
                    $NewTab = $driver.SwitchTo().Window($WindowHandles[$i])
                    if ($NewTab.Url -match $UrlOrTitle -or $NewTab.Title -match $UrlOrTitle) {
                        $driver.SwitchTo().Window($WindowHandles[$i])
                        break
                    }
                }
            }
        }
    }
}

function Close-SeTab {
    <#
    .SYNOPSIS
    Closes tab based on given parameter
    
    .DESCRIPTION
    Will close the tab that you specify through the CurrentWindow switch, TabNumber or Regex match with the URL or Title. As a warning if the current window is closed the driver will be left focusing nothing and will need to have the window focus set to a window that exists.
    
    .PARAMETER DriverList
    Array of WebDriver objects
    
    .PARAMETER TabNumber
    Tab number to be closed (From left to right, 0-N)
    
    .PARAMETER UrlOrTitle
    Regex to match the URL or title, will close first one it runs into.
    
    .PARAMETER CurrentTab
    Closes the tab currently focused
    
    .EXAMPLE
    #Close the 2nd tab and set focus to the first for all drivers in list.
    Close-SeTab -DriverList $a, $b, $c -TabNumber 1 | Set-SeTab -TabNumber 0
    
    .NOTES
    The regex will be slow as it switches to every tab until it finds a match.
    #>
    [CmdletBinding(DefaultParameterSetName = "Number")]
    param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [OpenQA.Selenium.Remote.RemoteWebDriver[]]$DriverList,
        [Parameter(Mandatory = $true, ParameterSetName = "Number")]
        [int]$TabNumber,
        [Parameter(Mandatory = $true, ParameterSetName = "Regex")]
        [regex]$UrlOrTitle,
        [Parameter(Mandatory = $true, ParameterSetName = "Current")]
        [switch]$CurrentTab
    )
    process {
        ForEach ($driver in $DriverList) {
            if ($CurrentTab) {
                $driver.Close()
                Write-Verbose "Current window was closed, driver is now not focused on any tab and must be set again."
            }
            elseif ($PSCmdlet.ParameterSetName -eq "Number") {
                $CurrentWindow = $Driver.CurrentWindowHandle
                ($Driver.SwitchTo().Window($Driver.WindowHandles[$TabNumber])).Close()
                try {
                    $Driver.SwitchTo().Window($CurrentWindow)
                }
                catch {
                    Write-Verbose "Current window was closed, driver is now not focused on any tab and must be set again."
                }
            }
            elseif ($PSCmdlet.ParameterSetName -eq "Regex") {
                $CurrentWindow = $Driver.CurrentWindowHandle
                $WindowHandles = $driver.WindowHandles
                For ($i = 0; $i -lt $WindowHandles.Count; $i++) {
                    $NewTab = $driver.SwitchTo().Window($WindowHandles[$i])
                    if ($NewTab.Url -match $UrlOrTitle -or $NewTab.Title -match $UrlOrTitle) {
                        $NewTab.Close()
                        break
                    }
                }
                try {
                    $Driver.SwitchTo().Window($CurrentWindow)
                }
                catch {
                    Write-Verbose "Current window was close, driver is now not focused on any tab and must be set again."
                }
            }
        }
    }
}
function Get-SeScreenShot {
    <#
    .SYNOPSIS
    Screen shot the borwser page.
    
    .DESCRIPTION
    Screen shots the browser page being shown (does not screen shot the whole page, so it may be important to naviagte to the part first with eitehr Invoke-SeFind or Invoke-SeWaitUntil). Returns the screenshot objects, if any of the Destination or save options are chosen it will save the screenshot in the given format, else it will just return the screenshot. Can take an array of WebDrivers.
    
    .PARAMETER DriverList
   Array of WebDriver objects
    
    .PARAMETER Format
    Image format to be saved in
    
    .PARAMETER DestinationDirectory
    Destination for the image files, will create the directory if it doesn't exist
    
    .PARAMETER FileBaseName
    The file name for the images, will add a 0,1, 2.... based on the order of images
    
    .PARAMETER FullPage
    This will let you take a screen shot of the entire page instead of just what is visible. This will only work with Chrome as the ExecuteChromeCommand method is needed.
    
    .EXAMPLE
    Get-SeScreenshot -DriverList $a, $b -Format Png -DestinationDirectory "$home\desktop\temp" -FileBaseName browser_img
    #Will save two images browser_img_0.png and browser_img_1.png in the temp directory.
    
    .NOTES
    General notes
    #>
    [CmdletBinding(DefaultParameterSetName = "DontSave")]
    param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [OpenQA.Selenium.Remote.RemoteWebDriver[]]$DriverList,
        [Parameter(Mandatory = $true, ParameterSetName = "SaveAs")]
        [ValidateSet("Png", "Jpeg", "Gif", "Tiff", "Bmp")]
        $Format,
        [Parameter(Mandatory = $true, ParameterSetName = "SaveAs")]
        $DestinationDirectory,
        [Parameter(Mandatory = $true, ParameterSetName = "SaveAs")]
        [ValidateScript( {
                if ($_ -match ".*?\.[A-Za-z]+") {
                    throw "Do not include a file extension as it will be added from the Format parameter"
                }
                else {
                    $true
                }})]
        $FileBaseName,
        [switch]$FullPage
    )
    process {
        $image_number = 0
        foreach ($Driver in $DriverList) {
            if ($FullPage) {
                $metrics = [System.Collections.Generic.Dictionary[string, System.Object]]::new()
                $metrics["width"] = Invoke-SeJavaScript -DriverList $driver -Scripts "return Math.max(window.innerWidth,document.body.scrollWidth,document.documentElement.scrollWidth)"
                $metrics["height"] = Invoke-SeJavaScript -DriverList $driver -Scripts "return Math.max(window.innerHeight,document.body.scrollHeight,document.documentElement.scrollHeight)"
                $metrics["deviceScaleFactor"] = [double](Invoke-SeJavaScript -DriverList $driver -Scripts "window.devicePixelRatio")
                $metrics["mobile"] = Invoke-SeJavaScript -DriverList $driver -Scripts "return typeof window.orientation !== 'undefined'"
                Invoke-SeChromeCommand -DriverList $driver -commandName "Emulation.setDeviceMetricsOverride" -commandParameters $metrics
            }
            $ScreenShot = $driver.GetScreenShot()
            if ($PSCmdlet.ParameterSetName -eq "SaveAs") {
                if (-not (Test-Path $DestinationDirectory)) {
                    New-Item -ItemType Directory -Path $DestinationDirectory    
                }
                $ScreenShot.SaveAsFile("$DestinationDirectory\$($FileBaseName)_$image_number.$Format", [OpenQA.Selenium.ScreenshotImageFormat]::$Format)
            }
            $image_number++
            $ScreenShot
            Invoke-SeChromeCommand -DriverList $driver -commandName "Emulation.clearDeviceMetricsOverride"
        }
    }
}



function Get-SeElementScreenShot {
    <#
    .SYNOPSIS
    Tries to screen shot specific elements
    
    .DESCRIPTION
    Dependant on the Get-SeScreenshot function. The list of drivers should all be having the same pages and elements for each call of this function. It works by screenshotting the page for each driver then going through each of those screen shots and cropping out every Element in ElementList. Returns each of the images.
    
    This now does work with the full page for any element on the page (only works with chrome for the full page).

    .PARAMETER DriverList
    Arrat if WebDrivers
    
    .PARAMETER ElementList
    Array of elements to crop the picture to
    
    .PARAMETER Format
    Format to save image
    
    .PARAMETER DestinationDirectory
    Directory to save all of the images, will make the directort if it doesnt exist
    
    .PARAMETER FileBaseName
    Base file name, will have a number appened to the end of it for each file
    
    .PARAMETER SaveBaseImage
    Switch to designate if you would like to save the full page screen shot beside the images of the elements.

    .EXAMPLE
    $ElementList = Invoke-SeFineElements -DriverList $driver -By CssSelector -Locator table
    Get-SeElementScreenshot -DriverList $driver -ElementList $ElementList $Format Png $DestinationDirectory "$home\desktop\temp" -FileBaseName CropImage
    
    .NOTES
    This dows not seem to work well on certain computers. On my desktop it takes screenshots and crops as expected, but on my Microsoft Surface all of the images are cropped incorrectly.
    #>
    [CmdletBinding(DefaultParameterSetName = "DontSave")]
    [OutputType([System.Drawing.Bitmap])]
    param(
        [Parameter(Mandatory = $true)]
        [OpenQA.Selenium.Remote.RemoteWebDriver[]]$DriverList,
        [Parameter(Mandatory = $true)]
        [OpenQA.Selenium.IWebElement[]]$ElementList,
        [Parameter(Mandatory = $true, ParameterSetName = "SaveAs")]
        [ValidateSet("Png", "Jpeg", "Gif", "Tiff", "Bmp")]
        $Format,
        [Parameter(Mandatory = $true, ParameterSetName = "SaveAs")]
        [string]$DestinationDirectory,
        [Parameter(Mandatory = $true, ParameterSetName = "SaveAs")]
        [ValidateScript( {
                if ($_ -match ".*?\.[A-Za-z]+") {
                    throw "Do not include a file extension as it will be added from the Format parameter"
                }
                else {
                    $true
                }})]
        [string]$FileBaseName,
        [switch]$SaveBaseImage
    )
    end {
        $screenShots = Get-SeScreenShot -DriverList $DriverList -FullPage
        $driver_num = 0
        foreach ($ScreenShot in $ScreenShots) {
            if ($SaveBaseImage) {
                $ScreenShot.SaveAsFile("$($DestinationDirectory)\$($FileBaseName)_$($driver_num).$Format", [System.Drawing.Imaging.ImageFormat]::$Format)
            }
            foreach ($Element in $ElementList) {
                [System.Drawing.Bitmap] $image = New-Object System.Drawing.Bitmap((New-Object System.IO.MemoryStream ($ScreenShot.AsByteArray, $ScreenShot.Count)))

                [System.Drawing.Rectangle] $crop = New-Object System.Drawing.Rectangle([System.Math]::Abs($element.location.X), [System.Math]::Abs($element.location.Y), [System.Math]::Abs($element.size.Width), [System.Math]::Abs($element.size.Height))

                $image = $image.clone($crop, $image.PixelFormat)
                if ($PSCmdlet.ParameterSetName -eq "SaveAs") {
                    if (-not (Test-Path $DestinationDirectory)) {
                        New-Item -ItemType Directory -Path $DestinationDirectory    
                    }
                    $image.Save("$($DestinationDirectory)\$($FileBaseName)_$($driver_num)_$($image_num).$Format", [System.Drawing.Imaging.ImageFormat]::$Format)                    
                }
                $image_num++
                $image
            }
            $driver_num++
        }
    }
}

function Invoke-SeNavigate {
    <#
    .SYNOPSIS
    Invokes the driver navigate command
    
    .DESCRIPTION
    Uses the browser navigate class to move the browser forward, back, refresh the page or go to new url. Predefined navigate options can be tab autocompleted.
    
    .PARAMETER DriverList
    Array of WebDrivers
    
    .PARAMETER Navigate
    Navigate option :Back, Forward or Refresh. Each is the same as doing it manually in a browser window.
    
    .PARAMETER Url
    If chosen will navigate to that URL.
    
    .EXAMPLE
    Invoke-SeNavigate -Navigate Back -DriverList $a, $b, $c
    
    .NOTES
    General notes
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [OpenQA.Selenium.Remote.RemoteWebDriver[]]$DriverList,
        [Parameter(Mandatory = $true, ParameterSetName = "Navigate")]
        [ValidateSet("Back", "Forward", "Refresh")]
        $Navigate,
        [Parameter(Mandatory = $true, ParameterSetName = "Url")]
        [System.Uri]$Url
    )
    process {
        Foreach ($Driver in $DriverList) {
            if ($PSCmdlet.ParameterSetName -eq "Url") {
                $Driver.Navigate().GoToUrl($Url)
            }
            else {
                $Driver.Navigate().$Navigate()
            }
        }
    }
}

function Invoke-SeSwitchTo {
    <#
    .SYNOPSIS
    Invokes the drivers SwitchTo class options
    
    .DESCRIPTION
    Allows you to switch to various parts of the browser. There are options to switch to an ActiveElement, any Alert that is present, default content of page, a specified frame or parentframe or a window (tab).
    
    .PARAMETER DriverList
    This is a parameter to accept an array of WebDrivers, may not show up in Get-Help as it needs to be dynamic to get around a bug with dynamic params.

    .PARAMETER SwitchTo
    Option to witch to. Dynamic params will become available based on the option that is chosen.
    
    .EXAMPLE
    Invoke-SeSwitchTo
    
    .NOTES
    General notes
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [ValidateSet("ActiveElement", "Alert", "DefaultContent", "Frame", "ParentFrame", "Window")]
        [string]$SwitchTo
    )
    DynamicParam {
        $RuntimeParameterDictionary = New-Object System.Management.Automation.RuntimeDefinedParameterDictionary

        #In order not to mess up other dynamic params all custom objects also need to be dynamic
        $RuntimeParameterDictionary.Add('DriverList', (Get-DynamicParam -Name "DriverList" -type "OpenQA.Selenium.Remote.RemoteWebDriver[]" -mandatory -FromPipeline))
        switch ($SwitchTo) {
            "Frame" {
                $RuntimeParameterDictionary.Add('FrameIndex', (Get-DynamicParam -name 'FrameIndex' -type int -mandatory -SetName "Index"))
                $RuntimeParameterDictionary.Add('FrameName', (Get-DynamicParam -name 'FrameName' -type string -mandatory -SetName "Name"))
                $RuntimeParameterDictionary.Add('FrameElement', (Get-DynamicParam -name "FrameElement" -type OpenQA.Selenium.IWebElement -mandatory -SetName "Element"))
            }"Window" {
                $RuntimeParameterDictionary.Add("windowHandleOrName", (Get-DynamicParam -name "windowHandleOrName" -type string -mandatory))
            }
        }
        $RuntimeParameterDictionary
    }
    begin {
        #This standard block of code loops through bound parameters...
        #It is for the Dynamic Params to make sure they have a variable assigned to them.
        #If no corresponding variable exists, one is created
        #Get common parameters, pick out bound parameters not in that set
        Function _temp { [cmdletbinding()] param() }
        $BoundKeys = $PSBoundParameters.keys | Where-Object { (get-command _temp | Select-Object -ExpandProperty parameters).Keys -notcontains $_}
        foreach ($param in $BoundKeys) {
            if (-not ( Get-Variable -name $param -scope 0 -ErrorAction SilentlyContinue ) ) {
                New-Variable -Name $Param -Value $PSBoundParameters.$param
                Write-Verbose "Adding variable for dynamic parameter '$param' with value '$($PSBoundParameters.$param)'"
            }
        }
        #Appropriate variables should now be defined and accessible
        #Get-Variable -scope 0
    }
    process {
        foreach ($driver in $driverList) {
            switch ($SwitchTo) {
                "Frame" {
                    switch ($FrameIndex, $FrameName, $FrameElement) {
                        $NULL {}
                        default {
                            $Locator = $_
                            break
                        }
                    }
                    $driver.SwitchTo().Frame($Locator)
                    break
                }"Window" {
                    $driver.SwitchTo().Window($windowHandleOrName)
                    break
                }default {
                    $driver.SwitchTo().$SwitchTo()
                }
            }
        }
    }
}

function Invoke-SeKeyboard {
    <#
    .SYNOPSIS
    Invokes a key event, not very reliable.
    
    .DESCRIPTION
    The main purpose of this command is to Press/Release keys as it is only for a WebDriver object. You should use the Send-SeSendKeys for most cases in my opinion.
    
    .PARAMETER KeyEvent
    The key event to trigger
    
    .EXAMPLE
    Turns out this isn't even working, so just avoid this command.

    Invoke-SeKeyboard -KeyEvent PressKey -DriverList $d, $c -keyToPress Control
    Invoke-SeKeyboard -KeyEvent SendKeys -DriverList $d, $c -keySequence "T"
    Invoke-SeKeyboard -Keyevent ReleaseKey -DriverList $d, $c -keyToRelease Control
    
    .NOTES
    Most of the use cases for this could be done with the Send-SeKeys class.
    Turns out this isn't even working, so just avoid this command.

    If you want to open a new tab is Invoke-SeJavascript with a script of "window.open()"
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet("PressKey", "ReleaseKey", "SendKeys")]
        [string]$KeyEvent
    )
    DynamicParam {
        $RuntimeParameterDictionary = New-Object System.Management.Automation.RuntimeDefinedParameterDictionary
        $RuntimeParameterDictionary.Add('DriverList', (Get-DynamicParam -Name "DriverList" -type "OpenQA.Selenium.Remote.RemoteWebDriver[]" -mandatory -FromPipeline -SetName "Driver"))

        switch ($KeyEvent) {
            "PressKey" {
                $RuntimeParameterDictionary.Add("keyToPress", (Get-DynamicParam -name "keyToPress" -Type string -mandatory -ValidateSet ("Add", "Alt", "ArrowDown", "ArrowLeft", "ArrowRight", "ArrowUp", "Backspace", "Cancel", "Clear", "Command", "Control", "Decimal", "Delete", "Divide", "Down", "End", "Enter", "Equal", "Escape", "F1", "F2", "F3", "F4", "F5", "F6", "F7", "F8", "F9", "F10", "F11", "F12", "Help", "Home", "Insert", "Left", "LeftAlt", "LeftControl", "LeftShift", "Meta", "Multiply", "Null", "NumberPAd0", "NumberPad1", "NumberPad2", "NumberPad3", "NumberPad4", "NumberPad5", "NumberPad6", "NumberPad7", "NumberPad8", "NumberPad9", "PageDown", "PageUp", "Pause", "Return", "Right", "Semicolon", "Separator", "Shift", "Space", "Subtract", "Tab", "Up")))
            }"ReleaseKey" {
                $RuntimeParameterDictionary.Add("keyToRelease", (Get-DynamicParam -name "keyToRelease" -Type string -mandatory -ValidateSet ("Add", "Alt", "ArrowDown", "ArrowLeft", "ArrowRight", "ArrowUp", "Backspace", "Cancel", "Clear", "Command", "Control", "Decimal", "Delete", "Divide", "Down", "End", "Enter", "Equal", "Escape", "F1", "F2", "F3", "F4", "F5", "F6", "F7", "F8", "F9", "F10", "F11", "F12", "Help", "Home", "Insert", "Left", "LeftAlt", "LeftControl", "LeftShift", "Meta", "Multiply", "Null", "NumberPAd0", "NumberPad1", "NumberPad2", "NumberPad3", "NumberPad4", "NumberPad5", "NumberPad6", "NumberPad7", "NumberPad8", "NumberPad9", "PageDown", "PageUp", "Pause", "Return", "Right", "Semicolon", "Separator", "Shift", "Space", "Subtract", "Tab", "Up")))
            }"SendKeys" {
                $RuntimeParameterDictionary.Add("keySequence", (Get-DynamicParam -name "keySequence" -type string -mandatory ))
            }
        }
        $RuntimeParameterDictionary
    }
    begin {
        #This standard block of code loops through bound parameters...
        #It is for the Dynamic Params to make sure they have a variable assigned to them.
        #If no corresponding variable exists, one is created
        #Get common parameters, pick out bound parameters not in that set
        Function _temp { [cmdletbinding()] param() }
        $BoundKeys = $PSBoundParameters.keys | Where-Object { (get-command _temp | Select-Object -ExpandProperty parameters).Keys -notcontains $_}
        foreach ($param in $BoundKeys) {
            if (-not ( Get-Variable -name $param -scope 0 -ErrorAction SilentlyContinue ) ) {
                New-Variable -Name $Param -Value $PSBoundParameters.$param
                Write-Verbose "Adding variable for dynamic parameter '$param' with value '$($PSBoundParameters.$param)'"
            }
        }
    }
    process {
        foreach ($driver in $DriverList) {
            switch ($KeyEvent) {
                "PressKey" {
                    $driver.Keyboard.PressKey([OpenQA.Selenium.Keys]::$keyToPress)
                }"ReleaseKey" {
                    $driver.Keyboard.ReleaseKey([OpenQA.Selenium.Keys]::$keyToRelease)
                }"SendKeys" {
                    $driver.Keyboard.SendKeys($keySequence)
                }
            }
        }
    }
}
function Invoke-SeMouseClick() {
    <#
    .SYNOPSIS
    Invoke commands normally done with a mouse

    .DESCRIPTION
    Used to click / submit elements as well as to do various functions with the divers mouse class. I don't know how to use the ICoordinates class that is needed for manipulating the driver mouse, but this function's main use is to click elements.
    
    .PARAMETER DriverList
    Parameter description
    
    .PARAMETER ElementList
    Parameter description
    
    .EXAMPLE
    An example
    
    .NOTES
    General notes
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, ParameterSetName = "Driver")]
        [OpenQA.Selenium.Remote.RemoteWebDriver[]]$DriverList,
        [Parameter(Mandatory = $true, ParameterSetName = "Driver")]
        [ValidateSet("Click", "ContextClick", "DoubleClick", "MouseDown", "MouseMove", "MouseUp")]
        [string]$MouseEvent,
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, ParameterSetName = "Element")]
        [OpenQA.Selenium.IWebElement[]]$ElementList
    )
    DynamicParam {
        $RuntimeParameterDictionary = New-Object System.Management.Automation.RuntimeDefinedParameterDictionary
        
        switch ($MouseEvent) {
            "MouseMove" {
                $RuntimeParameterDictionary.Add("offsetX", (Get-DynamicParam -name "offsetX" -Type int))
                $RuntimeParameterDictionary.Add("offsetY", (Get-DynamicParam -name "offsetY" -type int))
            }
        }
        $RuntimeParameterDictionary
    }
    begin {
        #This standard block of code loops through bound parameters...
        #It is for the Dynamic Params to make sure they have a variable assigned to them.
        #If no corresponding variable exists, one is created
        #Get common parameters, pick out bound parameters not in that set
        Function _temp { [cmdletbinding()] param() }
        $BoundKeys = $PSBoundParameters.keys | Where-Object { (get-command _temp | Select-Object -ExpandProperty parameters).Keys -notcontains $_}
        foreach ($param in $BoundKeys) {
            if (-not ( Get-Variable -name $param -scope 0 -ErrorAction SilentlyContinue ) ) {
                New-Variable -Name $Param -Value $PSBoundParameters.$param
                Write-Verbose "Adding variable for dynamic parameter '$param' with value '$($PSBoundParameters.$param)'"
            }
        }
        #Appropriate variables should now be defined and accessible
        #Get-Variable -scope 0

        #Validate both offset's are either set or not set
        switch -regex ($offsetX) {
            "." {
                switch -regex ($offsetY) {
                    "." {
                        $true
                    }$NULL {
                        throw "offsetX and Y must both be set or both be empty"
                    }
                }
            }$NULL {
                switch -regex ($offsetX) {
                    $NULL {
                        $true
                    }"." {
                        throw "offsetX and Y must both be set or both be empty"
                    }
                }
            }
        }
    }
    process {
        if ($PSCmdlet.ParameterSetName -eq "Driver") {
            foreach ($driver in $DriverList) {
                if ($NULL -ne $offsetX) {
                    $DriverList.Mouse.$MouseEvent($Coordinates, $offsetX, $offsetY)
                }
                else {
                    $DriverList.Mouse.$MouseEvent($Coordinates)
                }
            }
        }
        else {
            foreach ($Element in $ElementList) {
                $ElementList.Click()
            }
        }
    }
}

function Invoke-SeChromeCommand {
    <#
    .SYNOPSIS
    Executes custom Chrome command
    
    .PARAMETER DriverList
    Must be a chrome WebDriver
    
    .PARAMETER commandName
    Name of command to execute
    
    .PARAMETER commandParameters
    Parameters of command to execute. It will default to an empty array if no params are needed
    
    .EXAMPLE
    $metrics = [System.Collections.Generic.Dictionary[string, System.Object]]::new()
    $metrics["width"] = Invoke-SeJavaScript -DriverList $driver -Scripts "return Math.max(window.innerWidth,document.body.scrollWidth,document.documentElement.scrollWidth)"
    $metrics["height"] = Invoke-SeJavaScript -DriverList $driver -Scripts "return Math.max(window.innerHeight,document.body.scrollHeight,document.documentElement.scrollHeight)"
    $metrics["deviceScaleFactor"] = Invoke-SeJavaScript -DriverList $driver -Scripts "window.devicePixelRatio"
    $metrics["mobile"] = Invoke-SeJavaScript -DriverList $driver -Scripts "return typeof window.orientation !== 'undefined'"
    $driver.ExecuteChromeCommand("Emulation.setDeviceMetricsOverride", $metrics)
    
    .NOTES
    Here is a site of possible commands (not all work very well), it is all the same style as the example shown: https://chromedevtools.github.io/devtools-protocol/
    #>
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [OpenQA.Selenium.Chrome.ChromeDriver[]]$DriverList,
        [Parameter(Mandatory = $true)]
        [string]$commandName,
        [System.Collections.Generic.Dictionary[string, System.Object]]$commandParameters = ([System.Collections.Generic.Dictionary[string, System.Object]]::new())
    )
    Process {
        Foreach ($Driver in $DriverList) {
            $Driver.ExecuteChromeCommand($commandName, $commandParameters)
        }
    }
}

function Invoke-SeManageCookies {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)]
        [ValidateSet("AddCookie", "DeleteAllCookies", "DeleteCookie", "DeleteCookieNamed", "GetCookieNamed", "GetAllCookies")]
        [string]$Option
    )
    DynamicParam {
        $RuntimeParameterDictionary = New-Object System.Management.Automation.RuntimeDefinedParameterDictionary

        #In order not to mess up other dynamic params all custom objects also need to be dynamic
        $RuntimeParameterDictionary.Add('DriverList', (Get-DynamicParam -Name "DriverList" -type "OpenQA.Selenium.Remote.RemoteWebDriver[]" -mandatory -FromPipeline))

        switch ($option)
        {
            "AddCookie"{
                $RuntimeParameterDictionary.Add("Cookie", (Get-DynamicParam -name "Cookie" -Type OpenQA.Selenium.Cookie -mandatory))
            }"DeleteCookie"{
                $RuntimeParameterDictionary.Add("Cookie", (Get-DynamicParam -name "Cookie" -Type OpenQA.Selenium.Cookie -mandatory))
            }"DeleteCookieName"{
                $RuntimeParameterDictionary.Add("Name", (Get-DynamicParam -name "Name" -type string -mandatory))
            }"GetCookieNamed"{
                $RuntimeParameterDictionary.Add("Name", (Get-DynamicParam -name "Name" -type string -mandatory))
            }
        }
        $RuntimeParameterDictionary
    }
    begin {
        #This standard block of code loops through bound parameters...
        #It is for the Dynamic Params to make sure they have a variable assigned to them.
        #If no corresponding variable exists, one is created
        #Get common parameters, pick out bound parameters not in that set
        Function _temp { [cmdletbinding()] param() }
        $BoundKeys = $PSBoundParameters.keys | Where-Object { (get-command _temp | Select-Object -ExpandProperty parameters).Keys -notcontains $_}
        foreach ($param in $BoundKeys) {
            if (-not ( Get-Variable -name $param -scope 0 -ErrorAction SilentlyContinue ) ) {
                New-Variable -Name $Param -Value $PSBoundParameters.$param
                Write-Verbose "Adding variable for dynamic parameter '$param' with value '$($PSBoundParameters.$param)'"
            }
        }
        #Appropriate variables should now be defined and accessible
        #Get-Variable -scope 0
    }
    process{
        switch ($Option)
        {
            "AddCookie"{
                $Driver.Manage().Cookies.AddCookie($cookie)
            }"DeleteAllCookies"{
                $Driver.Manage().Cookies.DeleteAllCookies()
            }"DeleteCookie"{
                $Driver.Manage().Cookies.DeleteCookie($cookie)
            }"DeleteCookieNamed"{
                $Driver.Manage().Cookies.DeleteCookieNamed($name)
            }"GetCookieNamed"{
                $Driver.Manage().Cookies.GetCookieNamed($name)
            }"GetAllCookies"{
                $Driver.Manage().Cookies.AllCookies
            }
        }
    }
}

function New-SeCookie {
    <#
    .SYNOPSIS
    Returns a new cookie object
    
    .PARAMETER name
    Mandatory param, name of cookie
    
    .PARAMETER value
    Mandatory param, value for cookie
    
    .PARAMETER domain
    Cookies domain attribute
    
    .PARAMETER path
    Cookie path attribute
    
    .PARAMETER expiry
    Cookie expiry attribute
    
    .EXAMPLE
    $cookie = New-SeCookie -name "test1" -value "testValue"
    
    .EXAMPLE
    $cookieParams = [System.Collections.Generic.Dictionary[string, System.Object]]::new()
    $cookieParams["name"] = "CookieName"
    $cookieParams["value"] = "CookieValue"
    $cookieParams["domain"] = "CookieDomain"
    $cookie = New-SeCookie -CookieDictionary $cookieParams

    .NOTES
    General notes
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, ParameterSetName = "string")]
        [string]$name,
        [Parameter(Mandatory = $true, ParameterSetName = "string")]
        [string]$value,
        [Parameter(ParameterSetName = "string")]
        [string]$domain,
        [Parameter(ParameterSetName = "string")]
        [string]$path,
        [Parameter(ParameterSetName = "string")]
        [System.Nullable[datetime]]$expiry,
        [Parameter(Mandatory = $true, ParameterSetName = "dictionary")]
        [System.Collections.Generic.Dictionary[string, System.Object]]$CookieParams
    )
    process{
        if($PSCmdlet.ParameterSetName -eq "string"){
            New-Object OpenQA.Selenium.Cookie ($name, $value, $domain, $path, $expiry)
        }else{
            New-Object OpenQA.Selenium.Cookie ($CookieParams)
        }
    }
}