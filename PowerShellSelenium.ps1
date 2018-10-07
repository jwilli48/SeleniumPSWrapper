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

function Start-SeChrome {
    <#
    .SYNOPSIS
    Returns a Selenium Chromedriver object
    .EXAMPLE
    $driver = Start-SeChrome -Maximized
    .EXAMPLE
    $driver = Start-SeChrome -Headless -MuteAudio
    #>
    [cmdletbinding()]
    [OutputType([OpenQA.Selenium.Chrome.Chromedriver])]
    param(
        [switch]$Headless,
        [switch]$MuteAudio,
        [switch]$Maximized,
        [switch]$Incognito
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
        if($incognito){
            $chrome_options.AddArgument("--incognito")
        }
        if ($chrome_options.Arguments -eq 0) {
            New-Object -TypeName OpenQA.Selenium.Chrome.ChromeDriver
        }
        else {
            New-Object -TypeName OpenQA.Selenium.Chrome.ChromeDriver($chrome_options)
        }
    }
}

function Start-SeFireFox {
    <#
    .SYNOPSIS
    Returns a Selenium FirefoxDriver object
    .EXAMPLE
    $driver = Start-SeFirefox -Maximized
    .EXAMPLE
    $driver = Start-SeFirefox -Headless -MuteAudio
    #>
    [cmdletbinding()]
    [OutputType()]
    param(
        [switch]$Headless,
        [switch]$MuteAudio,
        [switch]$Maximized,
        [switch]$Private
    )
    process {
        [OpenQA.Selenium.Firefox.FirefoxOptions]$firefox_options = New-Object OpenQA.Selenium.Firefox.FirefoxOptions
        if ($Headless) { $firefox_options.AddArgument("--headless")}
        if ($Maximized) { $firefox_options.AddArgument("--start-maximized")}
        if ($MuteAudio) { $firefox_options.SetPreference("media.volume_scale", "0.0")}
        if ($Private) { $firefox_options.AddArgument("-private")}
        if ($ProfileManager) {
            if ($NULL -ne $ProfileName) {
                $firefox_options.AddArgument("-P `"$ProfileName`"")
            }
            else {
                $firefox_options.AddArgument("-P")
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
    Takes in one or more WebDrivers and returns a custom object that contains the Url, Title, SessionID, CurrentWindowHandle (browser tab with focus), any other window handles (all other tabs), and then any capabilities of the driver. All of this data can be accessed from the driver object itself though so this may have limited use
    
    .PARAMETER driver
    The Selenium WebDriver object. This can be an array of drivers as well. 
    
    .EXAMPLE
    $cdriver = Start-SeChrome -Headless
    $fdriver = Start-SeFirefox -Headless
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
    
    .PARAMETER Async
    If set the scripts will be run as AsyncScripts
    
    .EXAMPLE
    Start-SeChrome | Invoke-Javascript -Script "while(!document.ready()); window.url = "https://gmail.com; return document.title;"
    
    .NOTES
    General notes
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, ValueFromPipeLine = $true)]
        [OpenQA.Selenium.Remote.RemoteWebDriver[]]$DriverList,
        [Parameter(Mandatory = $true)]
        [string[]]$Scripts,
        [switch]$Async
    )
    process {
        ForEach ($Driver in $DriverList) {
            if ($Async) {
                ForEach ($script in $scripts) {
                    try {
                        $driver.ExecuteAsyncScript($script);
                    }
                    catch {
                        Write-Verbose "ERROR: Invalid JavaScript used:`n$script"
                    }
                }
            }
            else {
                foreach ($script in $scripts) {
                    try {
                        $driver.ExecuteScript($script)
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
    Invoke-SeFindElement -Driver $driver -by CssSelector -Locator 'input[type="password"]'
    
    .NOTES
    This should only be used if you know the page has alreadyoaded completely as it will immediately search for the given element. See Invoke-SeWaitUntil to allow for the browser to wait for certain conditions to be filled. 
    #>
    [CmdletBinding()]
    [OutputType([OpenQA.Selenium.IWebElement])]
    param(
        [Parameter(Mandatory = $true, ParameterSetName = "Driver")]
        [OpenQA.Selenium.Remote.RemoteWebDriver[]]$DriverList,
        [Parameter(Mandatory = $true, ParameterSetName = "Element")]
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
        foreach($item in $itemList){
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
    The time to wait for the condition to be filled before throwing a TimeOut error. Default is 5 seconds
    
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
        foreach($driver in $DriverList){
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

function Close-SeDriver {
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
    General notes
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

function Set-SeTabFocus{
    <#
    .SYNOPSIS
    Switches driver tab to new tab
    
    .DESCRIPTION
    Changes tab to the given input, will throw an error if the tab does not exist
    
    .PARAMETER DriverList
    Parameter description
    
    .PARAMETER TabNumber
    Parameter description
    
    .EXAMPLE
    An example
    
    .NOTES
    General notes
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [OpenQA.Selenium.Remote.RemoteWebDriver[]]$DriverList,
        [Parameter(Mandatory = $true, ParameterSetName="Number")]
        $TabNumber
    )
    process{
        Foreach($driver in $DriverList){
            if($TabNumber -gt $driver.WindowHandles.Count){
                throw "Tab number can't be greater then the number of tabs"
            }
            $driver.SwitchTo().Window($driver.WindowHandles[$TabNumber])
        }
    }
}

function Get-SeScreenShot{
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
    
    .EXAMPLE
    Get-SeScreenshot -DriverList $a, $b -Format Png -DestinationDirectory "$home\desktop\temp" -FileBaseName browser_img
    #Will save two images browser_img_0.png and browser_img_1.png in the temp directory.
    
    .NOTES
    General notes
    #>
    [CmdletBinding(DefaultParameterSetName="DontSave")]
    param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [OpenQA.Selenium.Remote.RemoteWebDriver[]]$DriverList,
        [Parameter(Mandatory = $true, ParameterSetName="SaveAs")]
        [ValidateSet("Png", "Jpeg", "Gif", "Tiff", "Bmp")]
        $Format,
        [Parameter(Mandatory = $true, ParameterSetName="SaveAs")]
        $DestinationDirectory,
        [Parameter(Mandatory=$true, ParameterSetName="SaveAs")]
        [ValidateScript({
            if($_ -match ".*?\.[A-Za-z]+"){
                throw "Do not include a file extension as it will be added from the Format parameter"
            }else{
                $true
            }})]
        $FileBaseName
    )
    process{
        $image_number = 0
        foreach($Driver in $DriverList){
            $ScreenShot = $driver.GetScreenShot()
            if($PSCmdlet.ParameterSetName -eq "SaveAs"){
                if(-not (Test-Path $DestinationDirectory)){
                    New-Item -ItemType Directory -Path $DestinationDirectory    
                }
                $ScreenShot.SaveAsFile("$DestinationDirectory\$($FileBaseName)_$image_number.$Format", [OpenQA.Selenium.ScreenshotImageFormat]::$Format)
            }
            $image_number++
            $ScreenShot
        }
    }
}



function Get-SeElementScreenShot{
    <#
    .SYNOPSIS
    Tries to screen shot specific elements
    
    .DESCRIPTION
    This is not working. The offset for the location of the images is not correct based on the screen shot of the page and it is cropping the wrong part of the image.
    
    .PARAMETER DriverList
    Parameter description
    
    .PARAMETER ElementList
    Parameter description
    
    .PARAMETER Format
    Parameter description
    
    .PARAMETER DestinationDirectory
    Parameter description
    
    .PARAMETER FileBaseName
    Parameter description
    
    .EXAMPLE
    An example
    
    .NOTES
    Not working currently very well (crops image incorrectly). 
    When I have time this is probably a better way to do it:
    $Screenshot = $driver.GetScreenShot()
    $ScreenShot.Saveas("$home\temp\$tempname.format", [OpenQA.Selenium.ScreenshotImageFormat]::$Format)

    [System.Drawing.Image]$img = [System.Drawing.Image]::FromFile($tempfile)

    $rect = New-Object System.Drawing.Rectangle($element.location.X, $element.location.Y, $element.size.Width, $element.size.Height) #Maybe my issue is this isn't in the correct order.

    [System.Drawing.Bitmap]$BmpImage = New-Object System.Drawing.Bitmap($img)
    $crop = $BmpImage.Clone($Rect, $BmpImage.PixelFormat)
    $crop.Save("$home\$temp\$tempname.$format", [System.Drawing.Imaging.ImageFormat]::$format)
    #>
    [CmdletBinding(DefaultParameterSetName="DontSave")]
    [OutputType([System.Drawing.Bitmap])]
    param(
        [Parameter(Mandatory=$true)]
        [OpenQA.Selenium.Remote.RemoteWebDriver[]]$DriverList,
        [Parameter(Mandatory = $true)]
        [OpenQA.Selenium.IWebElement[]]$ElementList,
        [Parameter(Mandatory = $true, ParameterSetName="SaveAs")]
        [ValidateSet("Png", "Jpeg", "Gif", "Tiff", "Bmp")]
        $Format,
        [Parameter(Mandatory = $true, ParameterSetName="SaveAs")]
        [string]$DestinationDirectory,
        [Parameter(Mandatory=$true, ParameterSetName="SaveAs")]
        [ValidateScript({
            if($_ -match ".*?\.[A-Za-z]+"){
                throw "Do not include a file extension as it will be added from the Format parameter"
            }else{
                $true
            }})]
        [string]$FileBaseName
    )
    process{
        $ElementList = $ElementList | Select-Object -Unique
        $ScreenShots = Get-SeScreenShot -DriverList $DriverList
        $image_num = 0 
        foreach($ScreenShot in $ScreenShots){
            foreach($Element in $ElementList){
                [System.Drawing.Bitmap] $image = New-Object System.Drawing.Bitmap((New-Object System.IO.MemoryStream ($ScreenShot.AsByteArray, $ScreenShot.AsByteArray.Count)))

                [System.Drawing.Rectangle] $crop = New-Object System.Drawing.Rectangle($element.location.X, $element.location.Y, $element.size.width, $element.size.height)

                $image = $image.clone($crop, $image.PixelFormat)
                if($PSCmdlet.ParameterSetName -eq "SaveAs"){
                    if(-not (Test-Path $DestinationDirectory)){
                        New-Item -ItemType Directory -Path $DestinationDirectory    
                    }
                    $image.Save("$($DestinationDirectory)\$($FileBaseName)_$($image_num).$Format", [System.Drawing.Imaging.ImageFormat]::$Format)                    
                }
                $image_num++
                $image
            }
        }
    }
}