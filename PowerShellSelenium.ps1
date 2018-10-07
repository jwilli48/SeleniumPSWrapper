<#
Summer 2018
Joshua Williamson
Functions to ease use of Selenium Webdriver to be used inside of Windows PowerShell for Chrome browser automations.

.NOTE
    Dynamic Params: If you have a custom type as a param, the dynamic params will not show up visually with tab completion / Intellisense after that one has been set in the function.
    FIX: Make the custom object a dynamic parm that is always created
#>

#Load needed modules
[System.Reflection.Assembly]::LoadFrom("$home\Desktop\SeleniumTest\WebDriver.dll") | Out-Null
[System.Reflection.Assembly]::LoadFrom("$home\Desktop\SeleniumTest\WebDriver.Support.dll") | Out-Null
[System.Reflection.Assembly]::LoadFrom("$home\Desktop\SeleniumTest\Selenium.WebDriverBackedSelenium.dll") | Out-Null

function Get-DynamicParam {
    <#
    .SYNOPSIS
    Sets and returns an object that can be added to the RuntimeDefinedParameterDictionary object
    .PARAMETER name
    Name of the dynamic parameter
    .PARAMETER type
    The type of the dynamic parameter
    .PARAMETER Mandatory
    Sets Mandatory attribute to true
    .PARAMETER FromPipeline
    Will set the Parameter Attribute ValueFromePipeline to true
    .PARAMETER position
    Sets the position
    .PARAMETER ValidateSet
    An array of strings to validate the parmameter input
    .PARAMETER HelpMessage
    Will be assigned as the Parameters help message
    .PARAMETER SetName
    The ParameterSetName value for the Dynamic Param
    .PARAMETER ValueFromPipelineByPropertyName
    Sets property to true
    .PARAMETER ValueFromRemainingArguments
    Sets property to true
    .PARAMETER DontShow
    Sets property to true
    .PARAMETER DefaultValue
    Sets a default value, I have not yet fully tested how well it would work.
    .INPUTS
    None can be piped into this function.
    .OUTPUTS
    Will return a System.Management.Automation.RuntimeDefinedParameter
    .EXAMPLE
    function test{
        param(
            [string]$name
        )
        DynamicParam{
            $RuntimeParameterDictionary = New-Object System.Management.Automation.RuntimeDefinedParameterDictionary
            $RuntimeParameterDictionary.Add("ParamName", (Get-DynamicParam -name "ParamName" -type string -mandatory -FromPipeline -position 2 -ValidateSet "Value1", "Value2" -HelpMessage "You use it like this..." -SetName "TestCase" -ValidateScript {Write-Host "Test"}))
            $RuntimeParameterDictionary
        }
        begin{
            #This standard block of code loops through bound parameters...
            #If no corresponding variable exists, one is created
            #Get common parameters, pick out bound parameters not in that set
                Function _temp { [cmdletbinding()] param() }
                $BoundKeys = $PSBoundParameters.keys | Where-Object { (get-command _temp | Select-Object -ExpandProperty parameters).Keys -notcontains $_}
                foreach($param in $BoundKeys)
                {
                    if (-not ( Get-Variable -name $param -scope 0 -ErrorAction SilentlyContinue ) )
                    {
                        New-Variable -Name $Param -Value $PSBoundParameters.$param
                        Write-Verbose "Adding variable for dynamic parameter '$param' with value '$($PSBoundParameters.$param)'"
                    }
                }
            #Appropriate variables should now be defined and accessible
                Get-Variable -scope 0
        }
        process{
            Write-Host $ParamName
        }
    }
    .NOTES
    The function _temp in the first example can be extremely helpful when wanting to use the DynamicParams within your function blocks as it creates variables for each of them with their given names. Without that function you would have to do use $PSBoundParameters[$ParameterName] in place of that. You can set $ParmName = $PSBoundParameters[$ParameterName] as well.
    #>
    [CmdletBinding()]
    [OutputType([System.Management.Automation.RuntimeDefinedParameter])]
    param(
        [Parameter(Mandatory = $true)]
        [string]$name,
        [Parameter(Mandatory = $true)]
        [System.Type]$Type,
        [switch]$mandatory,
        [switch]$FromPipeline,
        $position,
        [string[]]$ValidateSet,
        [System.Management.Automation.ScriptBlock]$ValidateScript,
        $HelpMessage,
        [string]$SetName,
        [switch]$ValueFromPipelineByPropertyName,
        [switch]$ValueFromRemainingArguments,
        [switch]$DontShow,
        $DefaultValue
    )
    process {
        $AttributeCollection = New-Object 'System.Collections.ObjectModel.Collection[System.Attribute]'

        $ParameterAttributes = New-Object System.Management.Automation.ParameterAttribute
        if ($mandatory) { $ParameterAttributes.Mandatory = $true}
        if ($FromPipeline) { $ParameterAttributes.ValueFromPipeline = $true}
        if ($ValueFromPipelineByPropertyName) { $ParameterAttributes.ValueFromPipelineByPropertyName = $true}
        if ($ValueFromRemainingArguments) { $ParameterAttribute.ValueFromRemainingArguments = $True}
        if ($DontShow) { $ParameterAttributes.DontShow = $True }
        if ($NULL -ne $position) { $ParameterAttributes.Position = $position }
        if ($NULL -ne $SetName) {$ParameterAttributes.ParameterSetName = $SetName}
        if ($NULL -ne $HelpMessage) {$ParameterAttributes.HelpMessage = $HelpMessage}
        $AttributeCollection.add($ParameterAttributes)
        if ($NULL -ne $ValidateSet) {
            $ValidateSetAttribute = New-Object System.Management.Automation.ValidateSetAttribute -ArgumentList $ValidateSet
            $AttributeCollection.add($ValidateSetAttribute)
        }
        if ($NULL -ne $ValidateScript) {
            $ValidateScriptAttribute = New-Object System.Management.Automation.ValidateScriptAttribute -ArgumentList $ValidateScript
            $AttributeCollection.add($ValidateScriptAttribute)
        }
        $RuntimeParameter = New-Object -TypeName System.Management.Automation.RuntimeDefinedParameter -ArgumentList @($name, $type, $AttributeCollection)
        if ($NULL -ne $DefaultValue) {
            $RuntimeParameter.Value = $DefaultValue
        }
        $RuntimeParameter
    }
}

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
        [switch]$Maximized
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
       [OpenQA.Selenium.Remote.RemoteWebDriver[]]$driver
    )
    process {
        #Need to make this a foreach loop#
        [PSCustomObject] @{
            Url                 = $driver.url
            Title               = $driver.title
            SessionId           = $driver.sessionId
            CurrentWindowHandle = $driver.CurrentWindowHandle
            WindowHandles       = $driver.WindowHandles
            Capabilities        = $driver.Capabilities
        }
    }
}

function Invoke-SeJavaScript {
    <#
    .SYNOPSIS
    Run JavaScript on browser window that has focus
    
    .DESCRIPTION
    Runs one or more JavaScript scripts that are given as input on the browser. It will attempt to execute the scripts (they can be Async or not depending on if you set the switch value) and if it fails it will print an error message that includes the script that fails (if you have the verbose switch set, otherwise it should fail silently).
    
    .PARAMETER driver
    Selenium webdriver that you will execute the JavaScript on. 
    
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
        [OpenQA.Selenium.Remote.RemoteWebDriver]$driver,
        [Parameter(Mandatory = $true)]
        [string[]]$scripts,
        [switch]$Async
    )
    process {
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
        [OpenQA.Selenium.Remote.RemoteWebDriver]$driver,
        [Parameter(Mandatory = $true, ParameterSetName = "Element")]
        [OpenQA.Selenium.IWebElement]$element,
        [Parameter(Mandatory = $true)]
        [ValidateSet("TagName", "ClassName", "ID", "LinkText", "PartialLinkText", "CssSelector", "XPath", "Name")]
        $by,
        [Parameter(Mandatory = $true, HelpMessage = "Input the selector that will match your choosen search method")]
        $locator
    )
    process {
        if ($PSCmdlet.ParameterSetName -eq "Driver") {
            $search = $driver
        }
        else {
            $search = $element
        }
        $search.FindElements([OpenQA.Selenium.By]::$by($locator))
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
    The condition to wait for. This will define various other dynamic parameters needed once it has been set. 
    
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
    [cmdletbinding(DefaultParameterSetName = "By")]
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
        $RuntimeParameterDictionary.Add('Driver', (Get-DynamicParam -Name "Driver" -type "OpenQA.Selenium.Remote.RemoteWebDriver" -mandatory -FromPipeline ))
        
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
