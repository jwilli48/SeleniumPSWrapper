# SeleniumPSWrapper
PowerShell Wrapper for Selenium

My attempt to make as comprehensive a PS wrapper for Selenium as I can, specifically for Firefox and Chrome browsers.

Included in this repository:
* ChromeDriver.exe 
* Geckodriver.exe (FireFox)
* Selenium .dll files 
* WebDriver.chm, which is complied HTML file that has all of the documentation for selenium which you may need to unblock to make it actually work 
* Helper function Get-DynamicParam.ps1
* Expand-Nupkg funciton that will extract a .nupkg file if you want to update the selenium (as the .nupkg is how it is stored for C# and this is virtually just a bunch of POSH functions for the C# selenium classes and methods).
* The module, PowerShellSelenium.psm1

## Resources
I included some relevant links along with some of the functions which I will include here as well:
* [Selenium HomePage](https://www.seleniumhq.org/)
* [Selenium Documentation](https://seleniumhq.github.io/selenium/docs/api/dotnet/index.html) (This is an online / more updated version of the WebDriver.chm file)
* [Selenium Downloads](https://www.seleniumhq.org/download/)
    * You will want to download the version for C#. You will then most likely have to extract the .nupkg files to get the .dll files needed (that are then imported into PowerShell with the command at the top of the module).
        * The .dll files will be somewhere in the folders that come from extracting the .nupkg, some other documentation will also be extracted.
    * You can see all of the driver options to download as well. I will try to add creation commands for the other drivers, but if I haven't you should only need to create a New-SeDriver function for that browser and then the rest of the commands should work (though there may be extra capabilities that other drivers can do that FireFox and Chrome can't / some capabilities may not work the same)
* [ChromeDriver HomePage](http://chromedriver.chromium.org/)
* [ChromeDriver Download](http://chromedriver.chromium.org/downloads)
* [Chrome Commands Resource](https://chromedevtools.github.io/devtools-protocol/) (for Invoke-SeChromeCommand)
* [Chrome Command Switches ](https://peter.sh/experiments/chromium-command-line-switches/) (for New-SeChrome)
* [FireFox Geckodriver Download](https://github.com/mozilla/geckodriver/releases)
* [FireFox Command Switches](https://developer.mozilla.org/en-US/docs/Mozilla/Command_Line_Options#Original_Document_Information) (for New-SeFirefox)

## TODO