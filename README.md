# SeleniumPSWrapper
PowerShell Wrapper for Selenium

My attempt to make as comprehensive a PS wrapper for Selenium as I can, specifically for Firefox and Chrome browsers.

## TODO
    -Missing Manage class implementation (for cookies / logs)
    -I want to add more helper functions (element attributes)
        
    
    General:
        -OpenNewTab
            -With URL
            -Switch To
                -If true switch focus to new tab
        
    Manage:
        -Cookies
            -DeleteAll
            -DeleteName
            -DeleteCookie
            -GetName
            -GetAll
        -Logs
            -GetLog
            -AvailableLogs
        -Windows
            -Maximize
            -Minimize
            -Fullscreen
            -Properties
                -Position
                -Size
    Elements:
        -Get attribute