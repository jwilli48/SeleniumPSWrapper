function Get-DynamicParam {
    <#
    .SYNOPSIS
    Sets and returns an object that can be added to the RuntimeDefinedParameterDictionary object
    .PARAMETER Name
    Name of the dynamic parameter
    .PARAMETER Type
    The type of the dynamic parameter
    .PARAMETER Mandatory
    Sets Mandatory attribute to true
    .PARAMETER FromPipeline
    Will set the Parameter Attribute ValueFromePipeline to true
    .PARAMETER Position
    Sets the position
    .PARAMETER ValidateSet
    An array of strings to validate the parmameter input
    .PARAMETER ValidateScript
    Takes a script block as a param. If needed you can convert a string to a scriptblock pretty easily: 
        $script = [scriptblock]::Create($string)
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

    The name for this must match the same name as you are using for the first parameter of the RuntimeParameterDictionary.
    #>
    [CmdletBinding()]
    [OutputType([System.Management.Automation.RuntimeDefinedParameter])]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name,
        [Parameter(Mandatory = $true)]
        [System.Type]$Type,
        [switch]$Mandatory,
        [switch]$FromPipeline,
        $Position,
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