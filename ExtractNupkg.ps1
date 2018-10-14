function Expand-NupkgFile{
    [Cmdletbinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,
        [Parameter(Mandatory = $true)]
        [string]$Destination
    )
    & {Add-Type -A 'System.IO.Compression.Filesystem'
        [IO.Compression.ZipFIle]::ExtractToDirectory($Path, $Destination)}
}
