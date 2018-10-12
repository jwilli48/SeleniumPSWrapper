& {Add-Type -A 'System.IO.Compression.Filesystem'
    [IO.Compression.ZipFIle]::ExtractToDirectory($FileToExtract, $FileDestination)}