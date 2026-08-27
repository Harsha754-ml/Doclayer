param(
    [Parameter(Mandatory = $true)] [string] $Source,
    [Parameter(Mandatory = $true)] [string] $Output
)

$ErrorActionPreference = 'Stop'

function HashStr($s) {
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($s)
    $md5 = [System.Security.Cryptography.MD5]::Create()
    $h = $md5.ComputeHash($bytes)
    $sb = New-Object System.Text.StringBuilder
    foreach ($b in $h) { [void]$sb.Append($b.ToString('x2')) }
    return $sb.ToString().Substring(0, 16)
}

function Sanitize($s) {
    $s = [System.IO.Path]::GetFileNameWithoutExtension($s)
    $s = $s -replace '[^a-zA-Z0-9_]', '_'
    if ($s -match '^\d') { $s = '_' + $s }
    return $s
}

$counter = 0
$compIds = [System.Collections.Generic.List[string]]::new()
$sb = [System.Text.StringBuilder]::new()

function EmitDir($dir, $indent) {
    $files = Get-ChildItem -Path $dir -File
    $subdirs = Get-ChildItem -Path $dir -Directory

    foreach ($f in $files) {
        $counter++
        $compId = 'cmp_' + (HashStr $f.FullName)
        $fileId = 'file_' + (HashStr $f.FullName)
        $compIds.Add($compId)
        [void]$sb.AppendLine("$indent<Component Id=`"$compId`" Guid=`"*`">")
        [void]$sb.AppendLine("$indent  <File Id=`"$fileId`" Source=`"$($f.FullName)`" KeyPath=`"yes`" />")
        [void]$sb.AppendLine("$indent</Component>")
    }

    foreach ($d in $subdirs) {
        $dirId = 'dir_' + (HashStr $d.FullName)
        [void]$sb.AppendLine("$indent<Directory Id=`"$dirId`" Name=`"$($d.Name)`">")
        EmitDir $d.FullName "$indent  "
        [void]$sb.AppendLine("$indent</Directory>")
    }
}

[void]$sb.AppendLine('<Wix xmlns="http://wixtoolset.org/schemas/v4/wxs">')
[void]$sb.AppendLine('  <Fragment>')
[void]$sb.AppendLine('    <DirectoryRef Id="INSTALLFOLDER">')
EmitDir $Source '      '
[void]$sb.AppendLine('    </DirectoryRef>')
[void]$sb.AppendLine('  </Fragment>')
[void]$sb.AppendLine('  <Fragment>')
[void]$sb.AppendLine('    <ComponentGroup Id="AppFiles">')
foreach ($c in $compIds) {
    [void]$sb.AppendLine("      <ComponentRef Id=`"$c`" />")
}
[void]$sb.AppendLine('    </ComponentGroup>')
[void]$sb.AppendLine('  </Fragment>')
[void]$sb.AppendLine('</Wix>')

Set-Content -Path $Output -Value $sb.ToString() -Encoding UTF8
Write-Host "Generated $Output with $($compIds.Count) components."
