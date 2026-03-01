param(
    [string]$Port = 'COM3',
    [string]$Cmd  = 'KEY:A:50'
)

# シンプルな Notepad 前面化とシリアル送信スクリプト
$wshell = New-Object -ComObject WScript.Shell

# Notepad が動いていなければ起動
$np = Get-Process -Name notepad -ErrorAction SilentlyContinue
if (-not $np) {
    Start-Process notepad | Out-Null
    Start-Sleep -Milliseconds 300
    $np = Get-Process -Name notepad -ErrorAction SilentlyContinue
}

if (-not $np) {
    Write-Error 'Notepad process not found'
    exit 1
}

# AppActivate はプロセスIDまたはウィンドウタイトルで前面化
try {
    $wshell.AppActivate($np.Id) | Out-Null
} catch {
    # 失敗しても続行
}
Start-Sleep -Milliseconds 200

try {
    $sp = New-Object System.IO.Ports.SerialPort $Port,115200
    $sp.NewLine = "`n"
    $sp.ReadTimeout = 1000
    $sp.Open()
    $sp.WriteLine($Cmd)
    $sp.Close()
    Write-Output ('Sending to {0} : {1}' -f $Port,$Cmd)
    Write-Output 'Sent.'
} catch {
    $msg = $_.Exception.Message -replace "\r|\n", ' '
    Write-Error ('Serial error: {0}' -f $msg)
    if ($sp -and $sp.IsOpen) { $sp.Close() }
    exit 1
}
