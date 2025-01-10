#SJIS $Workfile: OdaiAday.ps1 $$Revision: 5 $$Date: 25/01/09 22:18 $
#$NoKeywords: $

Add-Type -AssemblyName System.Windows.Forms

# フォルダのパスを指定
$folderPath = "N:\ネタの種\書籍"

# 指定されたフォルダ内の.txtファイルだけを取得
$txtFiles = Get-ChildItem -Path $folderPath -File | Where-Object { $_.Extension -eq ".txt" }

$foundToc = $false
$randomLine = $null
$readStatus = $null

for ($i=0;-not $foundToc -and  $i -lt 100;$i ++) {
    # .txtファイルが存在する場合、ランダムに1つ選択
    if ($txtFiles.Count -gt 0) {
        $randomTextFile = $txtFiles | Get-Random

        # ファイルの内容を読み込んでブロックに分割
        $content = Get-Content -Path $randomTextFile.FullName -Raw
        $blocks = $content -split "(`r?`n){3,}"

        # (区切り文字を含む)3つ目のブロックに"読了"が含まれるか確認（ブロックが3つ以上ある場合）
        if ($blocks.Length -ge 3) {
            $readStatus = if ($blocks[2] -match "読了") { "既読" } else { "未読" }
        }

        # "^目[ 　]*次"を含むブロックから長さが0より長い行をランダムに選択
        $tocBlock = $blocks | Where-Object { $_ -match "^目[ 　]*次" }
        if ($tocBlock) {
            $lines = $tocBlock -split "`r?`n" | Where-Object { $_.Length -gt 0 -and $_ -notmatch("(目[ 　]*次|はじめに|まえがき|おわりに|まとめ|あとがき)")}
            if ($lines.Length -gt 0) {
                $foundToc = $true
                $randomLine = $lines | Get-Random
            }
        }else{
            $randomTextFile = $null
            $randomLine = $null
            $readStatus = $null
        }
    }
}
if($foundToc){
    # 結果をダイアログ表示
    $fileName = [System.IO.Path]::GetFileNameWithoutExtension($randomTextFile)
    $message = "$fileName($readStatus)`n`n$randomLine"
    Write-EventLog -LogName "OdaiAdayLog" -Source "OdaiAday_ps1" -EntryType Information -EventID 0 -Message $message # メッセージをイベントログに記録
    [System.Windows.Forms.MessageBox]::Show($message, "今日のお題", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    
}else{
    [System.Windows.Forms.MessageBox]::Show("目次がありません。", "今日のお題", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
}