#SJIS $Workfile: OdaiAday.ps1 $$Revision: 7 $$Date: 25/09/28 15:09 $
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
            $lines = $tocBlock -split "`r?`n" | Where-Object { $_.Length -gt 0 -and $_ -notmatch("(目[ 　]*次|はじめに|まえがき|おわりに|まとめ|あとがき|参考|付録|巻末資料|用語解説)")}
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

    #----- 同内容をE-Mail(hi-hoから携帯)する
    # SMTPクライアントを作成
    $smtp = New-Object System.Net.Mail.SmtpClient("smtp.example.com", 587)
    $smtp.EnableSsl = $false # $true
    $smtp.Credentials = New-Object System.Net.NetworkCredential("yourmail@example.com", "yourpassword")

    # メール内容を作成
    $emailmsg = New-Object System.Net.Mail.MailMessage
    $emailmsg.From = "yourmail@example.com"
    $emailmsg.To.Add("receiver@example.com")
    $emailmsg.Subject = "今日のお題"
    $emailmsg.Body = $message + [Environment]::NewLine

    # 送信
    $smtp.Send($emailmsg)

}else{
    [System.Windows.Forms.MessageBox]::Show("目次がありません。", "今日のお題", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
}