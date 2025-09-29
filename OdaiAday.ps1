#SJIS $Workfile: OdaiAday.ps1 $$Revision: 7 $$Date: 25/09/28 15:09 $
#$NoKeywords: $

Add-Type -AssemblyName System.Windows.Forms

# �t�H���_�̃p�X���w��
$folderPath = "N:\�l�^�̎�\����"

# �w�肳�ꂽ�t�H���_����.txt�t�@�C���������擾
$txtFiles = Get-ChildItem -Path $folderPath -File | Where-Object { $_.Extension -eq ".txt" }

$foundToc = $false
$randomLine = $null
$readStatus = $null

for ($i=0;-not $foundToc -and  $i -lt 100;$i ++) {
    # .txt�t�@�C�������݂���ꍇ�A�����_����1�I��
    if ($txtFiles.Count -gt 0) {
        $randomTextFile = $txtFiles | Get-Random

        # �t�@�C���̓��e��ǂݍ���Ńu���b�N�ɕ���
        $content = Get-Content -Path $randomTextFile.FullName -Raw
        $blocks = $content -split "(`r?`n){3,}"

        # (��؂蕶�����܂�)3�ڂ̃u���b�N��"�Ǘ�"���܂܂�邩�m�F�i�u���b�N��3�ȏ゠��ꍇ�j
        if ($blocks.Length -ge 3) {
            $readStatus = if ($blocks[2] -match "�Ǘ�") { "����" } else { "����" }
        }

        # "^��[ �@]*��"���܂ރu���b�N���璷����0��蒷���s�������_���ɑI��
        $tocBlock = $blocks | Where-Object { $_ -match "^��[ �@]*��" }
        if ($tocBlock) {
            $lines = $tocBlock -split "`r?`n" | Where-Object { $_.Length -gt 0 -and $_ -notmatch("(��[ �@]*��|�͂��߂�|�܂�����|������|�܂Ƃ�|���Ƃ���|�Q�l|�t�^|��������|�p����)")}
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
    # ���ʂ��_�C�A���O�\��
    $fileName = [System.IO.Path]::GetFileNameWithoutExtension($randomTextFile)
    $message = "$fileName($readStatus)`n`n$randomLine"
    Write-EventLog -LogName "OdaiAdayLog" -Source "OdaiAday_ps1" -EntryType Information -EventID 0 -Message $message # ���b�Z�[�W���C�x���g���O�ɋL�^
    [System.Windows.Forms.MessageBox]::Show($message, "�����̂���", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)

    #----- �����e��E-Mail(hi-ho����g��)����
    # SMTP�N���C�A���g���쐬
    $smtp = New-Object System.Net.Mail.SmtpClient("smtp.example.com", 587)
    $smtp.EnableSsl = $false # $true
    $smtp.Credentials = New-Object System.Net.NetworkCredential("yourmail@example.com", "yourpassword")

    # ���[�����e���쐬
    $emailmsg = New-Object System.Net.Mail.MailMessage
    $emailmsg.From = "yourmail@example.com"
    $emailmsg.To.Add("receiver@example.com")
    $emailmsg.Subject = "�����̂���"
    $emailmsg.Body = $message + [Environment]::NewLine

    # ���M
    $smtp.Send($emailmsg)

}else{
    [System.Windows.Forms.MessageBox]::Show("�ڎ�������܂���B", "�����̂���", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
}