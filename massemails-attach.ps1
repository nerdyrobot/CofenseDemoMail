####### Mailscript PowerShell script to send emails to cofensedemo.com domain ########
### Generates random sender email address and domain - The output is formatted abcdefghijk@domain1234.org
function Generate-Sender {

    $userNameLength = Get-Random -Minimum 4 -Maximum 11
    $userName = (65..90) + (97..122) | Get-Random -Count $userNameLength | ForEach-Object {[char]$_}

    $domainNumber = Get-Random -Minimum 1000 -Maximum 9999
    
    $sender = (-join $username) + "@domain$domainNumber.org"

    return $sender

}
### Generates semi-random subject consisting of the below words preceded by text and date/time
### Format AB 01/23/45 12:34 AM 
### Regex to find this subject to be used in Triage rule ‘(ABC).\d{1,2}\/\d{1,2}\/\d{2}.\d{1,2}:\d{1,2}.(AM|PM)’
function Generate-Subject {

    $dictionary = @('Up-Sell', 'Skill set', 'Non-Traditional Management', 'Cross Sell', 'Reinvigorate', 'Run it up the flagpole', 'Revenue', 'Knowledge Base', 'Touch Base', 'Laser-focused', 'Utilize', 
        'Game Plan', 'Strategize', 'Differentiation', 'Herding Cats', 'Cutting Edge', 'Ramp Up', 'Turnkey', 'Scope', 'Content Management', 'Green', 'Competitive', 'Guidance', 'Asset', 'Discovery', 'One to one', 
        'Re-engagement', 'Heavy lifting', 'Heads up', 'Back-end', 'Service Oriented', 'Engage', 'Monoply', 'Goal oriented', 'Facilitate', 'Monetize', 'Revisit', 'Center of Excellence', 'E-commerce', 'Brain Storm', 
        'Mind Shower', 'Upside', 'Business Opportunity', 'Demographic', 'ETA', 'Cost', 'Walk the Talk', 'B2B', 'Touchpoints')

    $date = Get-date -Uformat %H:%M:%S
    $subject = $dictionary | Get-Random -Count 3 ### This number determines the number of word pairs in the email subject
    return ("ABC") + ($date -join ' ') + (" ") + ($subject -join ' ') ###Change “ABC“ for custom subject prefix
}

$emailBody = @"
<!DOCTYPE HTML>
<html>
<head>
<title>Test Email</title>
</head>
<body>
<h4>There are super important things in the attachment. You will want to open it right away...</h4>
<h4>Don't even worry about what it might be... </h4>
<h4>Sincerly,</h4>
<h4>- "Someone who seems important"</h4>
</body>
</html>
"@

$smtpServer = "cofensedemo-com.mail.protection.outlook.com"

$file_dictionary = @('0ciavsIp6PRn.iso',
                    '0HPjNn8YgUNB.iso',
                    '1cKVGjenhISF.iso',
                    '6kaRtbG4ArOg.iso',
                    '6koSoZ5aZTRQ.iso',
                    '6N6NJo3NFkM1.iso',
                    '7rFudsVdR998.iso',
                    '94ffQcryspvl.iso',
                    'aLE2lbPQTCuF.iso',
                    'BAFqSV2uWv4j.iso',
                    'BvMx1te2fLu9.iso',
                    'cSpNN12e1Vqp.iso',
                    'dEbgvEmIeETw.iso',
                    'fCjNafGRhDrq.iso',
                    'fnM6Cmhf5nlN.iso',
                    'fxGmJiKlBXgZ.iso',
                    'G6yk2F5ykLtq.iso',
                    'haXpATF3pBVd.iso',
                    'HfTFtI9bUoYh.iso',
                    'hkQd9yUPRzkR.iso',
                    'jmipU0nrUASi.iso',
                    'Jqfak9jaJT9u.iso',
                    'k4OxzT2B7EAk.iso',
                    'l6xDOYHpJ2IQ.iso',
                    'lgts8hfE5FCs.iso',
                    'LJRbI9Oi7N1Z.iso',
                    'luPQ3hFKiE3Y.iso',
                    'mMfEYrPveCsq.iso',
                    'MYgV0LSsb2JK.iso',
                    'nsbfoAnEKv4S.iso',
                    'P9ci7xsiNuTz.iso',
                    'pIkrXiX04xHi.iso',
                    'pxezaMjVlxNj.iso',
                    'qnKsHlikbQUr.iso',
                    'RglPfyHyNrTb.iso',
                    'rM561JdZmULH.iso',
                    'rsIbX0o9m0i4.iso',
                    'S5nFDjSZ8ZKi.iso',
                    'SCzEcN4trdab.iso',
                    'TaZ9sCgWk4N4.iso',
                    'tpBq1Gf6Xv47.iso',
                    'tPnfelNCZYVY.iso',
                    'UJh4k3FmszZ6.iso',
                    'v5k990SbFKv9.iso',
                    'XmyTfKPTjUN6.iso',
                    'xQBXbTjagCUl.iso',
                    'xuRsbqquw8XG.iso',
                    'z2QipUtIFspB.iso',
                    'zNSS2J8DvfTN.iso',
                    'ZoKI99iOfko2.iso')

$file = $file_dictionary | Get-Random -Count 1


### Selects recipients from the .txt file referenced in the recipients folder in mailscript folder
### Generates an emails for each recipient in the list
$allUsers = Get-Content -Path "$PSScriptRoot\recipient\recipients.txt"

$allUsers | Where-Object {$_} | ForEach-Object {
    
    $to = $_
    $from = Generate-Sender
    $subject = Generate-Subject
    
    Write-Host "Sending email from $from to $to with subject `"$subject`""

### Uncomment ‘-Attachments $attachment’ below to enable attachments
    Send-MailMessage -From $from -To $to -Subject $subject -Body $emailBody -SmtpServer $smtpServer -BodyAsHtml -UseSsl -Attachments "$PSScriptRoot\attachment\$file"

}
