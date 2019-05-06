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

    $date = Get-date -Format g
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

$file_dictionary = @('0ciavsIp6PRn.txt',
                    '0HPjNn8YgUNB.txt',
                    '1cKVGjenhISF.txt',
                    '6kaRtbG4ArOg.txt',
                    '6koSoZ5aZTRQ.txt',
                    '6N6NJo3NFkM1.txt',
                    '7rFudsVdR998.txt',
                    '94ffQcryspvl.txt',
                    'aLE2lbPQTCuF.txt',
                    'BAFqSV2uWv4j.txt',
                    'BvMx1te2fLu9.txt',
                    'cSpNN12e1Vqp.txt',
                    'dEbgvEmIeETw.txt',
                    'fCjNafGRhDrq.txt',
                    'fnM6Cmhf5nlN.txt',
                    'fxGmJiKlBXgZ.txt',
                    'G6yk2F5ykLtq.txt',
                    'haXpATF3pBVd.txt',
                    'HfTFtI9bUoYh.txt',
                    'hkQd9yUPRzkR.txt',
                    'jmipU0nrUASi.txt',
                    'Jqfak9jaJT9u.txt',
                    'k4OxzT2B7EAk.txt',
                    'l6xDOYHpJ2IQ.txt',
                    'lgts8hfE5FCs.txt',
                    'LJRbI9Oi7N1Z.txt',
                    'luPQ3hFKiE3Y.txt',
                    'mMfEYrPveCsq.txt',
                    'MYgV0LSsb2JK.txt',
                    'nsbfoAnEKv4S.txt',
                    'P9ci7xsiNuTz.txt',
                    'pIkrXiX04xHi.txt',
                    'pxezaMjVlxNj.txt',
                    'qnKsHlikbQUr.txt',
                    'RglPfyHyNrTb.txt',
                    'rM561JdZmULH.txt',
                    'rsIbX0o9m0i4.txt',
                    'S5nFDjSZ8ZKi.txt',
                    'SCzEcN4trdab.txt',
                    'TaZ9sCgWk4N4.txt',
                    'tpBq1Gf6Xv47.txt',
                    'tPnfelNCZYVY.txt',
                    'UJh4k3FmszZ6.txt',
                    'v5k990SbFKv9.txt',
                    'XmyTfKPTjUN6.txt',
                    'xQBXbTjagCUl.txt',
                    'xuRsbqquw8XG.txt',
                    'z2QipUtIFspB.txt',
                    'zNSS2J8DvfTN.txt',
                    'ZoKI99iOfko2.txt')

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
