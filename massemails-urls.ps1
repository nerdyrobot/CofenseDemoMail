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
### Regex to find this subject to be used in Triage rule ‘(AB).\d{1,2}\/\d{1,2}\/\d{2}.\d{1,2}:\d{1,2}.(AM|PM)’
function Generate-Subject {

    $dictionary = @('Up-Sell', 'Skill set', 'Non-Traditional Management', 'Cross Sell', 'Reinvigorate', 'Run it up the flagpole', 'Revenue', 'Knowledge Base', 'Touch Base', 'Laser-focused', 'Utilize', 
        'Game Plan', 'Strategize', 'Differentiation', 'Herding Cats', 'Cutting Edge', 'Ramp Up', 'Turnkey', 'Scope', 'Content Management', 'Green', 'Competitive', 'Guidance', 'Asset', 'Discovery', 'One to one', 
        'Re-engagement', 'Heavy lifting', 'Heads up', 'Back-end', 'Service Oriented', 'Engage', 'Monoply', 'Goal oriented', 'Facilitate', 'Monetize', 'Revisit', 'Center of Excellence', 'E-commerce', 'Brain Storm', 
        'Mind Shower', 'Upside', 'Business Opportunity', 'Demographic', 'ETA', 'Cost', 'Walk the Talk', 'B2B', 'Touchpoints')

    $date = Get-date -Format g
    $subject = $dictionary | Get-Random -Count 3 ### This number determines the number of word pairs in the email subject
    return ("AB ") + ($date -join ' ') + (" ") + ($subject -join ' ') ###Change “AB “ for custom subject prefix

}
### Domains used for URL based scenarios feel free to add
### Format ‘HyperLinkText’ = ‘http/s://url.whatever’; <- semicolon required if not the last line in the list
$domains = @{
    'Phished by the Prince' = 'https://www.nigerianprince.com';
    'Commander Highland' = 'https://www.Commanderhighland.com';
    'GOT Spoilers' = 'https://www.gotspoilers.com';
    'GOT Season 1 Spoilers' = 'https://www.gotspoilersS1.com';
    'GOT Season 2 Spoilers' = 'https://www.gotspoilersS2.com';
    'GOT Season 3 Spoilers' = 'https://www.gotspoilersS3.com';
    'GOT Season 4 Spoilers' = 'https://www.gotspoilersS4.com';
    'GOT Season 5 Spoilers' = 'https://www.gotspoilersS5.com';
    'GOT Season 6 Spoilers' = 'https://www.gotspoilersS6.com';
    'GOT Season 7 Spoilers' = 'https://www.gotspoilersS7.com';
    'GOT Season 8 Spoilers' = 'https://www.gotspoilersS8.com';
    'In your base' = 'https://www.inyourbase.com';
    'AccountAchiever' = 'https://AccountAchiever.com';
    'AdministerArch' = 'https://AdministerArch.com';
    'AudioAmanda' = 'https://AudioAmanda.com';
    'AztecChicago' = 'https://AztecChicago.com';
    'BeTicket' = 'https://BeTicket.com';
    'BillyTorch' = 'https://BillyTorch.com';
    'BleachBand' = 'https://BleachBand.com';
    'BroadcastBulb' = 'https://BroadcastBulb.com';
    'BrushBeetle' = 'https://BrushBeetle.com';
    'ButtonBushes' = 'https://ButtonBushes.com';
    'CableCopy' = 'https://CableCopy.com';
    'CardCrowd' = 'https://CardCrowd.com';
    'CelticCairo' = 'https://CelticCairo.com';
    'CherriesCup' = 'https://CherriesCup.com';
    'ClaimChance' = 'https://ClaimChance.com';
    'ClammyEye' = 'https://ClammyEye.com';
    'CorrectChess' = 'https://CorrectChess.com';
    'CrimeChess' = 'https://CrimeChess.com';
    'DiamondDemo' = 'https://DiamondDemo.com';
    'DomainDance' = 'https://DomainDance.com';
    'DropThumb' = 'https://DropThumb.com';
    'EngineerIndustry' = 'https://EngineerIndustry.com';
    'FearInsurance' = 'https://FearInsurance.com';
    'GateSegment' = 'https://GateSegment.com';
    'InterestTeeth' = 'https://InterestTeeth.com';
    'JazzyMove' = 'https://JazzyMove.com';
    'JudgeJester' = 'https://JudgeJester.com';
    'LectureLinen' = 'https://LectureLinen.com';
    'LeftGuitar' = 'https://LeftGuitar.com';
    'LopezRocket' = 'https://LopezRocket.com';
    'MedusaMinimum' = 'https://MedusaMinimum.com';
    'MicroMachine' = 'https://MicroMachine.com';
    'ModifyGoose' = 'https://ModifyGoose.com';
    'ObtainableHammer' = 'https://ObtainableHammer.com';
    'PaperProgram' = 'https://PaperProgram.com';
    'PastelStreet' = 'https://PastelStreet.com';
    'PhotoPublic' = 'https://PhotoPublic.com';
    'PictureProcess' = 'https://PictureProcess.com';
    'PotatoPollution' = 'https://PotatoPollution.com';
    'PrivateCare' = 'https://PrivateCare.com';
    'QuirkyJudge' = 'https://QuirkyJudge.com';
    'ResearchQuill' = 'https://ResearchQuill.com';
    'RitzyPrice' = 'https://RitzyPrice.com';
    'RomanRadio' = 'https://RomanRadio.com';
    'ScareSilk' = 'https://ScareSilk.com';
    'ShipSquirrel' = 'https://ShipSquirrel.com';
    'ShockSidewalk' = 'https://ShockSidewalk.com';
    'ShoesSwim' = 'https://ShoesSwim.com';
    'SlaySquirrel' = 'https://SlaySquirrel.com';
    'SlideDonkey' = 'https://SlideDonkey.com';
    'SoloDerby' = 'https://SoloDerby.com';
    'SquashFaucet' = 'https://SquashFaucet.com';
    'StoneScissors' = 'https://StoneScissors.com';
    'TumbleTomatoes' = 'https://TumbleTomatoes.com';
    'TurtleEvident' = 'https://TurtleEvident.com';
    'UnderstoodHydrant' = 'https://UnderstoodHydrant.com';
    'WhimsicalWren' = 'https://WhimsicalWren.com';
    'WisdomExtra' = 'https://WisdomExtra.com';
    'WorryStop' = 'https://WorryStop.com';
    'ZippyGarden' = 'https://ZippyGarden.com';
    'Password Help' = 'https://www.Passwordhelp.com';
    'Totally Not AHacker' = 'https://www.TotallyNotAHacker.com';
    'Maybe A Hacker' = 'https://www.MaybeAHacker.com';
    'Likely A Hacker' = 'https://www.LikelyAHacker.com';
    'SuperChicken' = 'https://www.SuperChicken.com';
    'WonderChicken' = 'https://www.WonderChicken.com';
    'Ninja Jokes' = 'https://www.NinjaJokes.com';
    'TruckRobot' = 'https://www.TruckRobot.com';
    'RobotTruck' = 'https://www.RobotTruck.com';
    'RobotTruckRobot' = 'https://www.RobotTruckRobot.com';
    'CatfoodBreath' = 'https://www.CatfoodBreath.com';
    'Epsisode1wasTheBestone' = 'https://www.Epsisode1wasTheBestone.com';
    'JarJar is a Sith Lord' = 'https://jarjarsith.com';
    'Thanos Was Framed' = 'https://www.ThanosWasFramed.com';
    'Boondoggle' = 'https://www.boondoggle.com';
    'Google' = 'https://www.google.com';
    'Bing' = 'https://www.bing.com';
    'AOL' = 'https://www.aol.com';
    'Cofense' = 'https://www.cofense.com';
    'Yahoo' = 'https://www.yahoo.com';
    'GitHub' = 'https://www.github.com';
    'Gooogles' = 'https://www.gooogles.com';
    'Bling' = 'https://www.bling.com';
    'Ah-OL' = 'https://www.ah-ol.com';
    'Cofenseded' = 'https://www.cofenseded.com';
    'Yahooop' = 'https://www.yahooop.com';
    'GlitterHub' = 'https://www.glitterthub.com';
    'I am reading your emal' = 'https://shouldnothaveclicked.com'
}
### Adds URLs to email body using a selection of above URLs adjust the number of URLs below 
### 3+ URLs is recommended to keep Clusters accurate
$chosenDomains = $domains.GetEnumerator() | get-random -Count 3 ### This number determines the number of URLs in the email body
$strChosenDomains = $chosenDomains | ForEach-Object {
    return "<li><a href=`"$($_.Value)`">$($_.Name)</a></li>"
}

$strChosenDomains = $strChosenDomains -join "`n"

### URL based email
$emailBody = @"
<!DOCTYPE HTML>
<html>
<head>
<title>Test Email</title>
</head>
<body>
<h4>I assure you these are all trustworthy links...</h4>
<ul>
$strChosenDomains
</ul>
</body>
</html>
"@

$smtpServer = "cofensedemo-com.mail.protection.outlook.com"


$allUsers = Get-Content -Path "$PSScriptRoot\recipient\recipients.txt"

$allUsers | Where-Object {$_} | ForEach-Object {
    
    $to = $_
    $from = Generate-Sender
    $subject = Generate-Subject
    
    Write-Host "Sending email from $from to $to with subject `"$subject`""

    Send-MailMessage -From $from -To $to -Subject $subject -Body $emailBody -SmtpServer $smtpServer -BodyAsHtml -UseSsl

}