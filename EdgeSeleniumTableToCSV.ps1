$OutPath = ''
$web = @('https://aib.ie/foreign-exchange-rates-sell','https://aib.ie/foreign-exchange-rates-buy',
         'https://aib.ie/fxrates-calculator', 'https://aibni.co.uk/fx-rates-calculator',
         'https://aibni.co.uk/business/ways-to-bank/ibusiness-banking/ibb-fx-buy','https://aibni.co.uk/business/ways-to-bank/ibusiness-banking/ibb-fx-sell',       
         'https://aibgb.co.uk/ways-to-bank/ibusiness-banking/ibb-exchange-rates/allied-irish-bank-gb-buy-rates','https://aibgb.co.uk/ways-to-bank/ibusiness-banking/ibb-exchange-rates/allied-irish-bank-gb-sell-rates')

$CurrentDir = if ($psise) { Split-Path $psise.CurrentFile.FullPath } else { $PSScriptRoot }
if (!$outPath) { $outPath = Split-Path $CurrentDir }
if (($env:Path -split ';') -notcontains $CurrentDir) { $env:Path += "$CurrentDir;" }
Add-Type -Path "$CurrentDir\WebDriver.dll"
$options = New-Object OpenQA.Selenium.Edge.EdgeOptions ; $options.AddArguments("log-level=3")
$edgeVer = ((Get-AppxPackage -Name *MicrosoftEdge.*).Version -split '\.')[0]
$edge = New-Object OpenQA.Selenium.Edge.EdgeDriver("$CurrentDir\$edgeVer",$options)
$edge.Manage().Window.Minimize(); "`n"*3

$work1 = { $name = split-path $_ -Leaf
  $edge.FindElements([OpenQA.Selenium.By]::XPath('//*[@id="foreignbuy"]/table[1]/tbody/tr')).GetAttribute('innerHTML').
  Trim() -replace '<[^>]+>',';' -replace ';;',',' -replace ';','' | Out-File "$outPath\$name.csv" -Encoding utf8 -Force
  "Saved file : $name.csv"
}

$work2 = { 1..2 | % {
  $all = $edge.FindElements([OpenQA.Selenium.By]::XPath('//*[@id="foreignbuy"]/div[1]/div['+$_+']')).text -split "`r`n" 
  $t = $all | select -skip 2
  (0..(($t.count/3)-1) | % { $t[$_*3] +','+ $t[$_*3+1] +','+ $t[$_*3+2] -replace "`r"} | Out-File "$outPath\$($all[0]).csv" -Encoding utf8 -Force )
  "Saved file : $($all[0]).csv" }
}

$work3  = { 1,3 | % {
  $name = (split-path $web[3] -Leaf)+'-'+$edge.FindElements([OpenQA.Selenium.By]::XPath('//*[@id="foreignbuy"]/div/div[1]/div['+$_+']')).Text -replace ' ','-'
  $edge.FindElements([OpenQA.Selenium.By]::XPath('//*[@id="foreignbuy"]/div/div[1]/div['+($_+1)+']/div/div/div')).GetAttribute('innerHTML').
  Trim() -replace '<[^>]+>',';' -replace ';;;;',',' -replace ';','' | Out-File "$outPath\$name.csv" -Encoding utf8 -Force
  "Saved file : $name.csv"}
}

$work4 = { $name = split-path $_ -Leaf
  $edge.FindElements([OpenQA.Selenium.By]::XPath('//*[@id="foreignbuy"]/table/tbody/tr')).GetAttribute('innerHTML').
  Trim() -replace '<[^>]+>',';' -replace ';;;',',' -replace ';','' | Out-File "$outPath\$name.csv" -Encoding utf8 -Force
  "Saved file : $name.csv"
}

$host.UI.RawUI.WindowSize = New-Object System.Management.Automation.Host.size(80,20); Clear-Host; "`n"*7
"Path - $outPath`n"

$web | % {
  $perc = [math]::Round($web.IndexOf($_)/$web.Count * 100, 1);
  Write-Progress "Processing $_" "Complete : $perc %" -perc $perc
  $edge.Navigate().GoToURL($_); Sleep -Seconds 1
  switch ($_) {
    {$_ -in $web[0..1]} {&$work1}
    $web[2]             {&$work2}
    $web[3]             {&$work3}
    {$_ -in $web[4..7]} {&$work4}
 }
}

$edge.Quit()
