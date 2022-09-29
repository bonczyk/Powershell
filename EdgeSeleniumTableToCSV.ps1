$OutPath = ''
$sites = @('https://aib.ie/foreign-exchange-rates-sell',
           'https://aib.ie/foreign-exchange-rates-buy',
           'https://aib.ie/fxrates-calculator' )

$CurrentDir = if ($psise) { Split-Path $psise.CurrentFile.FullPath } else { $PSScriptRoot }
if (!$outPath) { $outPath = Split-Path $CurrentDir }
if (($env:Path -split ';') -notcontains $CurrentDir) { $env:Path += "$CurrentDir;" }
Add-Type -Path "$CurrentDir\WebDriver.dll"
$options = New-Object OpenQA.Selenium.Edge.EdgeOptions 
# $options.AddArguments("headless","log-level=3")
$edge = New-Object OpenQA.Selenium.Edge.EdgeDriver($CurrentDir,$options)
$edge.Manage().Window.Minimize(); "`n"*3
$sites[0..1] | % { $edge.Navigate().GoToURL($_);  sleep -Seconds 1
  $name = ($_ -replace 'https://aib.ie/')
  $edge.FindElements([OpenQA.Selenium.By]::XPath('//*[@id="foreignbuy"]/table[1]/tbody/tr')).GetAttribute('innerHTML').
  Trim() -replace '<[^>]+>',';' -replace ';;',',' -replace ';','' | Out-File "$outPath\$name.csv" -Encoding utf8 -Force
  "Saved file : $outPath\$name.csv"
}

$edge.Navigate().GoToURL($sites[2]); sleep -Seconds 1
1..2 | % {
  $all = $edge.FindElements([OpenQA.Selenium.By]::XPath('//*[@id="foreignbuy"]/div[1]/div['+$_+']')).text -split "`r`n" 
  $t = $all | select -skip 2
  (0..(($t.count/3)-1) | % { $t[$_*3] +','+ $t[$_*3+1] +','+ $t[$_*3+2] -replace "`r"} | Out-File "$outPath\$($all[0]).csv" -Encoding utf8 -Force )
  "Saved file : $outPath\$($all[0]).csv"
}

$edge.Quit()




<# Notes - other way to transform data 
#1..(($t.count/3)-1) | % { [PSCustomObject]@{ $t[0]=$t[$_*3]; $t[1]=$t[$_*3+1]; $t[2]=$t[$_*3+2]; }  } 
#$edge.FindElements([OpenQA.Selenium.By]::XPath('//*[@id="foreignbuy"]/div[1]/div[1]')).FindElements([OpenQA.Selenium.By]::ClassName("fx-rates-table__row")).Text | % { $_ -replace "`n",';' }
#$a = $edge.FindElements([OpenQA.Selenium.By]::ClassName('fx-rates-table__row')) | ? {$_.text} 
#>