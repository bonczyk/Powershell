$sites = @('https://aib.ie/foreign-exchange-rates-sell','https://aib.ie/foreign-exchange-rates-buy')

$drvPath = "C:\windows"
$options = New-Object OpenQA.Selenium.Edge.EdgeOptions 
$options.AddArguments("headless","log-level=3")
$edge = New-Object OpenQA.Selenium.Edge.EdgeDriver("C:\windows",$options)
$sites | % { $edge.Navigate().GoToURL($_) 
  sleep -Seconds 2
  $edge.FindElements([OpenQA.Selenium.By]::XPath('//*[@id="foreignbuy"]/table[1]/tbody/tr')).GetAttribute('innerHTML').
  Trim() -replace '<[^>]+>',';' -replace ';;',',' -replace ';','' | ConvertFrom-Csv -Delimiter ',' | ft
  "Saved file : " + $_ -replace 'https://aib.ie/'
}
$edge.Quit()