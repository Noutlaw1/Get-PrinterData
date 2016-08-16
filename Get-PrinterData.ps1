function Get-PrinterData($printer_address){

$printer = Invoke-WebRequest -Uri “$printer_address/info_suppliesStatus.html?tab=Home&menu=SupplyStatus”
if ($? -eq $false)
{
continue
}
$printer = $printer.ParsedHtml.IHTMLDocument3_documentElement.outerText
$printer = $printer.ToString() | Out-file “C:\users\nick outlaw\desktop\html.txt”

$printer = get-content “C:\users\nick outlaw\desktop\html.txt”
$colors = @()
$remaining = @()
$counter = 0

foreach ($line in $printer)
{
if ($line -match “Yellow Cartridge” -or $line -match “Cyan Cartridge” -or $line -match “Black Cartridge” -or $line -match “Magenta Cartridge”)
{
$colors += $line
}
if ($line -match “.[0-9]\d%” -or $line -match “..%”)
{
$counter = $counter + 1
if ($counter -lt 5)
{
$remaining += $matches[0]

}
}
}
write-host $colors
write-host $remaining
$counter = 0

Foreach ($color in $colors){
$data += @{$color = $remaining[$counter]}
$counter = $counter + 1
}

return $data
remove-file “C:\html.txt”
}

#Once the function runs, the rest of the script parses the data and puts it into a readable form. I used the Import-Excel module (https://github.com/dfinke/ImportExcel) to output the excel workbook. It took some workarounds to get it in the way I needed it, which is why it outputs several files over the execution of the script.

if (“C:\users\nick outlaw\desktop\html.txt”){ remove-item “C:\users\nick outlaw\desktop\html.txt”}
if (“C:\users\nick outlaw\desktop\printerink.csv”){remove-item “C:\users\nick outlaw\desktop\printerink.csv”}

#This is where the function gets the printers to retrieve data from, reading the ip addresses from a file and then looping with a foreach loop.

$printer_list = get-content “C:\printerlist.txt”
$collated_data = @{}
foreach ($line in $printer_list)
{

$printerstuff = Get-printerdata($line)
if ($printerstuff.count -eq 1)
{
$black = $printerstuff.’Black Cartridge ‘ | Out-string
$yellow = “N/A”
$cyan = “N/A”
$magenta = “N/A”
[pscustomobject]@{“IP” = $line; “Black ink” = $black; “Yellow” = $yellow; “Magenta” = $magenta; “Cyan” = $cyan} | Export-csv “C:\users\nick outlaw\desktop\printerink.csv” -append -UseCulture -NoTypeInformation

#Basically, it outputs all of the data to a CSV, then imports the CSV and converts it to an Excel workbook, before outputting it again. Not the most elegant solution, but it appears to work properly.

}
else
{
$yellow = $printerstuff.’Yellow Cartridge ‘| Out-string
$black = $printerstuff.’black Cartridge ‘| Out-string
$cyan = $printerstuff.’cyan Cartridge ‘| Out-string
$magenta = $printerstuff.’magenta Cartridge ‘| Out-string

[pscustomobject]@{“IP” = $line; “Black ink” = $black; “Yellow” = $yellow; “Magenta” = $magenta; “Cyan” = $cyan} | Export-csv “C:\users\nick outlaw\desktop\printerink.csv” -append -UseCulture -NoTypeInformation

}

}