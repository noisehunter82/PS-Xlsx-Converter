Function ExcelToCsv ($File) {
    $Excel = New-Object -ComObject Excel.Application
    $wb = $Excel.Workbooks.Open($File)
        
    $x = $File | Select-Object Directory, BaseName
    $n = [System.IO.Path]::Combine($x.Directory, (($x.BaseName, 'csv') -join "."))
    
    foreach ($ws in $wb.Worksheets) {
        $ws.SaveAs($n, 6)
    }
    $Excel.Quit()
}

Function CsvToHtml ($File) {
    $Header = @"
<style>
TABLE {border-width: 1px; border-style: solid; border-color: black; border-collapse: collapse;}
TH {border-width: 1px; padding: 3px; border-style: solid; border-color: black; background-color: #6495ED;}
TD {border-width: 1px; padding: 3px; border-style: solid; border-color: black;}
</style>
"@

    $x = $File | Select-Object Directory, BaseName
    $n = [System.IO.Path]::Combine($x.Directory, (($x.BaseName, 'html') -join "."))

    Import-Csv $File | ConvertTo-Html -Head $Header | Out-File $n
}
    
Get-ChildItem C:\Users\Jankowskim\Projects\PS-Scripts\*.xlsx |
ForEach-Object {
    ExcelToCsv -File $_
}
    
Get-ChildItem C:\Users\Jankowskim\Projects\PS-Scripts\*.csv |
ForEach-Object {
    CsvToHtml -File $_
} 