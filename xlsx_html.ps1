Function ExcelToCsv ($File) {
    $Excel = New-Object -ComObject Excel.Application
    $Wb = $Excel.Workbooks.Open($File)
        
    $X = $File | Select-Object Directory, BaseName
    $N = [System.IO.Path]::Combine($X.Directory, (($X.BaseName, 'csv') -join "."))
    
    foreach ($Ws in $Wb.Worksheets) {
        $Ws.SaveAs($N, 6)
    }
    $Excel.Quit()
}

Function CsvToJson ($File) {
    $X = $File | Select-Object Directory, BaseName
    $N = [System.IO.Path]::Combine($X.Directory, (($X.BaseName, 'json') -join "."))

    $Header = 'A', 'B'

    $O = Get-Content $File
    $O = $O[1..($O.Count - 1)]
    $O | Out-File -FilePath .\temp.csv
    Import-Csv -Path .\temp.csv -Header $Header | ConvertTo-Json -Compress | Out-File $N
    del .\temp.csv
}

Function CsvToHtml ($File) {
    $Header = @"
<style>
TABLE {border-width: 1px; border-style: solid; border-color: black; border-collapse: collapse;}
TH {border-width: 1px; padding: 3px; border-style: solid; border-color: black; background-color: #6495ED;}
TD {border-width: 1px; padding: 3px; border-style: solid; border-color: black;}
</style>
"@

    $X = $File | Select-Object Directory, BaseName
    $N = [System.IO.Path]::Combine($X.Directory, (($X.BaseName, 'html') -join "."))

    Import-Csv $File | ConvertTo-Html -Head $Header | Out-File $N
}

########### Execute 

Get-ChildItem .\*.xlsx |
ForEach-Object {
    ExcelToCsv $_
}

Get-ChildItem .\*.csv |
ForEach-Object {
    CsvToJson $_
}
    
Get-ChildItem .\*.csv |
ForEach-Object {
    CsvToHtml $_
} 