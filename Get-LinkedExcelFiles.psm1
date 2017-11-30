Function Get-LinkedExcelFiles {
<#
.SYNOPSIS
Get linked Excel files under specified path.

.DESCRIPTION
Get linked Excel files under specified path.
Refer to https://blogs.technet.microsoft.com/heyscriptingguy/2010/04/05/hey-scripting-guy-how-can-i-search-a-microsoft-excel-workbook-for-links-to-other-workbooks/
Modified to run as a cmdlet with additional parameters, and using Excel Workbook.LinkSources() method to improve performance.

.PARAMETER Path
Specify the path to look for Excel files.

.PARAMETER Match
Specify the formula to match.

.PARAMETER Filter
Specify the file filter.

.PARAMETER Recurse
Specify whether to get files recursively.

.PARAMETER Depth
Specify the depth of recursion.

.OUTPUTS
Object

.EXAMPLE
PS> Get-LinkedExcelFiles -Path "C:\Temp"
Get Excel files containing external links recursively under path "C:\Temp".

.EXAMPLE
PS> Get-LinkedExcelFiles -Path "C:\Temp" | Export-Csv -Path C:\Temp\LinkedExcelFiles.csv -NoTypeInformation
Get Excel files containing external links recursively under path "C:\Temp" and export output to CSV file.

.EXAMPLE
PS> Get-LinkedExcelFiles -Path "C:\Temp" -WhatIf
Get Excel files recursively under path "C:\Temp" but do not open files.

.EXAMPLE
PS> Get-LinkedExcelFiles -Path "C:\Temp" -Depth 2
Get Excel files containing external links to a depth of two subdirectories under path "C:\Temp".

.EXAMPLE
PS> Get-LinkedExcelFiles -Path "C:\Temp" -Recurse:$false
Get Excel files containing external links non-recursively under path "C:\Temp".

.EXAMPLE
PS> Get-LinkedExcelFiles -Path "C:\Temp" -Match ".*\\\\fileserver1.*\[.*\].*|.*\\\\fileserver2.*\[.*\].*"
Get Excel files containing links to "\\fileserver1" or "\\fileserver2" under path "C:\Temp".

.EXAMPLE
PS> Get-LinkedExcelFiles -Path "S:\Department\Shared" -Match ".*\\\\fileserver1.*\[.*\].*|.*\\\\filserver2.*\[.*\].*" -Filter "*.xlsx"
Get Excel files containing links to "\\fileserver1" or "\\fileserver2" under path "S:\Department\Shared" using filter "*.xlsx".

.NOTES
  Version: 0.3.0 - Added progress bar and number of files for WhatIf
  Date: 2017-11-30

  Version: 0.2.1 - Added default value for Match parameter
  Date: 2017-11-29
  
  Version: 0.2.0 - Added -WhatIf functionality
  Date: 2017-11-27

  Version: 0.1.0 - Initial version
  Date: 2017-11-27
  
  Author: Richard Lock
#>

  [CmdletBinding(SupportsShouldProcess=$true)]

  Param (
    [Parameter(Mandatory = $true, ValueFromPipeline = $true)] [String]$Path,
    [Parameter()] [String]$Match = ".*\[.*\].*",
    [Parameter()] [String]$Filter = "*.xls",
    [Parameter()] [Boolean]$Recurse = $true,
    [Parameter(ParameterSetName = "Depth")] [Int32]$Depth
  )
  
  Begin {
    $gciParameters = @{ 
      Path = $Path
      Filter = $Filter
      Recurse = $Recurse
      Depth = $Depth
    }
    If ($PSCmdlet.ParameterSetName -ne "Depth" -or $Recurse -eq $false) {
      $gciParameters.Remove("Depth")
    }
    
    $excelWorkbooks = Get-Childitem @gciParameters
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $results = @()
  }
        
  Process {
    ForEach ($excelWorkbook In $excelWorkbooks) {
      If ($PSCmdlet.ShouldProcess($excelWorkbook.FullName, "Get Excel links")) {
        $i++
        
        $workbook = $excel.Workbooks.Open($excelWorkbook.FullName, 0, $true) # Open workbook, don't update links, read-only
        $linkSources = $workbook.LinkSources()
        ForEach ($linkSource in $linkSources) {
          $result = New-object PSObject
          $result | Add-Member -Name Workbook -Value $excelWorkbook.FullName -Membertype NoteProperty
          $result | Add-Member -Name Link -Value $linkSource -Membertype NoteProperty
          $results += $result
        }
                
        $workbook.Saved = $true
        $workbook.Close()
        
        Write-Progress -Activity "Searching $($excelWorkbooks.Count) Excel files under ""$Path"" for links" -PercentComplete (($i / $excelWorkbooks.Count) * 100)
      }
    }

    If (!$PSCmdlet.ShouldProcess("$($excelWorkbooks.Count) files", "Get Excel links")) {
    }

    $results
  }

  End {
    $excel.Quit()
    $excel = $null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
  }
}
