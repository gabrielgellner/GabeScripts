## a wc like tool that works with word doc[x] files
function Measure-Doc {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true, ValueFromPipeLine=$true)]
        [string] $fileName # not sure what to do about sending arrays of filenames ...
    )

    Begin {
        # I get an error the first time I run the script that the assembly was not found/loaded ... 
        # need to figure out how to fix this
        # I need enum values from the office COM space (no longer true in powershell 5.0 it seems)
        $wdstats = [Microsoft.Office.Interop.Word.WdStatistic]
        
        $wordApp = New-Object -ComObject Word.Application
    }

    Process {
        $file = Convert-Path $fileName # how can I do this without the temp var $file?
        $doc = $wordApp.Documents.Open($file)
        
        # I might want to add the file size as well (call Length for PS like naming)
        $output = New-Object -TypeName psobject -Property (
            @{
                'Name'=$fileName;
                'Lines'=$doc.ComputeStatistics($wdstats::wdStatisticLines);
                'Words'=$doc.ComputeStatistics($wdstats::wdStatisticWords);
                'Characters'=$doc.ComputeStatistics($wdstats::wdStatisticCharacters);
                'Paragraphs'=$doc.ComputeStatistics($wdstats::wdStatisticParagraphs);
                'Pages'=$doc.ComputeStatistics($wdstats::wdStatisticPages);
            }
        )
        $doc.Close([ref]$false)

        Write-Output $output
    }

    End {
        $wordApp.Quit() # this doesn't seem to always work word still open after
    }
}

## du like command, might have to write this in C# to get acceptable speed
function Get-DiskUsage { # not sure if I should use Measure-<blah> in this case
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true, ValueFromPipeLine=$true)]
        [string] $folder # not sure what to do about sending arrays of filenames ...
    )

    Process {
        # might be good to see if use COM FileSystemObject is faster
        $fsize = (Get-ChildItem $folder -Recurse -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum)
        if ($fsize.Sum -eq $null) {$fsize = 0} else {$fsize = $fsize.Sum} # not sure if this slows it down a lot

        $out = New-Object -TypeName psobject -Property (
            @{
                'Name'=$folder;
                'Length'=$fsize;
            }
        )

        Write-Output $out
    }
}