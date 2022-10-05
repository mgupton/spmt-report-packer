#
# Created By: Michael Gupton
# Date Created: 2022-9-30
#
# Description:
#
# The script spmt-report-packer.ps1 is a quick-n-dirty script that packs
# up the report output from the Microsoft Sharepoint Migration Tool
# (SPMT). For every job SPMT creates a directory under
# %APPDATA%\Microsoft\MigrationTool\<domain name>\ with a unique
# auto-generated name. The name of the directory does not indicate the
# site the job output represents. When many sites are being migrated it
# can become inconvenient to have to drill into each directory and figure
# out which site it corresponds to. In addition, there could be previous
# jobs that were created but abandoned. The script spmt-report-packer.ps1
# is a quick-n-dirty script that lists all the job output directories and
# zips up the report subdirectory and gives it a name that identifies the
# site it corresponds to.

$jobs = @{}

ls -Directory | foreach {
    $rpt = "$_\Report"
    $summary = import-csv "$rpt\SummaryReport.csv"
    $id = split-path -leaf $_

    $summary | foreach {
        $lastrun = [datetime]::parseexact($_."Start time", 'M/d/yyyy h:m:s tt', $null)
        $sitename = split-path -leaf $_.Source

        $job = [PSCustomObject]@{
            id = $id
            sitename = $sitename
            src = $_.Source
            rpt = $rpt
            lastrun = $lastrun
        }

        if ($jobs.Keys -notcontains $job.src) {
            $jobs[$job.src] = $job

        }
        else {
            if ($job.lastrun -gt $jobs[$job.src].lastrun) {
                $jobs[$job.src] = $job
            }
        }
    }
}

 foreach ($job in $jobs.GetEnumerator()) {
    $job = $job | select -ExpandProperty value
    write-host "$env:TEMP\$($job.sitename)_$($job.id)"
    Compress-Archive -Force -Path $job.rpt -DestinationPath "$env:TEMP\$($job.sitename)_$($job.id)"
}
