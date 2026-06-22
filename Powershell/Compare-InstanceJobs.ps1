#   This file is part of Compare-InstanceJobs.
#
#   This script is a derivative work based on "Compare-AGReplicaJobs":
#       Copyright 2020 Eitan Blumin <@EitanBlumin, https://www.eitanblumin.com>
#             while at Madeira Data Solutions <https://www.madeiradata.com>
#
#   Licensed under the MIT License (the "License");
#
#   Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
#
#   The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.
#
#   THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.


Function Compare-InstanceJobs
{
<#
.SYNOPSIS
Compare SQL Agent Jobs between any two SQL Server instances.
Compare-InstanceJobs Function: Compare-InstanceJobs
Based on "Compare-AGReplicaJobs" by Eitan Blumin (@EitanBlumin) | Madeira Data Solutions (@Madeira_Data)
License: MIT License

.DESCRIPTION
Compare-InstanceJobs compares SQL Agent Jobs between any two SQL Server instances, regardless of whether they
participate in an Availability Group or not. One instance acts as the "source" (the master / reference) and the
other as the "target" (the instance to be aligned to the source).

The cmdlet performs the following operations:
- Connects to the specified -SourceInstance and -TargetInstance using SMO.
- Collects the SQL Agent jobs from each instance, applying any of the optional filters (category, database
  context, enabled/disabled state, and "name contains").
- Compares the relevant jobs between the two instances.
- Generates an HTML report summarizing the differences.
- Generates a DROP/CREATE change-script to be applied to the target instance in order to "align" it to the source.
- Saves the report and change-script to an output folder, and (optionally) sends them by e-mail.

The four optional filters are all opt-in. When none are specified, ALL local SQL Agent jobs on both instances
are compared.

.PARAMETER SourceInstance
The "master" / reference SQL Server instance. Its jobs are treated as the desired state. The generated change
script aligns the target instance to match this one. Example: "SQLPROD01" or "SQLPROD01\INST2".

.PARAMETER TargetInstance
The SQL Server instance to compare against (and to be aligned to the source). Example: "SQLPROD02".

.PARAMETER JobCategories
Optional. One or more SQL Agent job categories to include. Jobs in any other category are ignored.
Leave empty to include jobs of all categories.

.PARAMETER DatabaseContext
Optional. One or more database names. Only jobs that contain at least one job step running against one of these
databases (the job step's DatabaseName, typically T-SQL steps) are included. Leave empty to ignore database context.

.PARAMETER JobState
Optional. Filter jobs by their enabled state. Valid values: "Enabled", "Disabled", "All". Default is "All".

.PARAMETER JobNameContains
Optional. Include only jobs whose name matches this pattern (PowerShell -like / LIKE semantics). If the value
contains no wildcard characters (* or ?), it is treated as a "contains" match (wrapped as *value*). If it does
contain wildcards, it is used verbatim. Leave empty to include jobs with any name.

.PARAMETER outputFolder
Optional. A folder path where the HTML report and the .sql change-script are saved. Leave empty to use the local
temporary folder.

.PARAMETER emailFrom
Optional. The e-mail address of the sender. Required only if you want an e-mail report.

.PARAMETER emailTo
Optional. One or more e-mail addresses for the recipients. Required only if you want an e-mail report.

.PARAMETER emailServerAddress
Optional. The SMTP server to use for sending the e-mail. Default is $PSEmailServer.

.PARAMETER emailServerPort
Optional. The port number for the SMTP server. Default is 25.

.PARAMETER emailCredential
Optional. A PSCredential object for authenticating against the SMTP server.

.PARAMETER emailUseSSL
Optional switch. Specify to use SSL/TLS for the SMTP connection.

.EXAMPLE
C:\PS> # Minimum parameters - compare ALL local jobs between two instances, output to temp folder:
C:\PS> Import-Module .\Compare-InstanceJobs.ps1; Compare-InstanceJobs -SourceInstance "SQLPROD01" -TargetInstance "SQLPROD02"

.EXAMPLE
C:\PS> # Compare only enabled jobs in specific categories, save to a folder and e-mail the report:
C:\PS> Compare-InstanceJobs -SourceInstance "SQLPROD01" -TargetInstance "SQLPROD02" `
C:\PS>     -JobCategories "Production Jobs","Maintenance" -JobState Enabled `
C:\PS>     -outputFolder "C:\Reports\" -From "no-reply@acme-corp.com" -To "dba@acme-corp.com" -EmailServer "smtp.acme-corp.com"

.EXAMPLE
C:\PS> # Compare only jobs whose name contains "Backup" and that touch a specific database:
C:\PS> Compare-InstanceJobs -SourceInstance "SQLPROD01" -TargetInstance "SQLPROD02" -JobNameContains "Backup" -DatabaseContext "SalesDB"

.EXAMPLE
C:\PS> # Using a PSCredential object for an authenticated SMTP server over SSL:
C:\PS> $username = "db_alerts@domain.com"
C:\PS> $password = ConvertTo-SecureString "mypassword" -AsPlainText -Force
C:\PS> $psCred   = New-Object System.Management.Automation.PSCredential -ArgumentList ($username, $password)
C:\PS> Compare-InstanceJobs -SourceInstance "A" -TargetInstance "B" -From "db_alerts@domain.com" -To "sysadmin@domain.com" -EmailServer "smtp.domain.com" -Port 587 -UseSsl -Credential $psCred

.EXAMPLE
C:\PS> # Display help:
C:\PS> Get-Help Compare-InstanceJobs -Full

.NOTES
This is a derivative of the open-source "Compare-AGReplicaJobs" project developed by Eitan Blumin while an
employee at Madeira Data Solutions, Madeira Ltd. The job-snapshot / change-script generation approach is reused;
the Availability Group auto-discovery has been replaced with an explicit two-instance comparison plus optional
job filters.

.LINK
https://madeiradata.github.io/mssql-jobs-hadr

.LINK
https://eitanblumin.com
#>
[CmdletBinding()]
Param(

    [Parameter(Mandatory=$true, Position=0,
    HelpMessage="Enter the source (master/reference) SQL Server instance whose jobs represent the desired state.")]
    [Alias("Source","Reference","PrimaryInstance","SourceServer")]
    [ValidateNotNullOrEmpty()]
    [String]
    $SourceInstance,

    [Parameter(Mandatory=$true, Position=1,
    HelpMessage="Enter the target SQL Server instance to compare against (and to be aligned to the source).")]
    [Alias("Target","Difference","SecondaryInstance","TargetServer")]
    [ValidateNotNullOrEmpty()]
    [String]
    $TargetInstance,

    [Parameter(Mandatory=$false,
    HelpMessage="Enter a list of one or more job categories to include. Leave empty to include all categories.")]
    [Alias("Categories")]
    [AllowEmptyCollection()]
    [String[]]
    $JobCategories = @(),

    [Parameter(Mandatory=$false,
    HelpMessage="Enter a list of one or more database names. Only jobs with a step targeting one of these databases are included. Leave empty to ignore database context.")]
    [Alias("Databases","DatabaseNames")]
    [AllowEmptyCollection()]
    [String[]]
    $DatabaseContext = @(),

    [Parameter(Mandatory=$false,
    HelpMessage="Filter jobs by enabled state: Enabled, Disabled, or All. Default is All.")]
    [Alias("State","Status")]
    [ValidateSet("Enabled","Disabled","All")]
    [String]
    $JobState = "All",

    [Parameter(Mandatory=$false,
    HelpMessage="Include only jobs whose name matches this pattern (LIKE / -like semantics). A value without wildcards is treated as a 'contains' match.")]
    [Alias("NameContains","NameLike")]
    [String]
    $JobNameContains = "",

    [Parameter(Mandatory=$false,
    HelpMessage="Enter a folder path where to save the report and change-script. Leave empty to use the local temporary folder.")]
    [String]
    $outputFolder = "",

    [Parameter(Mandatory=$false,
    HelpMessage="Enter the e-mail address of the sender. Required only for an e-mail report.")]
    [Alias("From","Sender","EmailSender")]
    [String]
    $emailFrom = "",

    [Parameter(Mandatory=$false,
    HelpMessage="Enter a list of one or more e-mail addresses for the recipients. Required only for an e-mail report.")]
    [Alias("To","Recipients","EmailRecipients")]
    [String[]]
    $emailTo = @(),

    [Parameter(Mandatory=$false,
    HelpMessage='Enter an address for the SMTP server to use for sending the e-mail. Default is $PSEmailServer.')]
    [Alias("EmailServer","SMTPServer","SMTP")]
    [String]
    $emailServerAddress = $PSEmailServer,

    [Parameter(Mandatory=$false,
    HelpMessage="Enter a port number for the e-mail server. Default is 25.")]
    [Alias("Port","EmailPort","SMTPPort")]
    [Int32]
    $emailServerPort = 25,

    [Parameter(Mandatory=$false,
    HelpMessage="Enter a credential object for the e-mail server.")]
    [Alias("Credential","EmailCredentials")]
    [AllowNull()]
    [PSCredential]
    $emailCredential = $null,

    [Parameter(Mandatory=$false,
    HelpMessage="Specify whether to use SSL for the e-mail server.")]
    [Alias("UseSsl")]
    [Switch]
    $emailUseSSL
)
Begin
{
    # Load SMO (same approach as the original script).
    $asm = [System.Reflection.Assembly]::LoadWithPartialName('Microsoft.SqlServer.SMO')

    # Scripting options used to generate the DROP part of the change-script.
    $ScriptOptionsForDrop = New-Object Microsoft.SqlServer.Management.Smo.ScriptingOptions
    $ScriptOptionsForDrop.IncludeIfNotExists = $true
    $ScriptOptionsForDrop.ScriptDrops = $true

    # The set of (scalar) properties used by Compare-Object to decide whether two jobs of the same name differ.
    # NOTE: Job steps and schedules are compared via deterministic "signature" strings (see Get-CompareableJob)
    #       rather than as raw collection objects, because Compare-Object cannot meaningfully diff array properties.
    $JobPropertiesToCompare = @(
        "Name","Category","IsEnabled","DeleteLevel","Description","EmailLevel","EventLogLevel","NetSendLevel",
        "OperatorToEmail","OperatorToNetSend","OperatorToPage","StartStepID","OwnerLoginName",
        "StepsSignature","SchedulesSignature"
    )

    # The per-step properties that participate in the comparison. This single list is used BOTH to build the
    # step signature (in Get-CompareableJob) and to report the exact differing step properties (in
    # Get-JobStepDifferences), so the two can never drift apart. Steps are matched between instances by NAME
    # (step names are unique within a job), and the step's ID (its execution position) is itself one of the
    # compared properties - so the same-named step landing at a different position is reported as an "ID" change.
    $JobStepCompareProperties = @(
        "ID","SubSystem","DatabaseName","DatabaseUserName","OnSuccessAction","OnFailAction","Command"
    )

    # Normalize the "name contains" pattern into a -like pattern.
    # If the caller supplied wildcards (* or ?), respect them; otherwise treat the value as a substring "contains".
    $NamePattern = $JobNameContains
    if (-not [string]::IsNullOrEmpty($NamePattern) -and $NamePattern -notmatch '[\*\?]') {
        $NamePattern = "*$NamePattern*"
    }

    # Build a normalized snapshot of a single SMO Job, including deterministic signatures of its steps and
    # schedules, plus portable DROP and CREATE change-scripts.
    function Get-CompareableJob {
    param([object]$JobObject)

        # --- Job steps snapshot ---
        $JobSteps = @()
        $JobObject.JobSteps | ForEach-Object {
            $currStep = @{
                Name                        = $_.Name;
                Command                     = $_.Command;
                CommandExecutionSuccessCode = $_.CommandExecutionSuccessCode;
                DatabaseName                = $_.DatabaseName;
                DatabaseUserName            = $_.DatabaseUserName;
                ID                          = $_.ID;
                JobStepFlags                = $_.JobStepFlags;
                OnFailAction                = $_.OnFailAction;
                OnFailStep                  = $_.OnFailStep;
                OnSuccessAction             = $_.OnSuccessAction;
                OSRunPriority               = $_.OSRunPriority;
                OutputFileName              = $_.OutputFileName;
                ProxyName                   = $_.ProxyName;
                RetryAttempts               = $_.RetryAttempts;
                RetryInterval               = $_.RetryInterval;
                Server                      = $_.Server;
                SubSystem                   = $_.SubSystem
            }
            $JobSteps += [pscustomobject]$currStep
        }

        # --- Job schedules snapshot ---
        $JobSchedules = @()
        $JobObject.JobSchedules | ForEach-Object {
            $currSchedule = @{
                Name                       = $_.Name;
                ActiveEndDate              = $_.ActiveEndDate;
                ActiveEndTimeOfDay         = $_.ActiveEndTimeOfDay;
                ActiveStartDate            = $_.ActiveStartDate;
                ActiveStartTimeOfDay       = $_.ActiveStartTimeOfDay;
                FrequencyInterval          = $_.FrequencyInterval;
                FrequencyRecurrenceFactor  = $_.FrequencyRecurrenceFactor;
                FrequencyRelativeIntervals = $_.FrequencyRelativeIntervals;
                FrequencySubDayInterval    = $_.FrequencySubDayInterval;
                FrequencySubDayTypes       = $_.FrequencySubDayTypes;
                IsEnabled                  = $_.IsEnabled
            }
            $JobSchedules += [pscustomobject]$currSchedule
        }

        # --- Deterministic signatures (ordered) so Compare-Object can detect step/schedule changes ---
        # Steps are keyed by NAME (sorted), so a step that merely moved position surfaces as an ID change
        # (ID is part of $JobStepCompareProperties) rather than as a wholesale add/remove.
        $StepsSignature = (
            $JobSteps | Sort-Object Name | ForEach-Object {
                $step = $_
                "Step '{0}'|{1}" -f $step.Name, (($JobStepCompareProperties | ForEach-Object { "$_=$($step.$_)" }) -join "|")
            }
        ) -join [Environment]::NewLine

        $SchedulesSignature = (
            $JobSchedules | Sort-Object Name | ForEach-Object {
                "Sched|{0}|Enabled={1}|FreqInt={2}|RecFactor={3}|RelInt={4}|SubDayType={5}|SubDayInt={6}|StartDate={7}|StartTime={8}" -f `
                    $_.Name, $_.IsEnabled, $_.FrequencyInterval, $_.FrequencyRecurrenceFactor, `
                    $_.FrequencyRelativeIntervals, $_.FrequencySubDayTypes, $_.FrequencySubDayInterval, `
                    $_.ActiveStartDate, $_.ActiveStartTimeOfDay
            }
        ) -join [Environment]::NewLine

        [pscustomobject]@{
            Name               = $JobObject.Name;
            Category           = $JobObject.Category;
            IsEnabled          = $JobObject.IsEnabled;
            DeleteLevel        = $JobObject.DeleteLevel;
            Description        = $JobObject.Description;
            EmailLevel         = $JobObject.EmailLevel;
            EventLogLevel      = $JobObject.EventLogLevel;
            NetSendLevel       = $JobObject.NetSendLevel;
            JobSteps           = $JobSteps;
            JobSchedules       = $JobSchedules;
            StepsSignature     = $StepsSignature;
            SchedulesSignature = $SchedulesSignature;
            OperatorToEmail    = $JobObject.OperatorToEmail;
            OperatorToNetSend  = $JobObject.OperatorToNetSend;
            OperatorToPage     = $JobObject.OperatorToPage;
            StartStepID        = $JobObject.StartStepID;
            OwnerLoginName     = $JobObject.OwnerLoginName;
            JobID              = $JobObject.JobID;
            # DROP by job NAME (not id) so the script is portable between instances.
            # Job.Script() returns a StringCollection (one batch per element); join with newlines for clean SQL.
            DropScript         = (@($JobObject.Script($ScriptOptionsForDrop)) | ForEach-Object {
                                      $_.Replace("@job_id=N'$($JobObject.JobID)'","@job_name=N'$($JobObject.Name)'")
                                  }) -join [Environment]::NewLine;
            CreateScript       = (@($JobObject.Script()) -join [Environment]::NewLine)
        }
    }

    # Collect and filter the jobs of a single instance, returning normalized "compareable" objects.
    function Get-FilteredInstanceJobs {
    param([string]$InstanceName)

        Write-Verbose "Connecting to instance [$InstanceName] ..."
        $smo = New-Object Microsoft.SqlServer.Management.SMO.Server($InstanceName)

        # Touch a property to force a connection now, so failures surface with a clear message.
        try {
            $null = $smo.JobServer.Jobs.Count
        }
        catch {
            throw "Failed to connect to or enumerate jobs on SQL Server instance [$InstanceName]: $($_.Exception.Message)"
        }

        $collected = @()

        $smo.JobServer.Jobs | Where-Object {
                # Only local jobs (exclude multi-server / MSX-managed jobs), matching the original script.
                $_.JobType -eq "Local"
            } | Where-Object {
                # 1) Category filter
                ($JobCategories.Count -eq 0) -or ($JobCategories -contains $_.Category)
            } | Where-Object {
                # 2) "Name contains" (LIKE) filter
                [string]::IsNullOrEmpty($NamePattern) -or ($_.Name -like $NamePattern)
            } | Where-Object {
                # 3) Enabled / Disabled / All filter
                ($JobState -eq "All") -or `
                ($JobState -eq "Enabled"  -and $_.IsEnabled) -or `
                ($JobState -eq "Disabled" -and -not $_.IsEnabled)
            } | ForEach-Object {
                $job = $_
                # 4) Database context filter: include only if at least one step targets one of the named databases.
                $passesDbFilter = ($DatabaseContext.Count -eq 0) -or `
                                  (@($job.JobSteps | Where-Object { $DatabaseContext -contains $_.DatabaseName }).Count -gt 0)

                if ($passesDbFilter) {
                    $collected += Get-CompareableJob $job
                }
            }

        Write-Verbose "Found $($collected.Count) matching job(s) on [$InstanceName]."
        # Comma forces the array to be returned as a single object (prevents unrolling of 0/1-element arrays).
        return ,$collected
    }

    # Given the source and target snapshots of a job that exists on both sides, return a list of segments
    # describing what differs. Scalar properties appear as their name; the JobSteps difference is expanded to
    # show which steps and which step properties differ. Segments are joined by the caller with "; ".
    function Get-DifferingProperties {
    param([object]$SourceJob, [object]$TargetJob)

        $segments = @()

        foreach ($prop in $JobPropertiesToCompare) {
            if ($prop -eq "Name") { continue }   # Name is the key, not a difference

            $sourceValue = if ($SourceJob -ne $null) { $SourceJob.$prop } else { $null }
            $targetValue = if ($TargetJob -ne $null) { $TargetJob.$prop } else { $null }

            # Compare as strings so enums/bools/ints/dates compare by their stable text representation.
            if (([string]$sourceValue) -ne ([string]$targetValue)) {
                switch ($prop) {
                    "StepsSignature" {
                        $stepMessages = Get-JobStepDifferences -SourceJob $SourceJob -TargetJob $TargetJob
                        if ($stepMessages.Count -gt 0) {
                            $segments += "JobSteps {" + ($stepMessages -join " | ") + "}"
                        } else {
                            $segments += "JobSteps"
                        }
                    }
                    "SchedulesSignature" { $segments += "JobSchedules" }
                    default              { $segments += $prop }
                }
            }
        }

        return $segments
    }

    # Compare the job steps of two job snapshots (matched by step NAME) and return human-readable messages
    # describing, per step, which step properties differ (or whether the step exists on only one side).
    # The step's ID (execution position) is one of the compared properties, so a moved step shows up as an ID diff.
    function Get-JobStepDifferences {
    param([object]$SourceJob, [object]$TargetJob)

        $messages    = @()
        $sourceSteps = @($SourceJob.JobSteps)
        $targetSteps = @($TargetJob.JobSteps)
        $allStepNames = @($sourceSteps | ForEach-Object { $_.Name }) + @($targetSteps | ForEach-Object { $_.Name }) | Sort-Object -Unique

        foreach ($name in $allStepNames) {
            $s = $sourceSteps | Where-Object { $_.Name -eq $name } | Select-Object -First 1
            $t = $targetSteps | Where-Object { $_.Name -eq $name } | Select-Object -First 1

            if ($s -eq $null) {
                $messages += "Step '$name' [only on target]"
            }
            elseif ($t -eq $null) {
                $messages += "Step '$name' [only on source]"
            }
            else {
                $stepDiffs = @()
                foreach ($p in $JobStepCompareProperties) {
                    if (([string]$s.$p) -ne ([string]$t.$p)) { $stepDiffs += $p }
                }
                if ($stepDiffs.Count -gt 0) {
                    $messages += "Step '$name' [" + ($stepDiffs -join ", ") + "]"
                }
            }
        }

        return $messages
    }
}
Process
{
    if ($SourceInstance -eq $TargetInstance) {
        Write-Warning "SourceInstance and TargetInstance are identical ('$SourceInstance'). The comparison will show no differences."
    }

    # Gather jobs from both instances.
    $SourceJobs = Get-FilteredInstanceJobs -InstanceName $SourceInstance
    $TargetJobs = Get-FilteredInstanceJobs -InstanceName $TargetInstance

    Write-Verbose "=== Comparing $($SourceJobs.Count) source job(s) against $($TargetJobs.Count) target job(s) ==="

    # Compare. Each differing job yields a '<=' row (present/different on source) and/or a '=>' row (present/different on target).
    $comparedJobs = Compare-Object -ReferenceObject $SourceJobs -DifferenceObject $TargetJobs -Property $JobPropertiesToCompare

    # ----- Build the HTML report body -----
    $sourceKey = $SourceInstance.Replace("\","_")
    $targetKey = $TargetInstance.Replace("\","_")
    $scriptCid = "changescript_$targetKey"

    $title    = "SQL Agent Job Comparison"
    $subtitle = "Source (master): <b>$SourceInstance</b> &nbsp;|&nbsp; Target (to align): <b>$TargetInstance</b>"

    # Describe which optional filters were applied, for transparency in the report.
    $filterParts = @()
    if ($JobCategories.Count -gt 0)            { $filterParts += "Categories: " + ($JobCategories -join ", ") }
    if ($DatabaseContext.Count -gt 0)          { $filterParts += "Database context: " + ($DatabaseContext -join ", ") }
    if ($JobState -ne "All")                   { $filterParts += "State: $JobState" }
    if (-not [string]::IsNullOrEmpty($JobNameContains)) { $filterParts += "Name like: $NamePattern" }
    $filterText = if ($filterParts.Count -gt 0) { "Filters &rarr; " + ($filterParts -join " &nbsp;|&nbsp; ") } else { "Filters &rarr; none (all jobs compared)" }

    # Columns shown in the comparison table. The Status column position is derived from this same list so the
    # "nowrap" CSS rule (in the document head below) stays correct even if the column order changes.
    $reportColumns  = @('Job','Source','Status','Target','Different Properties')
    $statusColIndex = [Array]::IndexOf($reportColumns, 'Status') + 1   # CSS :nth-child() is 1-based

    $htmlBody = ""

    if ($comparedJobs -eq $null -or @($comparedJobs).Count -eq 0)
    {
        $htmlBody = "<h1>$title</h1><p>$subtitle</p><p>$filterText</p><p><b>No differences found.</b> The two instances are in sync for the selected jobs.</p>"
    }
    else
    {
        # Group by job name to derive a single status (and the differing properties) per job.
        # NOTE: angle brackets below are written as plain '<' / '>' on purpose - ConvertTo-Html
        #       HTML-encodes table cell values for us, so pre-encoding here would double-encode.
        $reportRows = $comparedJobs | Group-Object Name | ForEach-Object {
            $name      = $_.Name
            $sides     = @($_.Group | ForEach-Object { $_.SideIndicator } | Sort-Object -Unique)
            $sourceJob = $SourceJobs | Where-Object { $_.Name -eq $name } | Select-Object -First 1
            $targetJob = $TargetJobs | Where-Object { $_.Name -eq $name } | Select-Object -First 1

            if ($sides -contains "<=" -and $sides -contains "=>") {
                $status    = "<>"
                $diffProps = (Get-DifferingProperties -SourceJob $sourceJob -TargetJob $targetJob) -join "; "
            }
            elseif ($sides -contains "<=") {
                $status    = ">="
                $diffProps = "(job exists only on source)"
            }
            else {
                $status    = "<="
                $diffProps = "(job exists only on target)"
            }

            [pscustomobject]@{
                Job                    = $name
                Source                 = $SourceInstance
                Status                 = $status
                Target                 = $TargetInstance
                'Different Properties' = $diffProps
            }
        }

        $htmlBody += $reportRows `
            | ConvertTo-Html -Property $reportColumns -Fragment `
                -PreContent "<h1>$title</h1><p>$subtitle</p><p>$filterText</p>" `
                -PostContent "<h3>Change Script</h3><ul><li><a href='cid:{$scriptCid}'>{$scriptCid}</a></li></ul><p><i>The change script aligns the target instance to the source: it DROPs jobs that are missing-from-source or different on the target, then (re)CREATEs the source versions.</i></p>"
    }

    # ----- Build the DROP/CREATE change-script for the target instance -----
    $changeScript = ""
    $newLineGo = [Environment]::NewLine + "GO" + [Environment]::NewLine

    if ($comparedJobs -ne $null)
    {
        # DROP on target: jobs flagged as '=>' (extra on target, or the target side of a "different" job).
        $droppedNames = @()
        $comparedJobs | Where-Object { $_.SideIndicator -eq "=>" } | ForEach-Object {
            $name = $_.Name
            if ($droppedNames -notcontains $name) {
                $droppedNames += $name
                $targetJob = $TargetJobs | Where-Object { $_.Name -eq $name } | Select-Object -First 1
                if ($targetJob -ne $null) {
                    $changeScript += "-- DROP (exists on target, missing-or-different on source): [$name]" + [Environment]::NewLine
                    $changeScript += $targetJob.DropScript
                    $changeScript += $newLineGo
                }
            }
        }

        # CREATE on target: jobs flagged as '<=' (present on source, missing-or-different on target).
        $createdNames = @()
        $comparedJobs | Where-Object { $_.SideIndicator -eq "<=" } | ForEach-Object {
            $name = $_.Name
            if ($createdNames -notcontains $name) {
                $createdNames += $name
                $sourceJob = $SourceJobs | Where-Object { $_.Name -eq $name } | Select-Object -First 1
                if ($sourceJob -ne $null) {
                    $changeScript += "-- CREATE (align target to source): [$name]" + [Environment]::NewLine
                    $changeScript += $sourceJob.CreateScript
                    $changeScript += $newLineGo
                }
            }
        }
    }

    $hasDifferences = -not [string]::IsNullOrWhiteSpace($changeScript)

    # ----- Resolve output folder -----
    if ([string]::IsNullOrEmpty($outputFolder) -or -not (Test-Path $outputFolder)) {
        if (-not [string]::IsNullOrEmpty($outputFolder)) {
            Write-Warning "Output folder '$outputFolder' not found. Falling back to the temporary folder."
        }
        $outputFolder = [System.IO.Path]::GetTempPath()
    }

    $dateStamp    = Get-Date -Format "yyyyMMdd"
    $sqlFileName  = "$($targetKey)_align_jobs_$dateStamp.sql"
    $sqlFilePath  = Join-Path $outputFolder $sqlFileName
    $htmlFilePath = Join-Path $outputFolder "jobs_comparison_report_$dateStamp.html"

    $filesToAttach = @()

    # Write the change-script only when there is something to apply.
    if ($hasDifferences) {
        $scriptHeader  = "-- Change script to align [$TargetInstance] to [$SourceInstance]" + [Environment]::NewLine
        $scriptHeader += "-- Generated on $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")" + [Environment]::NewLine + [Environment]::NewLine
        ($scriptHeader + $changeScript) | Out-File -FilePath $sqlFilePath -Encoding UTF8
        $filesToAttach += $sqlFilePath
        Write-Host "Output change-script: $sqlFilePath"

        # Wire the report's change-script link to the actual file name.
        $htmlBody = $htmlBody.Replace("{$scriptCid}", $sqlFileName)
    } else {
        # Remove the change-script placeholder/section when there is nothing to apply.
        $htmlBody = $htmlBody.Replace("{$scriptCid}", "(none)")
    }

    # Always write the HTML report so the caller gets output in their folder regardless of result.
    # The document head sets basic table styling and forces the Status column to never wrap.
    $reportHead = @"
<title>$title</title>
<style type="text/css">
    body   { font-family: Segoe UI, Tahoma, Arial, sans-serif; font-size: 13px; }
    table  { border-collapse: collapse; margin-top: 8px; }
    th, td { border: 1px solid #cccccc; padding: 4px 8px; text-align: left; vertical-align: top; }
    th     { background-color: #f2f2f2; }
    /* Status column (column $statusColIndex) should never wrap */
    td:nth-child($statusColIndex), th:nth-child($statusColIndex) { white-space: nowrap; text-align: center; }
</style>
"@

    $reportBody = ConvertTo-Html -Body $htmlBody -Head $reportHead -PostContent "<p>Generated on $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")</p>"
    $reportBody | Out-File -FilePath $htmlFilePath -Encoding UTF8
    $filesToAttach = @($htmlFilePath) + $filesToAttach

    if ($hasDifferences) {
        Write-Host "Discrepancies found. Report saved to: $htmlFilePath"
    } else {
        Write-Host "No discrepancies found. Report saved to: $htmlFilePath"
    }

    # ----- Optionally send the e-mail report -----
    $emailRequested = ($emailTo.Count -gt 0) -and (-not [string]::IsNullOrEmpty($emailFrom)) -and (-not [string]::IsNullOrEmpty($emailServerAddress))

    if ($emailRequested) {
        Write-Verbose "Sending e-mail report ..."

        $subjectState = if ($hasDifferences) { "Discrepancies found" } else { "In sync" }

        $mailParams = @{
            From        = $emailFrom
            To          = $emailTo
            Subject     = "SQL Agent Job Comparison ($subjectState): $SourceInstance vs $TargetInstance"
            Body        = ($reportBody -join [Environment]::NewLine)
            BodyAsHtml  = $true
            Attachments = $filesToAttach
            SmtpServer  = $emailServerAddress
            Port        = $emailServerPort
        }

        if ($emailUseSSL) {
            $mailParams["UseSsl"] = $true
        }
        if ($emailCredential -ne $null) {
            $mailParams["Credential"] = $emailCredential
        }

        Send-MailMessage @mailParams
        Write-Host "E-mail report sent to: $($emailTo -join ', ')"
    }
    elseif ($emailTo.Count -gt 0) {
        Write-Warning "E-mail recipients were specified, but -emailFrom and/or an SMTP server are missing. No e-mail was sent."
    }

    if (-not $hasDifferences) {
        Write-Host "OK"
    }
}
}