
#region <functions>

<#
    .SYNOPSIS
        This function creates provides a simple way to log messages.

    .DESCRIPTION
        The function provides the felexability of printing to screen only or printing to screen and log file.
        You can set the Severity of the message and add colors when printing to screen.

    .PARAMETER Severity
        [string] The logged message sevirity

    .PARAMETER Text
        [string] The text being loged

    .PARAMETER ForegroundColor
        [System.ConsoleColor] The color of the message whene printing to screen. Default color is White.

    .PARAMETER LogFileName
        [string] The file name including the path where the log file will be located


    .EXAMPLE            
        # Print to screen only.
        Write-Text -Text "Hellow World!"

        # Print to screen and log to file.
        Write-Text -Text "Hellow World!" -LogFileName "C:\Scrpits\PowerShellApp.log";

        # Print to screen and log to file while specifying a Sevirity and  Color.
        Write-Text -Severity Info  -Text "Hellow World!" -ForegroundColor Green -LogFileName "C:\Scrpits\PowerShellApp.log";
        Write-Text -Severity Error -Text "Hellow World!" -ForegroundColor Red   -LogFileName "C:\Scrpits\PowerShellApp.log";

    .NOTES
        Author:  Yaniv Etrogi
        Website: https://sqlserverutilities.com
        Email:   yaniv.etrogi@gmail.com   
#>
function Write-Text
{
    [CmdletBinding()]  
    Param
    (   
        [Parameter(Mandatory=$false, ValueFromPipelineByPropertyName=$false)] [ValidateSet('Info','Warn','Error')] 
        [string]$Severity = 'Info',

        [Parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$false)] [ValidateNotNullOrEmpty()] 
        [string]$Text,

        [Parameter(Mandatory=$false, ValueFromPipelineByPropertyName=$false)]
        [System.ConsoleColor]$ForegroundColor = [System.ConsoleColor]::White,  

        [Parameter(Mandatory=$false, ValueFromPipelineByPropertyName=$false)]
        [string]$LogFileName        
    )
    Begin
    {
        # Validate
        # None user session must spply a log file.
        if(-not [Environment]::UserInteractive -and [string]::IsNullOrEmpty($LogFileName))
        {
            Throw "Please specify a value for $LogFileName. Terminating.";
        }
    
        try
        {
            [string]$dt = (Get-Date).ToString('yyyy-MM-ddTHH:mm:ss');
            [string]$message = '{0}  {1}  {2}' -f $dt, $Severity, $Text; 
        
            # User session we print.
            if ([Environment]::UserInteractive) {Write-Host -ForegroundColor $ForegroundColor $message; };   

        
            if (-not [string]::IsNullOrEmpty($LogFileName)) 
            {                                
                Add-Content -Path $LogFileName -Value $message;
            }        
        }
        catch
        {       
            throw $_;
        }
    }
}


<#
    .SYNOPSIS
        This function creates an SQLCredentials object to be used for connectiong to SQL Server using SQL Server authentication mode.

    .DESCRIPTION
        The function usese a saved encrypted password file that needs to be generated one time only.
        Once the file exists the function reads it from disk eliminating the need for a user input allowing scripts to be automated.

        Note that by when using ConvertFrom-SecureString and ConvertTo-SecureString while not specifying the -Key option powershell will use by default "Windows Data Protection API" (DPAPI).
        This is great but limmits your password file to be valid only on the machine weher it was created.
        If you need your script to work across multipple servers then it is simpler to use the -Key option will instruct powershell to use the "Advanced Encryption Standard" (AES) instead of DPAPI.
        Doing so results in a password file that will work accross multiple machines allowing you to create the password file one time and then just copy it along the script to other machines. 
    
    .PARAMETER FileFullName
        [string] The file name including the path where to store the encrypted password file. i.e.: "C:\Scrpits\sql_password.txt"

    .PARAMETER UserName
        [string] The The sql user for which the password is created.


    .EXAMPLE            
        $SqlCredential = Get-SQLCredentials -FileFullName 'C:\Scrpits\sql_password.txt' -UserName 'CloudMonitoring';

    .NOTES
        Author:  Yaniv Etrogi
        Website: https://sqlserverutilities.com/
        Email:   yaniv.etrogi@gmail.com 
#>
function Get-SQLCredentials 
{
    [OutputType([System.Data.SqlClient.SqlCredential])]
    param 
    (        

        [Parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$false)]
        [string]$FileFullName,

        [Parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$false)]
        [string]$UserName
    )
        
    try
    {
        [string]$key_file_full_name = Join-Path $PSScriptRoot 'key.txt';    

        # Key
        if (-not ($key))
        {
            # If the key file does not exists create it.
            if (-not (Test-Path -Path $key_file_full_name -PathType Leaf))
            {
                [System.Array]$key = [System.Text.Encoding]::UTF8.GetBytes("A-32-byte-key-AES-256-encryption");
                Set-Content -Path $key_file_full_name -Value $key;            
            }
            else
            {
                # Load the key file.
                [System.Array]$key = Get-Content $key_file_full_name; 
            }
        }

        # User Input: If the file does not exists create the credentials and save to file.
        if (-not (Test-Path -Path $FileFullName -PathType Leaf) )
        {   
            Write-Host "SQL Credentials not found. Please enter your credentials.";
            $secure_string = Read-Host "Enter your SQL password for user $($UserName)" -AsSecureString;            

            [System.Object]$encrypted = ConvertFrom-SecureString -SecureString $secure_string -Key $key;
            [System.Object]$encrypted | Set-Content $sql_password_full_name;    

            # Required for the very first run only when the password file is generated.
            [securestring]$sql_password = Get-Content $sql_password_full_name | ConvertTo-SecureString -Key $key;
            [pscredential]$cred = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $UserName, $sql_password;
            $cred.Password.MakeReadOnly();
            [System.Data.SqlClient.SqlCredential]$SqlCredential = New-Object System.Data.SqlClient.SqlCredential($cred.username, $cred.password);               
        } 
        # Load the file
        else 
        {                    
           [securestring]$sql_password = Get-Content $sql_password_full_name | ConvertTo-SecureString -Key $key;
           [pscredential]$cred = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $UserName, $sql_password;
           $cred.Password.MakeReadOnly();
           [System.Data.SqlClient.SqlCredential]$SqlCredential = New-Object System.Data.SqlClient.SqlCredential($cred.username, $cred.password);               
        }
    }
    catch
    {    
        Throw $_;
    }

     return $SqlCredential;
}



<#
    .SYNOPSIS
        This function creates an PSCredential object to be used for connectiong to SQL Server using SQL Server authentication mode.

    .DESCRIPTION
        The function creates a PSCredential object and saves it to file. Sacing the PSCredential object to file is a one time task.
        Once the file exists the function reads it from disk eliminating the need for a user input allowing scripts to be automated.       
    
    .PARAMETER FileFullName
        [string] The file name including the path where to store the encrypted password file. i.e.: "C:\Scrpits\sql_password.txt"

    .EXAMPLE                    
        $PSCredential = Get-Credentials -FileFullName 'C:\Scrpits\pscredentials.xml';

    .NOTES
        Author:  Yaniv Etrogi
        Website: https://sqlserverutilities.com/
        Email:   yaniv.etrogi@gmail.com 
#>
function Get-Credentials 
{
    [OutputType([System.Management.Automation.PSCredential])]
    param 
    (
        [Parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$false)]
        [string]$FileFullName
    )
    
    try
    {
        # If the file exists load the credentials.
        if (Test-Path -PathType Leaf -Path $FileFullName) 
        {        
            $PSCredential = Import-Clixml -Path $FileFullName;            
        } 
        else 
        {        
            # User Input.
            Write-Host "Credentials not found. Please enter your credentials.";
            [System.Management.Automation.PSCredential]$PSCredential = Get-Credential;

            # Save.
            [System.Management.Automation.PSCredential]$PSCredential | Export-Clixml -Path $FileFullName;        
        }
    }
    catch
    {    
        Throw $_;
    }

     return $PSCredential;
}


<#
.SYNOPSIS
    This function provides a simple way to access sql server.

.DESCRIPTION
    This function is a warapper over the NET Class: System.Data.SqlClient.
    The function supports 3 execution methods:
      1. ExecuteNonQuery (for DML (insert, update, delete, select...into, merge))
      2. ExecuteScalar   (to return a single value such as an Identity being retreived using SCOPE_IDENTITY() or any other salar value) 
      3. DataSet         (to return one or more data tables)
        
    The authentication methods supported:
      1. IntegratedSecurity (Windows Authentication).
      2. PSCredential for a connection using sql authentication.
      3. SQLCredential for a connection using sql authentication.

    IntegratedSecurity should always be the preffered method but when SQL Authentication is required use PSCredential or SQLredential.
    When using PSCredential or SQLredential use the 2 supporting functions: Get-Credentials and Get-SQLCredentials respectivly.
    Using these two funtions lets you create the credentials needed and save these to disk for future usage naking your script ready for automation, 
    that is no user input required (except for the very first time) for entering credential and in addition the credential are saved on disk as SecureString.


.PARAMETER SqlInstance
    [string] The sql instance we connect to.

.PARAMETER Database
    [string] The database we connect to.

.PARAMETER CommandText
    [string] The sql command to execute.

.PARAMETER CommandType
    [string] The sql command type to execute (NonQuery, Scalar, Dataset).

.PARAMETER IntegratedSecurity
    [switch] To be used when the required authentication is IntegratedSecurity also known as Windows Authentication. 

.PARAMETER PSCredential
    [PSCredential] Authenticate to the sql server instance using the PSCredential

.PARAMETER SQLCredential
    [SQLCredential] Authenticate to the sql server instance using the SQLCredential

.PARAMETER CommandTimeOut
    [int]The default value is 30 seconds. This value can be overriden if passing a value. Asign 0 for infinite.  

.PARAMETER ApplicationName 
    [string] The default is "PowerShellApp". This value can be overriden if passing a value.

.PARAMETER ApplicationIntent
    [switch] When this property is turned on the sql connection will be directed to an Alwayson Availabilty Group Secndary Replica.

.EXAMPLE
    # Generate a PSCredential object and save to file as a one time task. 
      From this point onwards the script can access the saved PSCredential stored in the file allowing the script to run as an automated task.
    
    # PSCredential for a connection using sql authentication.
    if (-not $PSCredential) 
    {
        $pscredentials_file_full_name = Join-Path $PSScriptRoot 'pscredentials.xml';
        $PSCredential = Get-Credentials -FileFullName $pscredentials_file_full_name;
    }


    # Generate an encrypted password and save to file as a one time task. 
      From this point onwards the script can access the saved password stored in the file allowing the script to run as an automated task.

    # SqlCredential for a connection using sql authentication.
    if (-not $SqlCredential) 
    {
        $sql_password_full_name = Join-Path $PSScriptRoot 'sql_password.txt';
        $SqlCredential = Get-SQLCredentials -FileFullName $sql_password_full_name -UserName 'CloudMonitoring';
    }


    
    # Scalar command type using IntegratedSecurity (Windows Authentication)
    $val = Invoke-Sqlcommand -SqlInstance $env:COMPUTERNAME -Database master -CommandText "SELECT @@SERVERNAME" -CommandType Scalar -IntegratedSecurity;

    # Scalar command type using PSCredential
    $val = Invoke-Sqlcommand -SqlInstance $env:COMPUTERNAME -Database master -CommandText "SELECT @@SERVERNAME" -CommandType Scalar -PSCredential $PSCredential;
    
    # Scalar command type using SqlCredential
    $val = Invoke-Sqlcommand -SqlInstance $env:COMPUTERNAME -Database master -CommandText "SELECT @@SERVERNAME" -CommandType Scalar -SqlCredential $SqlCredential;


    
    # DataSet command type using IntegratedSecurity (dynamic column names).
    [string]$query = "SELECT TOP (2) name, database_id, recovery_model_desc FROM sys.databases;";
    $ds = Invoke-Sqlcommand -SqlInstance $env:COMPUTERNAME -Database master -CommandText $query -CommandType DataSet -IntegratedSecurity
    # Iterate through each row
    foreach ($row in $ds.Tables[0].Rows) 
    {
        [string]$all_columns = $null;
        # Loop over each column index (dynamically based on DataTable columns)
        foreach ($columnIndex in 0..($ds.Tables[0].Columns.Count - 1)) 
        {
            [string]$columnName = $ds.Tables[0].Columns[$columnIndex].ColumnName;  # Get column name dynamically
            $value = $row[$columnName]  # Get the value dynamically using the column name        
            $all_columns += "$($columnName): $($value)  ";
        }
            Write-Host $all_columns;
    }



    # DataSet command type using IntegratedSecurity (static column names).
    [string]$query = "SELECT TOP (1) name, recovery_model_desc FROM sys.databases;";
    $ds = Invoke-Sqlcommand -SqlInstance $env:COMPUTERNAME -Database master -CommandText $query -CommandType DataSet -IntegratedSecurity:
    # Iterate through each row
    foreach ($row in $ds.Tables[0].Rows) 
    {
       $name = $row.Item('name');
       $recovery_model_desc = $row.Item('recovery_model_desc');
       Write-Host $name $recovery_model_desc;
    }


.NOTES
    Author:  Yaniv Etrogi
    Website: https://sqlserverutilities.com/
    Email:   yaniv.etrogi@gmail.com
#>
function Invoke-Sqlcommand
{
[CmdletBinding()]
Param
(    
    [Parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$false)]
    [string]$SqlInstance,

    [Parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$false)]
    [string]$Database,

    [Parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$false)]
    [string]$CommandText,

    [Parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$false)][ValidateSet('NonQuery' ,'Scalar' ,'DataSet')] 
    [string]$CommandType,       

    [Parameter(Mandatory=$false, ValueFromPipelineByPropertyName=$false)]
    [switch]$IntegratedSecurity,  
    
    [Parameter(Mandatory=$false, ValueFromPipelineByPropertyName=$false)]
    [System.Management.Automation.PSCredential]$PSCredential,

    [Parameter(Mandatory=$false, ValueFromPipelineByPropertyName=$false)]
    [System.Data.SqlClient.SqlCredential]$SqlCredential,
    
    [Parameter(Mandatory=$false, ValueFromPipelineByPropertyName=$false)]
    [int32]$CommandTimeOut = 30,

    [Parameter(Mandatory=$false, ValueFromPipelineByPropertyName=$false)]
    [string]$ApplicationName = "PowerShellApp",
    
    [Parameter(Mandatory=$false, ValueFromPipelineByPropertyName=$false)]
    [switch]$ApplicationIntent        
)

Begin
{
    # Validate that only one authentication method is provided        
    if ($PSCredential -and $SqlCredential) 
    {
        throw "You cannot provide both PSCredential and SqlCredential at the same time.";
    }
    if ($PSCredential -and $IntegratedSecurity) 
    {
        throw "You cannot provide both PSCredential and IntegratedSecurity at the same time.";
    }
    if ($SqlCredential -and $IntegratedSecurity) 
    {
        throw "You cannot provide both SqlCredential and IntegratedSecurity at the same time.";
    }

    
    try
    {
        # Construct connection string based on the authentication method.
        # IntegratedSecurity
        if ($IntegratedSecurity.IsPresent)
        {
            [string]$ConnectionString = "Server=$($SqlInstance); Database=$($Database); Integrated Security=$true; Application Name=$($ApplicationName);";                
        }            
        # PSCredential
        elseif($PSCredential)
        {
            #Convert PowerShell PSCredential to SQLCredential.
            $SecurePassword = $PSCredential.Password;
            $PasswordBSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecurePassword);
            $PlainPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($PasswordBSTR);
            
            [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($PasswordBSTR);
            [string]$ConnectionString = "Server=$($SqlInstance); Database=$($Database); Integrated Security=$false; User Id=$($PSCredential.UserName); Password=$PlainPassword; Application Name=$($ApplicationName);";                                
            $PlainPassword = $null;
        }    
        # SqlCredential 
        elseif($SqlCredential)
        {
            [string]$ConnectionString = "Server=$($SqlInstance); Database=$($Database); Integrated Security=$false; Application Name=$($ApplicationName); ";              
        }
       

        if ($ApplicationIntent.IsPresent) { $ConnectionString += " ApplicationIntent=ReadOnly;"; }


        #Write-Host $ConnectionString -ForegroundColor Green

        $SqlConnection = New-Object System.Data.SqlClient.SqlConnection($ConnectionString);
        if($SqlCredential){$SqlConnection.Credential = $SqlCredential; }
        
        $SqlCommand = $SqlConnection.CreateCommand();           
        $SqlCommand.CommandText = $CommandText;      
        $SqlCommand.CommandTimeout = $CommandTimeOut;
        $SqlConnection.Open(); 
                         
              
        # NonQuery
        if($CommandType -eq 'NonQuery')
        {      
            $sqlCommand.ExecuteNonQuery();                
            return;
        }

        # Scalar
        if($CommandType -eq 'Scalar')
        {                                   
            $Val = $sqlCommand.ExecuteScalar();                
            return $Val;
        }
        
        # DataSet
        if($CommandType -eq "DataSet")
        {
            $DataSet = New-Object System.Data.DataSet;
            $SqlDataAdapter = New-Object System.Data.SqlClient.SqlDataAdapter;
            $SqlDataAdapter.SelectCommand = $SqlCommand;    
            $SqlDataAdapter.Fill($DataSet) | Out-Null;                  
            return $DataSet;   
        }
    }
    catch
    {       
        Throw $_;
    }
    finally
    {
        If ($SqlConnection.State -eq 'Open') 
        {
            $SqlConnection.Close(); 
        }
        $SqlConnection.Dispose();
        $SqlCommand.Dispose();
        
        #[System.Data.SqlClient.SqlConnection]::ClearAllPools();  
    }
}
}

#endregion



# PSCredential for a connection using sql authentication.
if (-not $PSCredential) 
{
    $pscredentials_file_full_name = Join-Path $PSScriptRoot 'pscredentials.xml';
    $PSCredential = Get-Credentials -FileFullName $pscredentials_file_full_name;
}



# SqlCredential for a connection using sql authentication.
if (-not $SqlCredential) 
{
    $sql_password_full_name = Join-Path $PSScriptRoot 'sql_password.txt';
    $SqlCredential = Get-SQLCredentials -FileFullName $sql_password_full_name -UserName 'CloudMonitoring';
}


[string]$program_name = $MyInvocation.MyCommand.Name.Split(".")[0];
[string]$log_file_full_name = Join-Path $PSScriptRoot $program_name;
$log_file_full_name += '.log';


$val = Invoke-Sqlcommand -SqlInstance $env:COMPUTERNAME -Database master -CommandText "SELECT @@SERVERNAME" -CommandType Scalar -IntegratedSecurity;
$val = Invoke-Sqlcommand -SqlInstance $env:COMPUTERNAME -Database master -CommandText "SELECT @@SERVERNAME" -CommandType Scalar -PSCredential $PSCredential;
$val = Invoke-Sqlcommand -SqlInstance $env:COMPUTERNAME -Database master -CommandText "SELECT @@SERVERNAME" -CommandType Scalar -SqlCredential $SqlCredential;
Write-Text -Severity Info -Text $val -ForegroundColor Yellow -LogFileName $log_file_full_name;



# DataSet command type using IntegratedSecurity (dynamic column names).
[string]$query = "SELECT TOP (2) name, database_id, recovery_model_desc FROM sys.databases;";
$ds = Invoke-Sqlcommand -SqlInstance $env:COMPUTERNAME -Database master -CommandText $query -CommandType DataSet -IntegratedSecurity
# Iterate through each row
foreach ($row in $ds.Tables[0].Rows) 
{
    [string]$all_columns = $null;
    # Loop over each column index (dynamically based on DataTable columns)
    foreach ($columnIndex in 0..($ds.Tables[0].Columns.Count - 1)) 
    {
        [string]$columnName = $ds.Tables[0].Columns[$columnIndex].ColumnName;  # Get column name dynamically
        $value = $row[$columnName]  
        $all_columns += "$($columnName): $($value)  ";
    }
    Write-Text -Severity Info -Text $all_columns -ForegroundColor Yellow;
}



