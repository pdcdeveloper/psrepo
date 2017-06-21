<##################################################################################
    
    Export-CsvFromWorkSchedule

Date            : Monday, June 19, 2017
Original Author : pdcdeveloper (https://github.com/pdcdeveloper)
Co-Authors      : 

The MIT License (MIT)

Copyright (c) 2016 pdcdeveloper

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.


Synopsis:
    See function Export-CsvFromWorkScheduleSelection.SYNOPSIS.
    
Description:
    See function Export-CsvFromWorkScheduleSelection.DESCRIPTION.
    
For co-authors:
    Scroll all the way down to MAIN.
    
    It is recommended to use the directory path provided by the $PWD variable.
This is the present working directory and will always be the directory path that
this script is located in.
    
Revisions:

    +20170619
        

##################################################################################>


function Export-CsvFromWorkScheduleSelection {
<#
.SYNOPSIS
    This function will return a ranges of rows from an employee's column by parsing through
a column of dates.  Both columns will be exported to the given output file location as csv
if provided an output directory.


.DESCRIPTION
    The input csv file must have a column of date strings and a unique column for each employee.
See the test values 'workschedule.csv' at the bottom of this script.

 Date string values are parsed using the .NET DateTime struct's static TryParse() method.  Not all formats
have been tested, but date strings in the format "yyyy/MM/dd" will work.  For more formats, see
(on Windows) Start>Settings>Time & language>Date & time.

    The values under the given employee name are read and printed, but never parsed.  It is assumed
that the values are in a 'calendar memo' format and must be interpreted by the recipient of this script's
output.


.PARAMETER [string]InputFilePath
    The literal path of the input file.  Must be comma-seperated values.


.PARAMETER [string]OutputDirectory
    The literal path of the output directory.  The filename is automatically generated using
the Employee parameter and will have a '.csv' extension.  The path to the directory and file
is tested before creating the file. If this parameter is not used, the results of this function
will not be exported.


.PARAMETER [string]Employee
    Column header.  Each row in this column has information about the employee's work schedule.


.PARAMETER [string]DatesHeader
    Column header.  Each row in this column will be parsed for date strings.


.PARAMETER [string]StartingDate
    Determines the starting index by comparing against the row of values in the Dates column.
An example valid date string formats is "yyyy/MM/dd".
    This parameter is for historical reasons.  If this parameter is not used, the value
[DateTime]::Today will be used everytime this function is called.


.PARAMETER [int]MaxResults
    The desired range of values from StartingDate.


.PARAMETER [Switch]OverwriteOutputFile
    Calling this parameter will overwrite the file generated at the output directory.  There is
no confirmation dialog.

#>
    [CmdletBinding()]
    Param (
        [parameter(Mandatory=$true)]
        [alias("LiteralFilePath")]
        [ValidateNotNullOrEmpty()]
        [string]$InputFilePath,

        [parameter(Mandatory=$false)]
        [alias("OutputPath")]
        [ValidateNotNullOrEmpty()]
        [string]$OutputDirectory,

        [parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [alias("DatesColumn","DatesHeaderName")]
        [string]$DatesHeader,

        [parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [alias("EmployeeName","EmployeeId")]
        [string]$Employee,
        
        [parameter(Mandatory=$false)]
        [ValidateNotNullOrEmpty()]
        [string]$StartingDate,

        [parameter(Mandatory=$false)]
        [ValidateRange(1,256)]
        [int]$MaxResults = 42,

        [parameter(Mandatory=$false)]
        [Switch]$OverwriteOutputFile
    )



    Process {
        # Input file path validation.
        if (!(Test-Path $InputFilePath)) {
            Write-Error -Message "Input file was not found.";
            return;
        }

        Write-Host "Input file path is $InputFilePath";


        # Output path validation.
        if (!([string]::IsNullOrEmpty($OutputDirectory))) {
            if (!(Test-Path $OutputDirectory)) {
                Write-Error -Message "Output directory path is invalid.";
                return;
            }
        }


        # Locals.  Trust the GC.
        $inputCsv = Import-Csv -LiteralPath $InputFilePath;
        [int]$startingIndex = -1;
        $start = [DateTime]::Today;
        $date = [DateTime]::Today;


        # Extra validation for the date column.
        if (!($inputCsv | Get-Member -Name $DatesHeader)) {
            Write-Error -Message "Unable to find the dates column.";
            return;
        }


        # Extra validation for the employee column.
        if (!($inputCsv | Get-Member -Name $Employee)) {
            Write-Error -Message "Unable to find the employee column.";
            return;
        }


        # Extra validation for the starting date.
        if (!([string]::IsNullOrEmpty($StartingDate))) {
            if (!([DateTime]::TryParse($StartingDate, [ref]$start))) {
                Write-Error -Message "Starting date input is not in a valid format.  Today's date will be used.";
                $start = [DateTime]::Today;
            }
        }


        # Compare the first date.
        if (!([DateTime]::TryParse($inputCsv.$DatesHeader[0], [ref]$date))) {
            Write-Error -Message "Unable to parse the date values under the date column.";
            return;
        }
        elseif ($date -gt $start) {
            Write-Host "The DeLorean did not reach 88mph.  Missing flux capacitor.";
            $startingIndex = 0;
        }
        elseif ($date -eq $start) {
            $startingIndex = 0;
        }


        # Compare the last date.
        if ($startingIndex -lt 0) {
            if (!([DateTime]::TryParse($inputCsv.$DatesHeader[$inputCsv.Length - 1], [ref]$date))) {
                Write-Error -Message "Unable to parse the date values under the date column.";
                
                return;
            }
            elseif ($date -lt $start) {
                Write-Host "Missing crystal ball.";
                $startingIndex = $inputCsv.Length - 1;
            }
            elseif ($date -eq $start) {
                $startingIndex = $inputCsv.Length - 1;
            }
        }


        # Iterate through the csv's date column to find the starting index.
        # Truncated to skip the first and last indices.
        if ($startingIndex -lt 0) {
            for ($i = 1; $i -lt $inputCsv.Length - 1; $i++) {
                if (!([DateTime]::TryParse($inputCsv.$DatesHeader[$i], [ref]$date))) {
                    Write-Error -Message "Unable to parse the date values under the date column.";
                    return;
                }

                if ($date -eq $start) {
                    $startingIndex = $i;
                    break;
                }
            }
        }


        # Make sure a starting index was found.
        if ($startingIndex -lt 0 -or $startingIndex -ge $inputCsv.Length) {
            Write-Error -Message "Unable to find a valid starting index.  $startingIndex";
            return;
        }



        # Filter the csv and collect the results.
        $prettyPrint = $null;
        if (!([string]::IsNullOrEmpty($OutputDirectory))) {
            # Remove illegal characters from the file name.
            $outFileName = $Employee;
            $outFileName = $outFileName.Replace("\","");
            $outFileName = $outFileName.Replace("/","");
            $outFileName = $outFileName.Replace(":","");
            $outFileName = $outFileName.Replace("*","");
            $outFileName = $outFileName.Replace("?","");
            $outFileName = $outFileName.Replace("`"","");
            $outFileName = $outFileName.Replace("<","");
            $outFileName = $outFileName.Replace(">","");
            $outFileName = $outFileName.Replace("|","");

            $ofp = $OutputDirectory.TrimEnd('\') + '\' + $outFileName + " work schedule.csv";
            Write-Host "Writing to $ofp";

            # TODO: Check if the InputFilePath is the same as the generated output file path.  For real...

            
            if ($OverwriteOutputFile) {
                $inputCsv[$startingIndex..($startingIndex + $MaxResults - 1)] | Select-Object $DatesHeader, $Employee -OutVariable prettyPrint | Export-Csv -LiteralPath $ofp -NoTypeInformation -Encoding UTF8;
            }
            else {
                $inputCsv[$startingIndex..($startingIndex + $MaxResults - 1)] | Select-Object $DatesHeader, $Employee -OutVariable prettyPrint | Export-Csv -LiteralPath $ofp -NoTypeInformation -Encoding UTF8 -NoClobber;
            }
        }
        else {
            $inputCsv[$startingIndex..($startingIndex + $MaxResults - 1)] | Select-Object $DatesHeader, $Employee -OutVariable prettyPrint;
        }

        # Show
        $prettyPrint | Format-Table

        return;
    }
}




<##################################################################################
MAIN
##################################################################################>

# Ease of access.
$EMPLOYEE_NAME = "employee header";
$DATES_COLUMN = "date header";
$STARTING_DATE = "";
$MAX_RESULTS = "256";
$INPUT_FILE_PATH = "$PWD\workschedule.csv";
$OUTPUT_DIRECTORY = $PWD;
$ALWAYS_OVERWRITE_FILE = $true;



if ($ALWAYS_OVERWRITE_FILE) {
    Export-CsvFromWorkScheduleSelection -InputFilePath $INPUT_FILE_PATH -OutputDirectory $OUTPUT_DIRECTORY -Employee $EMPLOYEE_NAME -DatesHeader $DATES_COLUMN -StartingDate $STARTING_DATE -MaxResults $MAX_RESULTS -OverwriteOutputFile
}
else {
    Export-CsvFromWorkScheduleSelection -InputFilePath $INPUT_FILE_PATH -OutputDirectory $OUTPUT_DIRECTORY -Employee $EMPLOYEE_NAME -DatesHeader $DATES_COLUMN -StartingDate $STARTING_DATE -MaxResults $MAX_RESULTS
}
