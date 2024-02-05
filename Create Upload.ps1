# Created by Zach Dokuchic at Oshki-Pimache-O-Win: The Wenjack Education Institute for use with the MyCreds project.
cls

function Get-CSV-Files {
    param (
        [Parameter(Mandatory = $true)]
        $folderPath
    )

    # Get all CSV files in the working directory
    $csvFiles = Get-ChildItem -Path $folderPath\* -Include "*.csv"

    if ($csvFiles.Count -eq 0) {
        Write-Error "Error: No CSV files found in folder."
        exit
    }
    else {
        if ($csvFiles.Count -eq 1) {
            return $csvFiles[0].FullName
        }
        else {
            
            # Print options for user to select
            Write-Host "Select the CSV file to process:"
    
            for ($i=0; $i -lt $csvFiles.Count; $i++) {
                Write-Host "$i) $($csvFiles[$i].Name)"
            }
    
            # Prompt user for selection
            $selectedOption = Read-Host "Enter the number of the CSV file to process (0-$($csvFiles.Count-1))"
    
            # Validate user input
            if ($selectedOption -notmatch '^\d+$' -or $selectedOption -ge $csvFiles.Count) {
                Write-Error "Error: Invalid selection."
                exit
            }
        }
        return $csvFiles[$selectedOption].FullName
    }
}

# Edit these files and locations to your instance
$workingDirectory = "$PSScriptRoot"
$workingDirectory = "\\oshkisv42\Office_Data\OSHKI-PIMACHE-O-WIN\Oshki Academics\MyCreds"
$folderDest = "$workingDirectory\Zip File"
$sourcePath = "$workingDirectory\Reports"
$sourceFile = Get-CSV-Files -folderPath $sourcePath
$accessChargeMethod = ""
$accessChargeAmount = ""
$accessChargeCurrency = ""
$accessChargePeriod = ""

# Set the batch name based on user input
$batchName = Read-Host -Prompt "What would you like to name this batch"
$completionSemester = Read-Host -Prompt "What Semester is this for (i.e Winter 2023)"
cls

# Setup variables
$csvExportLocation = $folderDest+"\"+$batchName+".csv"
$zipExportLocation = $folderDest+"\"+$batchName+".zip"
$itemNumber=0

# Adds the header to each XML file created
$xmlTemplate = @"
<?xml version="1.0" encoding="utf-8"?>
<Student xmlns="https://core.digitary.net/schema/mycreds/vsp/2022/11/01"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  xsi:schemaLocation="https://core.digitary.net/schema/mycreds/vsp/2022/11/01 https://core.digitary.net/schema/mycreds/vsp/2022/11/01">
@@STUDENTNODES@@
</Student>
"@

# Template for XML creation
$studentTemplate = @"

    <Person>
		<SchoolAssignedPersonID>{0}</SchoolAssignedPersonID>
		<Birth>
			<BirthDate>{1}</BirthDate>
		</Birth>
		<Name>
			<FirstName>{2}</FirstName>
			<LastName>{3}</LastName>
		</Name>
		<Contacts>
			<Email>{4}</Email>
		</Contacts>
	</Person>
	<LearningRecord>
		<MicrocredentialAward>
			<AwardID>{5}</AwardID>
			<AwardLevel>{6}</AwardLevel>
			<AwardTitle>
            <Title>{7}</Title>
            </AwardTitle>
			<AwardDescription>{8}</AwardDescription>
			<AwardHonours>{9}</AwardHonours>
			<AwardProgram>
				<ProgramName>{10}</ProgramName>
			</AwardProgram>
			<AwardDate>{11}</AwardDate>
			<IssuedDate>{12}</IssuedDate>
		</MicrocredentialAward>
		<MicrocredentialOrganization>
			<IssuingBodyID>{13}</IssuingBodyID>
			<IssuingBodyName>{14}</IssuingBodyName>
			<Contacts>
				<Address>
					<CountryCode>{15}</CountryCode>
				</Address>
			<URL>{16}</URL>
			</Contacts>
		</MicrocredentialOrganization>
	</LearningRecord>
	
"@

# Generate a list of all files in the folder and pipe it to ForEach-Object
Get-ChildItem -Path $sourceFile -File | ForEach-Object {  	

    # Import the CSV file
    $data = Import-Csv -Path $_.FullName
    $fullName = $data.FirstName+ " " +$data.LastName

	$xmlOutput = foreach ($Student in $data) 
    {$studentTemplate -f $Student.SchoolAssignedPersonID, $Student.BirthDate, $Student.FirstName, $Student.LastName, $Student.Email, $Student.AwardID, $Student.AwardLevel, $Student.AwardTitle, $Student.AwardDescription, $Student.AwardHonours, $Student.ProgramName, $Student.AwardDate, $Student.IssuedDate, $Student.IssuingBodyID, $Student.IssuingBodyName, $Student.CountryCode, $Student.URL}

    # Outputs the total number of rows found (debugging)
    Write-Output (-join('Total rows to process: ',$xmlOutput.count))
    Write-Output "------------------------------------"
    Write-Output ""

    # Iterates through the xmlOutput array
    while ($itemNumber -lt ($xmlOutput.count))
	{
        
        # Outputs which row is currently being checked
        Write-Output (-join('Checking Row: ',($itemNumber+1)))

        # Combines destination path and file name with extension .xml
	    $filePathdest = (-join($folderDest,'\',$data[$itemNumber].SchoolAssignedPersonID,'.xml'))
        
        # Generate and save the XML if there is only 1 row, the variable stays a string the if statement manipulates it based on those details.
        if ($xmlOutput.Count -eq 1){
            $xmlTemplate -replace '@@STUDENTNODES@@',$xmlOutput | Set-Content -Path $filePathdest -Encoding utf8
        }
        if ($xmlOutput.Count -gt 1){
            $xmlTemplate -replace '@@STUDENTNODES@@',$xmlOutput[$itemNumber] | Set-Content -Path $filePathdest -Encoding utf8
        }

        # Output the log as files are generated
        Write-Output (-join('XML generated for ',$data[$itemNumber].FirstName,' ',$data[$itemNumber].LastName))

        # Convert the Award Title into the document type via parsing
        $documentType = $data[$itemNumber].AwardTitle.toLower()
        $documentType = $documentType -replace '\s','_'

        # Create the output CSV files
        Write-Output (-join(($data[$itemNumber].FirstName,' ',$data[$itemNumber].LastName,' added to the destination CSV')))
        Write-Output ""
        $newRow = [PSCustomObject] @{
        "id" = $data[$itemNumber].SchoolAssignedPersonID;
        "fullName" = $data[$itemNumber].FirstName+ " " +$data[$itemNumber].LastName;
        "email" = $data[$itemNumber].Email;
        "file" = $data[$itemNumber].SchoolAssignedPersonID+".xml";
        "documentType" = $documentType;
        "display_name" = $data[$itemNumber].AwardTitle;
        "initial_login_type" = "email";
        "initial_login_idp" = "digitary";
        "initial_login_value" = $data[$itemNumber].Email;
        "access_charge_method" = $accessChargeMethod;
        "access_charge_amount" = $accessChargeAmount;
        "access_charge_currency" = $accessChargeCurrency;
        "access_charge_period" = $accessChargePeriod;
        "Batch Name" = $batchName;
        "Program ID" = $data[$itemNumber].ProgramName;
        "Completion Semester" = $completionSemester;}

        # Add to the item number variable to setup the next loop
        $itemNumber++

        # Export to the CSV for each row (same csv each time)
        $newRow | Export-CSV $csvExportLocation -Force -NoTypeInformation -Append
        }
    }

# Deletes any old archives from previous runs
Remove-Item $folderDest\*.zip -Exclude *$batchName*.zip

# Creates the ZIP file for upload
Write-Output "------------------------------------"
(-join('Created Archive: ',($zipExportLocation)))
Compress-Archive -Path $folderDest\* -DestinationPath $zipExportLocation -Force

# Deletes all files created during the process except the archive
Remove-Item $folderDest\* -Exclude *.zip
Write-Output "------------------------------------"
Write-Output "Done"
Start-Sleep -Seconds 15
