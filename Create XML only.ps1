## Created by Zach Dokuchic | Oshki-Wenjack for use with the MyCreds project.
cls

# Added padding so both 
# files have the configuration 
# sections in the same spot

# Edit these files and locations to your instance
$folderDest = 'C:\IT\XML Conversion\Destination'
$sourcePath = 'C:\IT\XML Conversion\Source.csv'

# create a template Here-string for the XML (all <Person> nodes need to be inside a root node <Student>)
$xmlTemplate = @"
<?xml version="1.0" encoding="utf-8"?>
<Student xmlns="https://core.digitary.net/schema/mycreds/vsp/2022/11/01"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  xsi:schemaLocation="https://core.digitary.net/schema/mycreds/vsp/2022/11/01 https://core.digitary.net/schema/mycreds/vsp/2022/11/01">
@@STUDENTNODES@@
</Student>
"@

# and also a template for the individual <Person> nodes
# inside are placeholders '{0}' we will fill in later
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
			<AwardTitle>{7}</AwardTitle>
			<AwardDescription>{8}</AwardDescription>
			<AwardHonours>{9}</AwardHonours>
			<MicrocredentialAwardProgram>
				<ProgramName>{10}</ProgramName>
			</MicrocredentialAwardProgram>
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

$itemNumber=0
# Generate a list of all files in the folder and pipe it to ForEach-Object
Get-ChildItem -Path $sourcePath -Filter '*.csv' -File | ForEach-Object {  	
   
    # Import the CSV file
    $data = Import-Csv -Path $_.FullName

	$xmlOutput = foreach ($Student in $data) 
    {
		$studentTemplate -f $Student.SchoolAssignedPersonID, $Student.BirthDate, $Student.FirstName, $Student.LastName, $Student.Email, $Student.AwardID, $Student.AwardLevel, $Student.AwardTitle, $Student.AwardDescription, $Student.AwardHonours, $Student.ProgramName, $Student.AwardDate, $Student.IssuedDate, $Student.IssuingBodyID, $Student.IssuingBodyName, $Student.CountryCode, $Student.URL
    }

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
        Write-Output (-join('XML generated for ',$data[$itemNumber].FirstName,' ',$data[$itemNumber].LastName,', save location = ',$filePathdest))
        Write-Output ""
        $itemNumber++
    }
}	
Write-Output "------------------------------------"
Write-Output "Done"
Read-Host -Prompt "Press Enter to exit"