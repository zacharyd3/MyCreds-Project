##Created by Zach Dokuchic at Oshki-Wenjack for use with the MyCreds project.

# Folder containing source CSV files
$folderPath = 'C:\IT\XML Conversion\Source CSV'

# Destination folder for the new files
$folderPathDest = 'C:\IT\XML Conversion\Destination XML'

# create a template Here-string for the XML (all <Person> nodes need to be inside a root node <Student>)
$xmlTemplate = @"
<?xml version="1.0" encoding="utf-8"?>
<Student>
@@STUDENTNODES@@
</Student>
"@

# and also a template for the individual <Person> nodes
# inside are placeholders '{0}' we will fill in later
$studentTemplate = @"
    <Person> 
		<SchoolAssignedPersonID>{0}</SchoolAssignedPersonID > 
		<Birth> 
			<BirthDate>{1}</Birthdate> 
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
			<AwardDate>{11}</AwardDate 
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
Get-ChildItem -Path $folderPath -Filter '*.csv' -File | ForEach-Object {
    
	# Combines destination path and file name with extension .xml
    $filePathdest = Join-Path -Path $folderPathDest -ChildPath ('{0}.xml' -f $_.BaseName)
	
    # Import the CSV file
    $data  = Import-Csv -Path $_.FullName
    $Students = foreach ($Student in $data) {
        # output a <Student> section with placeholders filled in
        $studentTemplate -f $Student.SchoolAssignedPersonID, $Student.BirthDate, $Student.FirstName, $Student.LastName, $Student.Email, $Student.AwardID, $Student.AwardLevel, $Student.AwardTitle, $Student.AwardDescription, $Student.AwardHonours, $Student.ProgramName, $Student.AwardDate, $Student.IssuedDate, $Student.IssuingBodyID, $Student.IssuingBodyName, $Student.CountryCode, $Student.URL
    }
    
	# create the completed XML and write this to file
    $xmlTemplate -replace '@@STUDENTNODES@@', ($Students -join [environment]::NewLine) | Set-Content -Path $filePathdest -Encoding utf8
}


	