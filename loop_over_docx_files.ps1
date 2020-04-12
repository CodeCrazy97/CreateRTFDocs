cd 'C:\Users\Ethan\Documents\Writings'
$files = (Get-ChildItem -Path C:\Users\Ethan\Documents\Writings -Filter *.docx -Recurse -ErrorAction SilentlyContinue -Force).FullName
foreach ($file in $files) {
	$newFile = $file.substring(0, $file.LastIndexOf('.')) + '.rtf'
	
	# Create the RTF document.
	touch $newFile
	
	# Open both Word and RTF documents so I can manually copy/past.
	Invoke-Item $newFile
	Invoke-Item $file
	
	echo ''
	$User = Read-Host -Prompt 'Hit enter to continue...'
	
	# Move the Word document to backups folder.
	Move-Item -Path $file -Destination 'C:\Users\Ethan\Documents\Writings\MS Word Backups'
	
}
cd 'C:\Users\Ethan\Documents\Projects\Batch'
