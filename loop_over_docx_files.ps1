$topLevelFolder = Read-Host -Prompt 'Enter the top-level folder location you would like to loop over to create .rtf files from .docx files'
cd $topLevelFolder
$files = (Get-ChildItem -Path $topLevelFolder -Filter *.docx -Recurse -ErrorAction SilentlyContinue -Force).FullName
foreach ($file in $files) {
	$newFile = $file.substring(0, $file.LastIndexOf('.')) + '.rtf'
	
	# Create the RTF document.
	touch $newFile
	
	# Open both Word and RTF documents so I can manually copy/past.
	Invoke-Item $newFile
	Invoke-Item $file
	
	echo ''
	$next = Read-Host -Prompt 'Hit enter to continue...'
}

# Create backup location.
New-Item -ItemType Directory -Force -Path $topLevelFolder + '\MS Word Backups'
# Move the Word documents to backups folder. (Do this now instead of in above for-loop because user might accidentally hit enter when they didn't mean to.)
foreach ($file in $files) {
	Move-Item -Path $file -Destination $topLevelFolder + '\MS Word Backups'
}
