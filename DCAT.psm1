# 
#   Module      : DCAT 
#   Description : This module defines functions that support creation of DCAT files.
#   Creator  : matthew.lawler@mdba.gov.au 

#	Usage : 
#	The module needs to be imported obviously 
#	Import-Module DCAT -Force -Verbose 

#	Source directory - this is the directory from which recursion begins. 
#	Set-Location "G:\IT Languages\Powershell\Test_All_File_Types" 
#	Set-Location "G:\IT Languages\Powershell\New" 
#	Or J:\working documents\basin plan division\ecohydrology \projecs \ 2017Evaluation + ESLTOngoing? + HydrologyReporting + Articles 
#	J:\Working Documents and Drafts\Basin Plan Division\Eco-hydrology Analysis Branch\Projects
#	Set-Location	"J:\Working Documents and Drafts\Basin Plan Division\Eco-hydrology Analysis Branch\Publications"

#	Pipeline composable functions call 
#	Get-ChildItem -File -Recurse | 
#		Convert-FileToDCATProps | 
#		Convert-DCATPropsToDCATObj | 
#		Convert-DCATObjToDCATXML | 	
#	all DCAT files are copied to the target directory 
#		Add-SafeFile -TargetDir "J:\Working Documents and Drafts\Corporate Services\ICT Services\Data Management\CKAN\Initial" 

#############################

function Convert-FileToDCATProps{
#	::	[System.IO.FileInfo] -> Hashtable 
#	::	fileObject -> Hashtable
#	This is a pipelinable function. 
#	It takes in a file object, and produces a hashtable object.  
#	This is the subset of properties useful for constructing DCAT files. 

	[cmdletbinding()]
	param(
#	The $_. or $PSItem is the closure item, so it is always the last parm 
    [Parameter(
      Position=0,
#	0 is first argument 
      Mandatory=$true,
#	get item from pipeline 
      ValueFromPipeline=$true 
    )]
#	source is the original file 
    [System.IO.FileInfo]
    $source 	
	)

	process{ 

#	get all the properties from COMObject 
	$dirname = Split-Path -Path ($_.FullName) 
	$File = Split-Path ($_.FullName) -Leaf 
	
	$objShell = New-Object -ComObject Shell.Application
	$objFolder = $objShell.namespace($dirname) 
	
	$Item = $objFolder.items().item($File) 

	$fileProps = [pscustomobject]@{ 
#	get the directly available file properties 
		FullName = ($_.FullName) 
		Name_File = ($_.Name) 
		Extension = ($_.Extension) 
#	use DateCreated and DateModified instead as almost equal 
#		CreationTimeUtc = ($_.CreationTimeUtc) 
#		LastWriteTimeUtc = ($_.LastWriteTimeUtc) 
		
#	get the COMObject DCAT relevant file properties 
		Size = $objFolder.GetDetailsOf($Item, 1) 
		DateCreated = $objFolder.GetDetailsOf($Item, 4) 		
		DateModified = $objFolder.GetDetailsOf($Item, 3)
		Owner = $objFolder.GetDetailsOf($Item, 10) 
		Authors = $objFolder.GetDetailsOf($Item, 20) 	
		Company = $objFolder.GetDetailsOf($Item, 33) 	
#	this fixes the strange formating so the date matches "\d{1,2}\/\d{2}\/\d{4}\s\d{1,2}:\d{2}\s(A|P)M" or [datetime]::ParseExact($DateCreated, "d/MM/yyyy h:mm tt", $null) 
		ContentCreated = (($objFolder.GetDetailsOf($Item, 147) -replace '\/.', '/') -replace '\s\W\W', ' ') -replace '^\W', '' 
		DateLastSaved = (($objFolder.GetDetailsOf($Item, 149) -replace '\/.', '/') -replace '\s\W\W', ' ' ) -replace '^\W', '' 
		Language = $objFolder.GetDetailsOf($Item, 194) 			
		URL = $objFolder.GetDetailsOf($Item, 199) 	
		Tags = $objFolder.GetDetailsOf($Item, 18) 
		Subject = $objFolder.GetDetailsOf($Item, 22) 	
		Categories = $objFolder.GetDetailsOf($Item, 23) 
		Comments = $objFolder.GetDetailsOf($Item, 24) 

#	Duplicates: Name_File = Name ; Extension = FileExtension ; FullName = Path 
#		Name = $objFolder.GetDetailsOf($Item, 0) 
#		FileExtension = $objFolder.GetDetailsOf($Item, 159) 
#		Path = $objFolder.GetDetailsOf($Item, 189) 	
#	not needed
#		ContentStatus = $objFolder.GetDetailsOf($Item, 129) 
#		LinkTarget = $objFolder.GetDetailsOf($Item, 198) 		
	} 
	write-verbose ("1: fileProps = $fileProps ")  
	return $fileProps 
	}
}

#############################

function Get-ExtensionToMimeType{
#	::	String -> String 
#	This is NOT a pipelinable function.
#	It is not used elsewhere, so there is no need to export it.  
#	This enables a lookup from extension to Mime type.  
 
	[cmdletbinding()]
	param(
	[Parameter(
      Position=0,
#	0 is first argument 
      Mandatory=$true, 
      ValueFromPipeline=$false 
    )] 
#	does a static type check to ensure that the parm is a valid extension 
	[ValidateSet('.doc', '.docx', '.pps', '.ppt', '.pptx', '.vsd', '.vsdx', '.xls', '.xlsm', '.xlsx', '.accdb', '.b5', '.bat', '.bmp', '.cabal', '.cdr', '.css', '.csv', '.dat', '.db', '.dbc', '.dbm', '.ddl', '.emx', '.ent', '.epx', '.erwin', '.exe', '.gif', '.hs', '.htm', '.html', '.ibak', '.ini', '.jpeg', '.jpg', '.js', '.jsonld', '.jsp', '.kml', '.kmz', '.lhs', '.lnk', '.local', '.log', '.md', '.mdb', '.mht', '.mp3', '.mp4', '.mpp', '.msg', '.mvt', '.njx', '.org', '.pdf', '.png', '.ps', '.ps1', '.ps1xml', '.psafe3', '.rdf', '.rtf', '.SQL', '.sqlplus', '.svg', '.tex', '.text', '.thmx', '.tif', '.tmp', '.tr5', '.tsv', '.twb', '.txt', '.ucs', '.url', '.vss', '.vssx', '.webp', '.wmf', '.wmv', '.xml', '.xsd', '.zip')] 
#	Extension is the file type 
	[string]
    $Extension 
	)

	process{ 
		$ExtensionToMimeType = [pscustomobject]@{ 
		
		".doc" = "application/msword"
		".docx" = "application/vnd.ms-word.document.macroEnabled.12"
		".pps" = "application/vnd.ms-powerpoint"
		".ppt" = "application/vnd.ms-powerpoint"
		".pptx" = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
		".vsd" = "application/vnd.visio"
		".vsdx" = "application/vnd-ms-visio.drawing"
		".xls" = "application/vnd.ms-excel"
		".xlsm" = "application/vnd.ms-excel.sheet.macroEnabled.12"
		".xlsx" = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
		".accdb" = "application/msaccess"
		".b5" = "application/octet-stream"
		".bat" = "application/bat"
		".bmp" = "image/bmp"
		".cabal" = "application/octet-stream"
		".cdr" = "image/x-coreldraw"
		".css" = "text/css"
		".csv" = "text/csv"
		".dat" = "text/plain"
		".db" = "application/octet-stream"
		".dbc" = "application/octet-stream"
		".dbm" = "application/octet-stream"
		".ddl" = "application/sql"
		".emx" = "application/octet-stream"
		".ent" = "text/plain"
		".epx" = "application/octet-stream"
		".erwin" = "application/octet-stream"
		".exe" = "application/x-ms-dos-executable"
		".gif" = "image/gif"
		".hs" = "application/octet-stream"
		".htm" = "text/html"
		".html" = "text/html"
		".ibak" = "application/octet-stream"
		".ini" = "text/plain"
		".jpeg" = "video/JPEG"
		".jpg" = "image/jpeg"
		".js" = "application/javascript"
		".jsonld" = "application/ld+json"
		".jsp" = "application/jsp"
		".kml" = "application/vnd.google-earth.kml+xml"
		".kmz" = "application/vnd.google-earth.kmz"
		".lhs" = "application/octet-stream"
		".lnk" = "application/octet-stream"
		".local" = "application/octet-stream"
		".log" = "text/plain"
		".md" = "application/octet-stream"
		".mdb" = "application/x-msaccess"
		".mht" = "application/vnd.pwg-multiplexed"
		".mp3" = "audio/mpeg"
		".mp4" = "video/mp4"
		".mpp" = "application/vnd.ms-project"
		".msg" = "application/x-msg"
		".mvt" = "application/octet-stream"
		".njx" = "application/octet-stream"
		".org" = "application/vnd.lotus-organizer"
		".pdf" = "application/pdf"
		".png" = "image/png"
		".ps" = "application/postscript"
		".ps1" = "text/plain"
		".ps1xml" = "text/plain"
		".psafe3" = "application/octet-stream"
		".rdf" = "application/rdf+xml"
		".rtf" = "application/rtf"
		".SQL" = "application/sql"
		".sqlplus" = "text/plain"
		".svg" = "image/svg-xml"
		".tex" = "application/x-tex"
		".text" = "text/plain"
		".thmx" = "application/vnd.ms-officetheme"
		".tif" = "image/tiff"
		".tmp" = "application/octet-stream"
		".tr5" = "application/octet-stream"
		".tsv" = "text/tab-separated-values"
		".twb" = "application/octet-stream"
		".txt" = "text/plain"
		".ucs" = "application/octet-stream"
		".url" = "application/x-url"
		".vss" = "application/vnd-visio"
		".vssx" = "application/vnd-visio"
		".webp" = "image/webp"
		".wmf" = "image/x-wmf"
		".wmv" = "video/x-ms-wmv"
		".xml" = "text/xml"
		".xsd" = "text/xml"
		".zip" = "application/zip"

		} 
#		write-verbose ("1: ExtensionToMimeType = $ExtensionToMimeType ")  
		return $ExtensionToMimeType.$Extension  
	} 
} 

#############################

function Convert-DCATPropsToDCATObj{
#	::	[System.IO.FileInfo] -> Hashtable 
#	::	fileObject -> Hashtable
#	This is a pipelinable function. 
#	It takes in a DCAT properties custom object, and produces a DCAT object.  
#	This does mapping from file properties to DCAT terms with renaming and some transformations. 
 
	[cmdletbinding()]
	param(
#	The $_. or $PSItem is the closure item, so it is always the last parm 
    [Parameter(
      Position=0,
#	0 is first argument 
      Mandatory=$true,
#	get item from pipeline 
      ValueFromPipeline=$true 
    )]
#	source is the original file 
    [pscustomobject]
    $source 	
	)

	process{ 
	
	$guid = [guid]::NewGuid().ToString() 

	$DCATObject = [pscustomobject]@{ 
#	follows mapping order in 'draft metadata mapping v2.0.xls' 
#	see https://data.gov.au/dataset/ds-dga-e207243a-0b06-49c9-a232-95bb7eb442bb/details

		Title = ($_.Name_File) 
#	Does not get default when there are no comments - tried $null, but why does "" not work? 
		Description = if( ($_.Subject) -eq "" ) { if( ($_.Commments) -eq "" ) { 'There were no Subject or Commments in document metadata.' } else { ($_.Commments) } } else { ($_.Subject) } 
#	Maybe these should be separable somehow into an array? 	
		Keyword = if( ($_.Tags) -eq "" ) { if( ($_.Categories) -eq "" ) { 'There were no Tags or Categories in document metadata.' } else { ($_.Categories) } } else { ($_.Tags) } 	
#	after parsing, derive from document somehow  	
		Theme = "Healthy Working Basin" 
		
		Language = if( ($_.Language) -eq "" ) { "en-AU" } else { ($_.Language) } 	

		Licence = "$($_.Name_File) (c) by Murray-Darling Basin Authority. This work is licensed under a Creative Commons Attribution-ShareAlike 3.0 Unported License." 

		Rights = "Unclassified" 	

		DataStatus = "Inactive" 	

		UpdateFrequency = "As Required" 	

		ExposeUserContactInformation = "Yes" 
#	full path to file  
		LandingPage = ($_.FullName) 

#	dates 
		PublishDate = if( ($_.ContentCreated) -eq "" ) { ($_.DateCreated) } else { ($_.ContentCreated) }
		UpdateDate = if( ($_.DateLastSaved) -eq "" ) { ($_.DateModified) } else { ($_.DateLastSaved) } 
		
		Identifier = $guid 
#	Data portal address may change. 	
		DataPortal = "http://mdba.data.gov.au" 	
		MetadataURI = -join( "http://mdba.data.gov.au", "/", $guid ) 
#	Document access protocol needs to be agreed. Maybe ftp, or added in CKAN? 		
		DownloadURL = ($_.FullName) 
		
		FileSize = ($_.Size) 
		
		AccessURL = if( ($_.URL) -eq "" ) { "https://www.mdba.gov.au/" } else { ($_.URL) } 	
#	Use Get-ExtensionToMimeType to derive Mime Type 
		MediaType = 
		if( ('.doc', '.docx', '.pps', '.ppt', '.pptx', '.vsd', '.vsdx', '.xls', '.xlsm', '.xlsx', '.accdb', '.b5', '.bat', '.bmp', '.cabal', '.cdr', '.css', '.csv', '.dat', '.db', '.dbc', '.dbm', '.ddl', '.emx', '.ent', '.epx', '.erwin', '.exe', '.gif', '.hs', '.htm', '.html', '.ibak', '.ini', '.jpeg', '.jpg', '.js', '.jsonld', '.jsp', '.kml', '.kmz', '.lhs', '.lnk', '.local', '.log', '.md', '.mdb', '.mht', '.mp3', '.mp4', '.mpp', '.msg', '.mvt', '.njx', '.org', '.pdf', '.png', '.ps', '.ps1', '.ps1xml', '.psafe3', '.rdf', '.rtf', '.SQL', '.sqlplus', '.svg', '.tex', '.text', '.thmx', '.tif', '.tmp', '.tr5', '.tsv', '.twb', '.txt', '.ucs', '.url', '.vss', '.vssx', '.webp', '.wmf', '.wmv', '.xml', '.xsd', '.zip') -contains ($_.Extension) ) 
			{ Get-ExtensionToMimeType ($_.Extension) } 
			else 
#	apparently this is the default MIME type 
			{ 'application/octet-stream' } 
		
		Format = ($_.Extension) 
#	FOAF record may be needed rather than just the name 
		Publisher = if( ($_.Company) -eq "" ) { "Murray-Darling Basin Authority" } else { ($_.Company) } 
#	This is a sample only. Need to derive from Authors using a lookup. 
		Publisher_user = if( ($_.Authors) -eq "" ) { ($_.Owner) } else { ($_.Authors) } 
		Contact = "BEGIN:VCARD
VERSION:3.0
N:Bradshaw;Ben;;;
FN:Ben Bradshaw
EMAIL;type=INTERNET;type=WORK;type=pref:Ben.Bradshaw@mdba.gov.au
TEL;type=WORK;type=pref:+1 612 6279 0155 
X-ABUID:5AD380FD-B2DE-4261-BA99-DE1D1DB52FBE\:ABPerson
END:VCARD" 

		Jursidiction = "Commonwealth of Australia" 
		Homepage = "https://www.mdba.gov.au/" 
#	These dates require parsing within the file to discover earliest date mentioned.  	
		TemporalCoverageFrom = if( ($_.ContentCreated) -eq "" ) { ($_.DateCreated) } else { ($_.ContentCreated) }	
		TemporalCoverageTo = if( ($_.DateLastSaved) -eq "" ) { ($_.DateModified) } else { ($_.DateLastSaved) }			
#	This requires document parsing and a lookup based on the gazetteer to derive the latitude longitude box. 
		Geospatial = "Murray-Darling Basin" 

	} 

	write-verbose ("1: DCATObject = $DCATObject ")  
	return $DCATObject 
	}
}

#############################

function Convert-DCATObjToDCATXML{
#	::	[System.IO.FileInfo] -> Hashtable 
#	::	fileObject -> Hashtable
#	This is a pipelinable function. 
#	It takes in a file object, and produces a hashtable object.  
#	Note that the DCAT XML tags sometimes have different names than the DCAT object. 

#	to do 
#	1. XML type causes some errors - String is more permissive. Check this 
#	2. These properties do not have defined XML tags:  
#	DataStatus ExposeUserContactInformation MetadataURI DataPortal Jursidiction Homepage TemporalCoverageTo 
#	3.	Temporal data only has a single tag? 
 
  [cmdletbinding()]
  param(

    [Parameter(
      Position=0,
#	0 is first argument 
      Mandatory=$true,
#	get item from pipeline 
      ValueFromPipeline=$true 
    )]
#	properties is a hashtable [(k,v)] wrapped in a pscustomobject 
    [pscustomobject]
    $source 	
  )

  process{ 
  
	$DCATXML = ([String]"<dcat:Dataset rdf:about=`"$($_.LandingPage)`"> 
	<dct:title>$($_.Title)</dct:title> 
	<dct:description>$($_.Description)</dct:description> 	
	<dcat:keyword>$($_.Keyword)</dcat:keyword> 
	<dcat:theme>$($_.Theme)</dcat:theme> 	
	<dct:language>$($_.Language)</dct:language>
	<dct:license>$($_.License)</dct:license>	
	<dct:rights>$($_.Rights)</dct:rights>	
	<dct:accrualPeriodicity>$($_.UpdateFrequency)</dct:accrualPeriodicity>		
	<dct:landingPage>$($_.LandingPage)</dct:landingPage> 
	<dct:created>$($_.PublishDate)</dct:created> 
	<dct:last_modified>$($_.UpdateDate)</dct:last_modified> 	
	<dct:identifier>$($_.Identifier)</dct:identifier> 	
	<dcat:downloadURL>$($_.AccessURL)</dcat:downloadURL> 	
	<dcat:byteSize>$($_.FileSize)</dcat:byteSize> 	
	<dct:accessURL>$($_.AccessURL)</dct:accessURL> 
	<dcat:mediaType>$($_.MediaType)</dcat:mediaType> 	
	<dcat:format>$($_.Format)</dcat:format>	
	<dct:publisher>$($_.Publisher)</dct:publisher> 	
	<dcat:contactPoint>$($_.Publisher_user)</dcat:contactPoint> 
	<dct:temporal>$($_.TemporalCoverageFrom)</dct:temporal> 
	<dct:spatial>$($_.Geospatial)</dct:spatial> 	
	<dct:api_type>$($_.MediaType)</dct:api_type> 	
	</dcat:Dataset>") 

	write-verbose ("1: DCATXML = $DCATXML ")  

	return $DCATXML 
	}
}

#############################

function Convert-FilepathToFilename{ 
#	::	String -> String 
#	This is NOT a pipelinable function.
#	must follow path format rules, but not tested, of course. 
#	replaces \ and whitespace with _ in directory path 
#	Example: 
#	$source = "G:\IT Languages\Powershell\New" 
#	-> 
#	$target = "G_IT_Languages_Powershell_New" 
  [cmdletbinding()]
  param(
    [Parameter(
      Position=0,
#	0 is first argument 
      Mandatory=$true, 
      ValueFromPipeline=$false  
    )] 
	[String] 
    $source,
    [Parameter(
      Position=1,
      Mandatory=$true, 
      ValueFromPipeline=$false  
    )] 
	[Int] 
    $maxLength 
  )
  
#	use regex to replaces \ and whitespace with underscore, and removes : 
#	260 chars which is the maximum length for a path (file name and its directory route) 
#	use this length to truncate the string 
#	Dirpaths that produce duplicates shold not be a problem, as the XML will still be added to a file 
	$filename = (((-join (Split-Path -Path ($source))[0..$maxLength]) -replace '\\', '_') -replace '\s', '_') -replace ':', '' 
	write-verbose ("$_ is a file; new filename is $filename") 
	return $filename 
}

#############################

function Add-SafeFile{
#	::	String -> String -> [String] 
#	::	targetDir -> pipeline -> file 
#	This is a pipelinable function. 
#	typing rules only apply to input 
#	output type is defined by process {} output 
#	Usage: composable function to write object to a file. 
#	Really a pseudo monad. 

  [cmdletbinding()]
  param(
    [Parameter(
      Position=0,
#	0 is first argument 
      Mandatory=$true, 
      ValueFromPipeline=$false 
    )] 
#	targetDir is the location of all DCAT files 
	[string]
    $TargetDir, 
#	The $_. or $PSItem is the closure item, so it is always the last parm 
    [Parameter(
      Position=1,
#	1 is second argument 
      Mandatory=$true,
#	get item from pipeline 
      ValueFromPipeline=$true 
    )]
#	content is whatever has to added as content to a file 
    [String]
    $content 	
  )

  process{
#	taken from about tag which is extrscted from pipeline 
#	<dcat:Dataset rdf:about="G:\IT Languages\Powershell\Test_All_File_Types\Favorites.vssx"> 
	$filepathCheck = $_ -Match 'rdf:about="(.*)"' 
	if ($filepathCheck) { $filepath =  $matches[1] } else { $filepath = "ERROR_NOMATCH"}  
	write-verbose ("filepath = $filepath ")
#	260 chars which is the maximum length for a path (file name and its directory route) 
#	7 = \ + .dcat + 1 
	$filenamelength = 260 - 7 - $TargetDir.length 
	write-verbose ("filenamelength = $filenamelength ")	
#	create a target file to contain contents 
	$filename = -join( $TargetDir, "\", (Convert-FilepathToFilename $filepath $filenamelength), ".dcat" )  
	
#	if the file exists, then use add-content to append the pipeline object 
	if( Test-Path $filename -PathType Leaf ) { 
		write-verbose ("filename $filename exists ")
	}else { 
#	crate a new file and set-content 	
		New-Item -Path $filename -ItemType "file" -Force 
		write-verbose ("filename $filename created ")
	}
#	add the content 
	Add-Content -Path $filename -Value $_ 	
#	alles klar message 
	write-verbose ("content added") 	
  }
}

#############################

function Add-HeaderFooter{
#	::	String -> String -> String -> [String] -> [String] 
#	::	header -> footer -> nameSuffix -> body -> [header,body,footer] 
#	This is a pipelinable function. 
#	typing rules only apply to input 
#	output type is defined by process {} output 

#	Usage : 
#	Import-Module DCAT -Force -Verbose 

#	test directory 
#	Set-Location "G:\IT Languages\Powershell\New" 
#	Set-Location "J:\Working Documents and Drafts\Corporate Services\ICT Services\Data Management\CKAN\Initial"

#	$header = 'A Header' 
#	$header = '<rdf:RDF xmlns:foaf="http://xmlns.com/foaf/0.1/" xmlns:owl="http://www.w3.org/2002/07/owl#"
#         xmlns:rdfs="http://www.w3.org/2000/01/rdf-schema#"
#         xmlns:rdf="http://www.w3.org/1999/02/22-rdf-syntax-ns#"
#         xmlns:dcat="http://www.w3.org/ns/dcat#"
#         xmlns:ods="http://open-data-standards.github.com/2012/01/open-data-standards#"
#         xmlns:dct="http://purl.org/dc/terms/">'

#	$footer = '	</rdf:RDF>'

#	$nameSuffix = '.RDF' 

#	Get-ChildItem -File | 
#	Add-HeaderFooter $header $footer $nameSuffix -Verbose 
#	Add-HeaderFooter $header $footer $nameSuffix 

  [cmdletbinding()]
  param(
    [Parameter(
      Position=0,
#	0 is first argument 
      Mandatory=$true, 
      ValueFromPipeline=$false 
    )] 
#	header is the new string to attach to front of file 
	[string]
    $header, 
    [Parameter(
      Position=1,
#	1 is second argument 
      Mandatory=$true, 
      ValueFromPipeline=$false 
    )] 
#	footer is the new string to attach to end of file 
	[string]
    $footer, 
    [Parameter(
      Position=2,
#	2 is third argument 
      Mandatory=$true, 
      ValueFromPipeline=$false 
    )] 
#	a mandatory nameSuffix is the new filename with the prepended string 
#	note this is not the filename, but a string suffix added to the BaseName for each new nameSuffix file 
	[string]
    $nameSuffix,
#	The $_. or $PSItem is the closure item, so it is always the last parm 
    [Parameter(
      Position=3,
#	3 is fourth argument 
      Mandatory=$true,
#	get item from pipeline 
      ValueFromPipeline=$true 
    )]
#	source is the original file object 
    [System.IO.FileInfo]
    $source 	
  )

  process{
    if(!$source.exists){
      write-error "$source does not exist" 
      return;
    }
    try{
#	check for already created source files  
		if($source.BaseName -like "*$($nameSuffix)*"){ 
			write-verbose ("1: skipping as $source is actually a nameSuffix "); 
#	no need to remove item, as it will be overwritten  
#			Remove-Item -Path $source -Force 
		}else { 
#	create a new nameSuffix inserting the nameSuffix string as a name suffix 
		$targetFilepath = $source.BaseName + $nameSuffix + $source.Extension 
		write-verbose ("2: targetFilepath = $targetFilepath ") 
#	create a new targetFilepath file, overwriting any current files 
		New-Item -Name $targetFilepath -Type file -Force 
#	read the source file 	
		$body = Get-Content -Path $source 
		write-verbose ("3: body = $body ") 
#	add the header to the targetFilepath 
		Set-Content -Path $targetFilepath -Value $header 
#		Add-Content -Path $targetFilepath -Value $header 
#	add the body to the targetFilepath 	
		Add-Content -Path $targetFilepath -Value $body 
#	add the footer to the targetFilepath 
		Add-Content -Path $targetFilepath -Value $footer 
#	alles klar message 
        write-verbose ("4: Function to create $targetFilepath from $source OK") 	
		}
    }
    catch{
      write-verbose $_.Exception.Message 
    }
    finally{
#	successful end 
      if($error.count -eq 0){
        write-verbose ("5: Function end OK") 
      }
      else{
#	failure  
        $error.clear();
        write-verbose ("6: Function to create $targetFilepath from $source FAILED") 
      }
    }
  }
}

#############################

#	It seems that all functions are exposed by default. 
#	Use the Export-ModuleMember -Function $ToExport cmdlet. 
#	Place this at the end after all functions. 

Export-ModuleMember -Function Convert-FileToDCATProps
Export-ModuleMember -Function Get-ExtensionToMimeType 
Export-ModuleMember -Function Convert-DCATPropsToDCATObj 
Export-ModuleMember -Function Convert-DCATObjToDCATXML
Export-ModuleMember -Function Convert-FilepathToFilename 
Export-ModuleMember -Function Add-SafeFile 
Export-ModuleMember -Function Add-HeaderFooter 
