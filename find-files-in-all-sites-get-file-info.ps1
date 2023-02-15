
#goal - find all instances of a file that exist in a sharepoint tenant.
# run using an ACCOUNT that has access to ALL sharePoint sites.

#looking for filename containing
#use something that works in sharepoint search box , single quote if you need phrase searching 
$searchTerm = "'new supplier' NEAR form FileType:xlsm"
$exportFileNameStart = "c:\temp\new-supplier-search-results"

#get all site collection URLs
#Set Parameter
$TenantSiteURL="https://YourTenantName.sharepoint.com"
 
#Connect to the Tenant site
Connect-PnPOnline $TenantSiteURL -interactive 
 
#then use search of all tenant to get results
$SearchResults = Submit-PnPSearchQuery -Query $searchTerm -All -SelectProperties ListItemID -SortList @{Created="Descending"}
$SearchResults.RowCount


$Results = @()
foreach($ResultRow in $SearchResults.ResultRows) 
{ 
    #Get All Properties from search results
    $Result = New-Object PSObject 
    $ResultRow.GetEnumerator()| ForEach-Object { $Result | Add-Member Noteproperty $_.Key $_.Value} 
    $Results+=$Result
}
#sort results so that sites are grouped, which makes next foreach work
$Results = $Results | sort-object -property IdentityWebId

$Results | out-gridview

#get a list of all sites containing the search results

$uniqueSitesWithResults = $Results | select-object -unique -property IdentityWebId
 $uniqueSitesWithResults.count
 
#>> following takes time, 15 files in 10sec or so or 400 files in 5 mins 
#results are ordered by site, so connect to site, itterate results, when site id changes, connect new site, itterate items ...
#	so will need to keep track of current siteid and next siteid
# search using get-item so can see filesize
# act on name / file size info e.g. if close match delete or rename or update

$prevIdentityWebURL = ""
$counter = 0

$allFileDetails = @()

forEach ($resultItem in $results) {
	
	#following assumes that /sites/ preceeds every site name 
	$thisIdentityWebURL = $TenantSiteURL+'/sites/'+$resultItem.path.split('/')[4]
	
	#following ensures you connect to each different site in turn
	if($thisIdentityWebURL -ne $prevIdentityWebURL) {
		#new site so connect
		Connect-PnPOnline $thisIdentityWebURL -interactive 
	}
		#this gets the info about this file 
		$resultItemThisFile = get-pnplistitem -list $resultItem.path.split('/')[5] -id $resultItem.listitemid 
		
		#Get All Properties from file result 
		$fileDetail = New-Object PSObject 
		$resultItemThisFile.FieldValues.GetEnumerator()| ForEach-Object { $fileDetail | Add-Member Noteproperty $_.Key $_.Value} 
		$allFileDetails+=$fileDetail
		
		
	
	
	write-host $counter 
	
	$prevIdentityWebURL = $thisIdentityWebURL
		$counter = ++$counter

}
 
$allFileDetails.count 

$fileName = $exportFileNameStart + "-AllFiles.csv"
$allFileDetails | export-csv -path $fileName -noTypeInformation

#now just get some peoples names 
#people who modified files in search results recently

$byPersonThenSiteFile = $allfiledetails | sort-object -property Modified_x0020_By,FileRef | select FileLeafRef, Modified, Modified_x0020_By, FileRef, fileDirRef

$byPersonThenSiteFile.count

$arrByPersonFoundfiles = @()
foreach($thing in $byPersonThenSiteFile) {
	#convert array to object
	
	$itemDetail = New-Object PSObject 
	$itemDetail | Add-Member "file name" $thing.FileLeafRef 
	$itemDetail | Add-Member "modified" $thing.Modified
	
	$itemDetail | Add-Member "modified by" $thing.modified_x0020_By
	$thisSiteUrl = $TenantSiteURL+'/sites/'+$thing.FileRef.split('/')[2] 
	$itemDetail | Add-Member "Site url" $thisSiteUrl
	$thisfileFolderUrl  = $TenantSiteURL+'/sites/'+$thing.fileDirRef
	$itemDetail | Add-Member "Folder file in" $thisfileFolderUrl
	
			$arrByPersonFoundfiles+=$itemDetail

}

$arrByPersonFoundfiles.count

$fileName = $exportFileNameStart + "-ContactPeopleFiles.csv"

# from the CSV you should be able to build some links and tables for each person so it is easy for them to get to the files.
$arrByPersonFoundfiles | export-csv -path $fileName -noTypeInformation




