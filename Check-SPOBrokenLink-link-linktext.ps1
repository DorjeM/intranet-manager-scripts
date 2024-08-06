<#
		.SYNOPSIS
		This script tries to get the links in SharePoint online pages, and test each of those links to see if they are work. A 404 link checker in essence.

		.DESCRIPTION
  		This script requires the PNP powershell module to run.
    
		To run this script and get a list of all the links in all the pages in all the sites queried do the following:
		Open a PowerShell window
		Navigate to the folder you saved this PS1 file into e.g.   CD c:\temp
		Then run the script by typing in     .\Check-SPOBrokenLink-link-linktext.ps1
		You'll be prompted for all the variables required.
		For each SharePoint online site that has it's web page links tested, and every site that is linked to you will see a popup asking you to logon to that site.
		Finally the script will tell you where the output was saved once the script is completed.
		
		This function was originally inspired by https://pnp.github.io/script-samples/spo-modern-page-url-report/README.html?tabs=pnpps  but spo-moder-page-url-report didn't take into account enough of the SharePoint online nuances. So this function was developed.
		
		Which sites to check for broken links can be set using a list of site URLs seperated by a comma, or by using the getSiteListFromSPAdmin function to get ALL the site URLs in your tenant.
		
		The script then connects to each site in turn, and every ASPX file in that site. From the ASPX files the CanvasContent1 property is extracted. From that all the links are extracted. Each link is put into an array that holds the page and the link in that page.
		NOTE the mega menu links and links it the "Quick links" web part can't be simply accessed so they are not checked to see if they are ok.
		
		The the list of all links is reduced to just the unique links. This means we aren't testing the same link multiple times, speeding up the script.
		Each link is then tested
		The test results for each unique link are then used to update every link in the list of all links.
		
		Finally the list of all links, and the results of the test for each link are exported as a CSV file.
		
		As with anything like this, the script does the bulk of the work but there will still be links which need to be manually checkeed.
		This CSV can simply be turned into a report you can work through these manual checks.
		I suggest you open the CSV, then save it as an XLSX file, then update the I colum "ManualTestLinkUrl" with the following formula
			=HYPERLINK(E2,E2)
			If you get a #VALUE! error in this column it is likely due to the linkURL beeing longer than 256 characters (an excel limitation)
		
		That way you can sort by the LinkURL, click on the I colum "ManualTestLinkUrl" link and test those links which the script couldn't test or got inconclusive results from. 
		Column H "scriptStatus" has the following values
			fileDoesntExist, or no security access
			fileExists
			Forbidden
			libraryExists
			mail link, not checked
			manual check, colon url
			manual check, within page link
			OK
			UnsureLinkMayRedirect
			(blank)
		The following values definately mean that link is working:
			fileExists
			libraryExists
			OK
		The other values in column H you will have to check manually.
		Finaly update column K "HyperlinkToPage" with the following formula in J2
			=HYPERLINK(CONCAT(A2,"/sitepages/",B2),c2)
		This gives you a quick way to go to any pages that have links which need fixing
		In my case (blank) column G "scriptStatus" items were links to the root of sharepoint sites. 90% were sites I know exist so quick to check and update.
		
		As an example I ran this across 2 SharePoint communication sites and found 820 links. 468 unique links.
			314 needed manual checking (but due to duplication only needed to click on some of these)
			152 links were not working after manual testing 
			But they existed only on 40 pages
		It took me about 40 minutes to run the script and check the links, document them in the spreadsheet, send them to the site owners to fix up.
			
		
		Thanks to 
		Alister Air for asking for something to do this in the StepTwo forum in 2021, and then Alise Croft and Michael Hutchens for inspiring it's improvement in 2024. 
			
		.SETUP 
		To prepare this script to run you need to set the following variables		
			$finalReportCSV , choose a path where the CSV will be created 
			$MyTenantURL , so that the system knows how to build up links 
			$Library , you need to match this to the language your tenant was setup with 
			$showConsoleUpdates , set this to true or false. Use true the first few times you run the script 
#>


# clear everything in this instance of the interactive PowerShell window
cls

#>>>>>The following are variables you have to set.

$ArrayURLStatus = @()

#give a path to export the checked link information to, a unique file name will be generated for you
#$finalReportCSV = "c:\temp"
$finalReportCSV = read-host -Prompt "Enter the path to the folder the output CSV will be saved to e.g. c:\temp"

#$MyTenantURL = "https://MyTenant.sharepoint.com"
$MyTenantURL= read-host -Prompt "Enter the url of your SharePoint tenant e.g. https://myTenant.sharepoint.com"

#$Library = "Site Pages" # depends of the site's language
$Library = read-host -Prompt "Enter the name (not the url) of the library where ASPX files are saved in every site e.g. Site Pages"
#set this to $true for updates as the script runs to be shown to you. 
$screenUpdates = read-host -Prompt "Enter Yes to see script progress in the PowerShell console window. No will only show minimal info."

if($screenUpdates.ToLower() -eq "yes") {
	$showConsoleUpdates = $true
	} Else {
	$showConsoleUpdates = $false
}

#>> Comma delimited list of site URLs you want to test the page links in 
#  e.g. 
$siteList = read-host -Prompt "Enter a list of site URLs you want to check for broken links. Delimited by commas e.g. https://contoso.sharepoint.com/sites/intranet,https://contoso.sharepoint.com/sites/safetyintranet"

#$TenantSites = "https://contoso.sharepoint.com/sites/intranet,https://contoso.sharepoint.com/sites/safetyintranet" -split ","
$TenantSites = $siteList -split ","


#>> NOTE   You can get a list of all the Communication sites in your tenant from the user interface
#Go to the SharePoint admin site then click > active sites > filter communication sites > export > generate a list of urls speperated by a comma
# or you can use the function you'll find below getSiteListFromSPAdmin

#>>> END of variables you should set

#>>> Other variables required, do not change 

#the following regular expression works best
$urlPatern = '\shref=["](.*?)["].*?>(.*?)<\/a>' 
#returns group 1 (which is href) and group 2 (which is linktext) for all types of links in the Canvas page.

$i = 0
$j = 0


#this function is here to help you get a list of all sites, you need to modify the script and call it to take advantage of it.
Function getSiteListFromSPAdmin{
<#
		.SYNOPSIS
		Get a list of sites from the SharePoint admin site if you have an account with access 

		.DESCRIPTION
		
		.PARAMETER spAdminURL
		The URL to be checked

		.EXAMPLE
		getSiteListFromSPAdmin -spAdminURL https://contoso-admin.sharepoint.com


#>
	
    Param
    (
        [Parameter(Mandatory=$true)]$spAdminURL
    )

	#Connect to the Tenant site, 
	write-host "From the popup choose an account that has access to the SPAdmin site"
	start-sleep -seconds 10 
	Connect-PnPOnline $spAdminURL -Interactive
	 
	#sharepoint online pnp powershell get all sites
	Get-PnPTenantSite
	# if you get the error "Attempted to perform an unauthorized operation" then the account you've authenticated with doesn't have enough rights to connect to the sharepoint-admin site.
	
	#if you want to filter for one site then use   
	$siteUrls = Get-PnPTenantSite | ? {($_.Template -eq "SitePagePublishing#0") -and ($_.URL -eq "$MyTenantURL/sites/hr")} | Select -ExpandProperty URL 
	
	#if you want to filter for some sites starting with same url then
	#		$siteUrls = Get-PnPTenantSite | ? {($_.Template -eq "SitePagePublishing#0") -and ($_.URL.contains("$MyTenantURL/sites/intra"))} | Select -ExpandProperty URL
	#Read more: https://www.sharepointdiary.com/2016/02/get-all-site-collections-in-sharepoint-online-using-powershell.html#ixzz8eJLxvGOX
	
	if($showConsoleUpdates) { $siteUrls.count }
	
	Return $siteUrls

	#disconnect pnp-online so that connection to each of the subsites can use a different user account.
	disconnect-pnponline
}

#>> Uncomment the following step to use the function to get ALL the URLs
#>> beware this can be a very long process. Test multiple times with the comma delimited list of URLs first 
# $TenantSites = getSiteListFromSPAdmin -spAdminURL "https://contoso-admin.sharepoint.com"
# $TenantSites.count



Write-host $TenantSites.count " TenantSites will be checked"


Function Test-URL{
<#
		.SYNOPSIS
		Test if URL is reachable

		.DESCRIPTION
		This function was originally inspired by https://pnp.github.io/script-samples/spo-modern-page-url-report/README.html?tabs=pnpps  but spo-moder-page-url-report didn't take into account enough of the SharePoint online nuances. So this function was developed.
		
		Basic approach
		This function processes a URL from a SharePoint online page. 
		If the link is not a SharePoint one it tests the link using invoke-webrequest.
		Then the link is tested to see if it is in the same site as the current logged in user context. 
			If it is then it tries to prove the site, or the library, or the list, or the folder, or the file exists.
		
		Anti scraping software 
		Some sites have anti scraping software so that you need to send http headers, it is quite complex, so if the site is external to sharepoint and it flags as the link is broken it may just be that site has some Anti scraping solution in place. 
	
		Invoke-Webrequest 
		invoke-webrequest is hard to do with Microsoft sharepoint credentials without useing an Enterprise App ID and secret. This solution was intended for a technically compentent SharePoint online admin to run without resorting to IT support. So other SharePoint specific tests are done, only using invoke-webrequest when linkType = website 
		
		Thanks to 
		Alister Air for asking for something to do this in the StepTwo forum in 2021, and then Alise Croft and Michael Hutchens for inspiring it's improvement 
			
		
		.PARAMETER URL
		The URL to be checked

		.PARAMETER linkType
		Is the URL being checked same sharepoint site samesite. Or sametenant i.e. a sharepoint site in the same tenant. Or website i.e. an external web address nothing to do with SharePoint. If no value submitted then assume it is weblink.

		.EXAMPLE
		Test-URL -URL https://MyTenant.sharepoint.com/MyCommsSite/sitepages/home.aspx -linkType 'samesite'
		test-url  $url 'samesite'
		test-url  -URL $url -linkType 'samesite'
#>
	
    Param
    (
        [Parameter(Mandatory=$true)]$URL,
		[Parameter(Mandatory=$false)]$linkType,
		[Parameter(Mandatory=$false)]$showInfo
    )

	
	$Status = ""
	
	if([string]::IsNullOrEmpty($linkType)){
		$linkType = "weblink"
		}

	if($linkType -eq "weblink") {
		
		Try {
			$SiteStatus = Invoke-WebRequest $URL
			if($SiteStatus.StatusCode -eq "200"){
					if($showInfo) { write-host "Link OK" -f green }
					$Status = "OK"
				}			
        }
		catch{
        if($showInfo) { write-host "Dead Link"-f red }
        $Status = $error[0].Exception.StatusCode
		}
	
	}	
	
	#figure out the type of test to do on this SharePoint URL 
	if(($linkType -eq "samesite") -OR ($linkType -eq "sametenant")) {
		
		#change full $url to relative URL
		$relUrl = $url.ToLower().replace((get-pnpweb).url.ToLower(),'')
		
		#what type of thing is the URL
		switch -regex ($relUrl) {
			"(?i)\/Forms\/AllItems.aspx$" {
					if($showInfo) { write-host "The url is an SPonline library listing" }
					$endPoint = "splibrary"
					} #assume this is a SharePoint library 
			"(?i)\/lists\/" {
					if($showInfo) { write-host "The url is an SPonline list" }
					$endPoint = "splist"
					} #assume this is a SharePoint list 
			default {
				if($showInfo) { write-host "Assume this is a file or webpage" }
				$endPoint = "file"
				} #assume this is a SharePoint file or page
		}
		
		if($relUrl -eq (get-pnpsite).url) {
			if($showInfo) { write-host "The url is a SharePoint site" }
			$endPoint = "spsite" #assume this is a SharePoint site
		}
		
		#test depending on endPoint AND linkType
		switch ($endPoint) {
			"spsite" {
				$Status = "spSiteExists"
			}

			
			"splibrary" { 
				$relUrl = $relUrl.replace('/forms/allitems.aspx','')
				$relUrl = $relUrl.replace('%20',' ')
				
				if(get-pnpfolder -url $relUrl -ErrorAction SilentlyContinue) {
					$Status = "libraryExists"
				} else { 
					$Status = "libraryDoesntExist" 
				}
				
			}
			
			"splist" { 
				if(Get-PnPList -Identity $relUrl -ErrorAction SilentlyContinue) {
					$Status = "listExists"
				} else { 
					$Status = "listDoesntExist" 
				}
			
			} 
			
			default { 
				if(Get-PnPFile -Url $relUrl -ErrorAction SilentlyContinue) {
					$Status = "fileExists"
				} else { 
					$Status = "fileDoesntExist" 
					
				}
			
			}
			
		}
	
	}


	#if doesntexist then check if it is a folder
	if($Status -eq 'fileDoesntExist' -or $Status -eq 'libraryDoesntExist' -or $Status -eq 'listDoesntExist') {
		$relUrl = $relUrl.replace('%20',' ')
		$relUrl = (get-pnpweb).ServerRelativeUrl+$relUrl
		if(Get-PnPFolder -Url $relUrl -ErrorAction SilentlyContinue) {
			$Status = "folderExists"
			} 
	}
	
	#check for wopi urls which may redirect
	if($Status -eq 'fileDoesntExist') {
		if($relUrl -match '_layouts|aspx\?id|\/:.:\/') {
			$Status = 'UnsureLinkMayRedirect'
		} else {
			$Status = $Status+", or no security access"
		}
	}


    Return $Status
}



# START
##################################################################################################

# for each site get the links in every page 
# using the CanvasContent1 property of each page 
foreach($TenantSite in $TenantSites){
    Connect-PnPOnline $TenantSite -Interactive 
	
	$thisSiteHomePageUrl = get-pnphomepage
	#will be a value like sitepages/Home.aspx 
	
	$i = $i + 1
    
    if($showConsoleUpdates) { write-host "`nTenantSite: $TenantSite" }
    
    # get the contents of all the pages of the site
    $TenantSitePages = Get-PnPListItem -List $library -Fields CanvasContent1,Title,FileLeafRef | Where-Object { $_["FileLeafRef"] -like "*.aspx" }
	
    if($showConsoleUpdates) { $TenantSitePages.count }
	
    # for each page in this site
    foreach($TenantSitePage in $TenantSitePages){
		
		
		
        $j = $j + 1
 
	# modify the content to be more readable by a human being
		$TenantSitePageContentHumans = [System.Web.HttpUtility]::UrlDecode($TenantSitePage.FieldValues.CanvasContent1)
		$TenantSitePageContentHumans = $TenantSitePageContentHumans.replace("&#58;", ":")
        
	# detect all URLs
        $TenantSitePageContentHumanURLs = $TenantSitePageContentHumans | select-string -Pattern $URLPatern -AllMatches
        if($showConsoleUpdates) { $TenantSitePageContentHumanURLs.Matches.count }
		# NOTE often due to the way SharePoint manages links there will be 2 of the same link in the CanvasContent1, for each actual url in the sharepoint page users edit.

	# for each detected link
        foreach($TenantSitePageContentHumanURL in $TenantSitePageContentHumanURLs.Matches ){
            $thisLinkText = ""
			$thisLinkUrl = $TenantSitePageContentHumanURL.Groups[1].Value
			$thisLinkText = $TenantSitePageContentHumanURL.Groups[2].Value #can be blank e.g. for hero image links
			$urlParts = ""
			$linkURLBaseSite = ""
			
			
			#if the first letter of any item is / then append in front $MyTenantURL
			#	this changes /sites/xxxx  to  https://mytenant.sharepoint.com/sites/xxxx
				if($thisLinkUrl.Substring(0, [Math]::Min($thisLinkUrl.Length, 1)) -eq "/")
				{
					$thisLinkUrl = $MyTenantURL+$thisLinkUrl
				}
			
			#what if there is no aspx in the URL i.e. a link to the default or home page of site.
			#	then postpend the $thisSiteHomePageUrl variable 
				if($thisLinkUrl -contains $TenantSite) {
					if($thisLinkUrl -notcontains ".aspx") {
						$thisLinkUrl = $thisLinkUrl+"/"+$thisSiteHomePageUrl
						#write-host "yyy: $thisLinkUrl"
					}
				}
			
			#if the link goes to a sharepoint site other than one that has been connected to, the test-url function says the aspx file doesn't eixst.
			#	So if the site is different to the currently connected site then connect to the site before testing the link.
			#	by ordering the links being tested we ensure each site is only connected to ONCE, because site connections slow things down
			
			if($showConsoleUpdates) { write-host "$thisLinkUrl" }
			
			$linkType = ""
			
			#figure out what linkType this URL is
			switch -wildcard ($thisLinkUrl) {
				#note the order is important
				"$MyTenantURL*" {
					if($showConsoleUpdates) { write-host "The url is from this tenant " }
					$linkType = "sametenant"
					} 
				"$TenantSite*" {
					if($showConsoleUpdates) { write-host "The url is from this site " }
					$linkType = "samesite"
					} 
				"mailto*"  {
					if($showConsoleUpdates) { write-host "The url is from this site " }
					$linkType = "mailto"
					} 
				default {
					if($showConsoleUpdates) { write-host "The url a weblink" }
					$linkType = "weblink"
					} 
			}
			
			if($showConsoleUpdates) { write-host "LinkType:"$linkType }
			
			#figure out the SITE that the linkURL is in so that we can order the results, reducing the number of times each site needs to be connected to for checking 
			#reset site url for special type of link containing :/r the r indicating relative url e.g. https://MyTenant.sharepoint.com/:b:/r/sites/myCommSite/
			
			$urlParts = ""
			$linkURLBaseSite = ""
			$urlParts = $thisLinkUrl.split('/')
			switch -regex ($thisLinkUrl) {
				"\/\:.\:\/r" {
					$linkURLBaseSite = "https://"+$urlparts[2]+"/"+$urlparts[5]+"/"+$urlparts[6] }
				"\/\:.\:\/s" {
					$linkURLBaseSite = "https://"+$urlparts[2]+"/sites/"+$urlparts[5] }
				default {
					$linkURLBaseSite = "https://"+$urlparts[2]+"/"+$urlparts[3]+"/"+$urlparts[4] }
			}
			
			
			# test URL
            $URLStatus = ""
			
			$ObjURLStatus = [PSCustomObject]@{
						Site = $TenantSite
						PageURL = $TenantSitePage.FieldValues.FileLeafRef
						PageTitle = $TenantSitePage.FieldValues.Title
						MatchID = $TenantSitePageContentHumanURL.Index #not sure what this actually is
						linkURL = $thisLinkUrl
						linkUrlText = $thisLinkText
						linkType = $linkType
						linkURLBaseSite = $linkURLBaseSite
						Status = $URLStatus
					}
			 $ArrayURLStatus += $ObjURLStatus  
        }
    }
	#close connection to this site so can move to next 
	disconnect-pnponline 
}

#Number of links found in all the pages in all the sites tested
write-host -ForegroundColor Magenta "We found: "$ArrayURLStatus.count" links, in: "$TenantSites.count" sites"

#the arrray now contains all the items we need to report on
#but we need to only test the Unique linkUrls
#this is more efficient and faster than checking EVERY link 

#Get the unique links and their type
$uniqueLinkURLs = $ArrayURLStatus | select site, linkURLBaseSite, linkURL, linkType | Sort-Object linkURLBaseSite, linkURL, linkType -unique
#sort by linkType then linkUrl 
#	will put samesite first, then sametenant ordered by URL, then weblink
#	this means the next section can connect to each site just once if it needs to, again speeding things up.

$uniqueLinkURLs = $uniqueLinkURLs | Sort-Object linkURLBaseSite, linkURL, linkType 
write-host -ForegroundColor Green "We found: "$uniqueLinkURLs.count" unique links, that will be tested."

#now test each of these URLS and update all the rows in ArrayURLStatus with the Status
#$uniqueLinkURLs = $uniqueLinkURLs | select-object -first 20
#$uniqueLinkURLs | out-gridview

$arrUniqueLinkURLsChecked = @()
$previousSite = ""
$link = ""

#>>>>> need more work on next fornext loop , to take account of linkURLBaseSite and how it interacts with linkType 

foreach($link in $uniqueLinkURLs) {
	
	$urlStatus = ""
	$urlParts = ""
	$urlParts = $link.linkUrl.split('/')
	$thisSiteurl = ""
	
	#reset site url for special type of link containing :/r the r indicating relative url e.g. https://MyTenant.sharepoint.com/:b:/r/sites/MyCommSite/
	switch -regex ($link.LinkURL) {
		"\/\:.\:\/r" {
			$thisSiteurl = "https://"+$urlparts[2]+"/"+$urlparts[5]+"/"+$urlparts[6] }
		"\/\:.\:\/s" {
			$thisSiteurl = "https://"+$urlparts[2]+"/sites/"+$urlparts[5] }
		default {
			$thisSiteurl = "https://"+$urlparts[2]+"/"+$urlparts[3]+"/"+$urlparts[4] }
	}
	
	if($link.linkType -eq 'mailto') {
		$urlStatus = "mail link, not checked"
	}
	
	if($link.LinkUrl -match 'sourcedoc=' -and $link.linkType -ne 'weblink') {
		#we know test-url function can't figure these out so
		#	looks for urls with _layouts OR aspx?id OR /: some letter :/ in them
		if($showConsoleUpdates) { write-host "complex url with colons" }
		$urlStatus = "manual check, colon in url"
		
	} 	
	
	
	#do connection stuff 
	if(($previousSite -ne $thisSiteurl) -and ($link.linkType -eq 'samesite' -or $link.linkType -eq 'sametenant') -and ($urlStatus -eq '')) {
			#attempt a pnp connection
			try {disconnect-pnponline} catch {write-host "No connection to disconnect"}
			write-host -ForegroundColor Magenta "For link testing Connecting to: "$thisSiteurl
			write-host -ForegroundColor White "NOTE a get-pnpfile or get-pnpsite error may be displayed, it won't stop the script running for all the other linkUrls"
			Connect-PnPOnline $thisSiteurl -interactive
			#sleep 2 secs to avoid errors calling get-pnpsite too quickly
			sleep 2
			$currentSite = (get-pnpsite).url
		}
		
		
	#links in page using # in the linkURL come back from TEST-URL as doesn't exist
	if($link.linkURL -match '#') {
		$urlStatus = "manual check, within page link"
	}
	
	#check one of the url types we know exist in sharepoint that has a specific structure we can work around
	if(($link.LinkURL -match 'AllItems\.aspx\?id=') -and ($link.linkType -ne 'weblink') -and ($urlStatus -eq '')){

		If ($link.Type -eq 'sametenant'){
			#ignore sametenant this loop has connected to the current link's site so should Test-URL function will work with samesite setting
			$urlStatus = test-url -URL (($link.LinkURL -split "id=")[1] -split "&amp")[0] -linkType 'samesite' -showInfo $showConsoleUpdates
			} else {
			#test most links
			$urlStatus = test-URL -URL (($link.LinkURL -split "id=")[1] -split "&amp")[0] -linkType $link.linkType -showInfo $showConsoleUpdates
			}

	}
	
	#check another SharePoint link that has a specific structure, contains :/r/ , can't check :/s/ because they are all GUID based 
	if(($link.LinkURL -match '\/\:.\:\/r') -and ($link.linkType -ne 'weblink') -and ($urlStatus -eq '')){
		If ($link.Type -eq 'sametenant'){
			$urlStatus = test-url -URL (($link.LinkURL -split '\/\:.\:\/.','')[1] -split "\?")[0] -linkType 'samesite' -showInfo $showConsoleUpdates
		} else {
			$urlStatus = test-url -URL (($link.LinkURL -split '\/\:.\:\/.','')[1] -split "\?")[0] -linkType $link.linkType -showInfo $showConsoleUpdates
		}
		
	}
	
	
	#do nothing if the urlStatus is already predefined 
	if ($urlStatus.length -eq 0) {
	
		If ($link.Type -eq 'sametenant'){
			#ignore sametenant this loop has connected to the current link's site so should Test-URL function will work with samesite setting
			$urlStatus = test-url -URL $link.linkUrl -linkType 'samesite' -showInfo $showConsoleUpdates
			} else {
			#test most links
			$urlStatus = test-URL -URL $link.linkUrl -linkType $link.linkType -showInfo $showConsoleUpdates
			}
			
	}
	

	
	
	$ObjThisLink = [PSCustomObject]@{
						linkURL = $link.linkUrl
						linkType = $link.linkType
						Status = $urlStatus
					}
	$arrUniqueLinkURLsChecked += $ObjThisLink
	
	$previousSite = ""
	$previousSite = $thisSiteUrl 
	$urlStatus = ""		 
}


write-host -ForegroundColor White $arrUniqueLinkURLsChecked.count" unique linkUrls have been checked"

#Improvement join the unique checked links $arrUniqueLinkURLsChecked  with the full list of page links $ArrayURLStatus
#	the join should be done on 
#	$arrUniqueLinkURLsChecked linkURL property 
#	and
#	$ArrayURLStatus linkURL
#
#	I couldn't figure out how to merge the arrays so I just go through one of them again, which is quite quick.
#	
#	Go through ever item in $ArrayURLStatus and update the Status value with the value from the matching $arrUniqueLinkURLsChecked 

$arrAllPageLinksChecked = @()

foreach($item in $ArrayURLStatus) {

	if($showConsoleUpdates) { write-host -ForegroundColor Blue "Updating status of: "$item.linkUrl }
			$thisItemInfo = ""
			
			$thisItemInfo = [PSCustomObject]@{
						Site = $item.Site
						PageURL = $item.PageURL
						PageTitle = $item.PageTitle
						MatchID = $item.MatchID
						linkURL = $item.linkUrl
						linkUrlText = $item.linkUrlText
						linkType = $item.linkType
						ScriptStatus = ($arrUniqueLinkURLsChecked | where-object {$_.linkURL -eq $item.linkUrl}).Status
						ManualTestLinkUrl = ""
						ManualLinkUrlCheckStatus = ""
						HyperlinkToPage = ""
					}
			 $arrAllPageLinksChecked += $thisItemInfo


}	

$arrAllPageLinksChecked.count

#$arrAllPageLinksChecked | out-gridview
$pathToUse = "$finalReportCSV\SPOlinks-$(get-date -Format yyyy-MM-dd-hh-mm).csv"
$arrAllPageLinksChecked | export-csv -path $pathToUse -NoTypeInformation

Write-host "The link checker results are available here: $pathToUse"
Write-host "Add this formula to I2 =HYPERLINK(E2,E2)"
write-host "Add this formula to K2 "
$formula2 = "=HYPERLINK(CONCAT(A2,""/"
$formula2 = $formula2+$Library.replace(' ','')
$formula2 = $formula2+"/"",B2),CONCAT(C2,"" : find linkURLText col F and fix linkURL col E""))"
write-host $formula2
Write-host "Next, filter by column H to exclude all files actually found"
Write-host "Test the remaining files by clicking the links in column I"
Write-host "Update column J with what you discover, e.g. if the link doesn't work put 404, if it does work put OK"
write-host "Now filter again using column J, showing anything not = OK"
write-host "Now click each of the links in column K, and edit those files searching for the link text or link url, and changing it to a link that works."
write-host -ForegroundColor Magenta ">>> Good work! <<<"
