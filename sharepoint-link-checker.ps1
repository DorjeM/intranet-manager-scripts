<#
        .SYNOPSIS
			This script finds broken links in SharePoint online Modern pages, it is simplistic. A better solution is Check-SPOBrokenLink-link-linktext.ps1

        .DESCRIPTION
			This script was created to be run in an interactive powershell window.
			
			This script looks at SharePoint Online Modern pages
			Specifically it looks at the the page's  "FieldValues.CanvasContent1"  property value.
			NOTE this means it doesn't see the megamenu links OR some links within web parts (depending on how they are renedered)
			
			It uses regular expressions to find hyperlinks in this html.
			If figures out all the unique links that exist in the pages.
			It tests the unique links using Invoke-WebRequest
			It then reports back for every page that contains a link which doesn't work (or it can't understand) the following:
				page URLs so you can use excel or other tool to make a list of clickable links to open that page.
				the link href url
				the linktext (so you can find it in the page)

			During my first run the script found 4000 links in 641 site pages, found 350 broken links in 240 pages.
			Over 3 or 4 hours I was able to fix the links, delete the page that had the bad link (because it wasn't required) or ignore the page e.g. old news items linking to external sites 

			The script can check about 40 links per minute, I've optimised this as much as I can.
			
			NOTE you may have to alter the >> Make real URLs to test << section to match the types of links in your environment. 
			
			IMPROVEMENTS 
			The linkText of Modern Quick Links web part, isn't captured by my reglar expression.
			Links to videos on https://web.microsoftstream.com/ show up as broken links because stream redirects the user to the video so many times it works in the browser but not in my code.
			Links to systems which require browser authentication e.g. our capex system, show up as broken links because they returen 401 unauthorised errors.
			An improvement may be to return the HTTP status codes with the other data to the user.
			
			Thanks to 
			Alister Air for asking for something to do this in the StepTwo forum. It helped me clean up some very poor links on my own intranet. I hope it helps your intranet users.
			
        .PARAMETER
			There aren't any as such but you do need to update some variables
				
				$SiteURL   e.g.  "https://TenantName.sharepoint.com/sites/sitename"
				
				This script checks the "Site Pages" library. 
				
        .EXAMPLE
			Not really applicable. Open a powershell window, paste in the code and try it out.
			
        .INPUTS
			None
			
        .OUTPUTS
			The script sends you back a table using the out-gridview function 
			
        .NOTES
            Created by   : Dorje McKinnon 
			Created on the back of : Questions from 
            Date Coded   : 30 and 31 Mar 2022
			
			References   :
				https://blog.p-difm.com/sharepoint-online-check-for-broken-links/
					note last comment $TenantSitePage.FieldValues.CanvasContent1
				and https://github.com/Sayannara/Check-SPOBrokenLink
				and https://sharepains.com/2017/10/05/office-365-check-your-site-for-broken-links-in-sharepoint-online-part-1/
				and https://regex101.com/ to test the regex patterns (this took the longest time because the CANVAS links aren't like normal HTML links.
				and  https://petri.com/testing-uris-urls-powershell/ used for speeding up the invoke-webrequest
			
        .LINK
            Home
#>




#Set Variables
#>>> you will be prompted for for these <<<


$SiteURL = Read-Host -Prompt 'Enter your site relative URL e.g.  https://tenantname.sharepoint.com/sites/sitename'

#>>> end of anything you need to put in <<<<


#set other variables

#check no trailing slash
$SiteURL = $SiteURL.TrimEnd('/')

[regex]$reg1 = "http.*?.com"
$tenantUrl = $reg1.match($siteurl).value

$LibraryName = "Site Pages"

# $URLPatern = "(https|http)://.+?(`"|')"
# $URLPattern2 = "href\s*=\s*(?:[""'](?<1>[^""']*)[""']|(?<1>[^>\s]+))"
# $URLPattern3 = '(href=\")(.+?)(")"'
# $URLPattern4 = '(\shref=")(.*?)(\")'
 
 $URLPattern5 = '\shref="(.*?)".*?>(.*?)<\/a>' #returns group 1 (which is href) and group 2 (which is linktext) for all types of links in the Canvas page.
 
 

 
#functions used later on

Function Encode-HumanReadability{
 <#
		.SYNOPSIS
		Changes the encoding values for better human readability
		
		.PARAMETER SiteURL
		The content to be decoded
		.EXAMPLE
		$Content = Encode-HumanReadability -ContentRaw $MyText
#>

		
    Param
    (
        [Parameter(Mandatory=$true)]$ContentRaw
    )


    $ContentHuman = $ContentRaw.replace("&#58;", ":")
    $ContentHuman = $ContentHuman.replace("&quot;", '"')
    $ContentHuman = $ContentHuman.replace("&#123;", "<")
    $ContentHuman = $ContentHuman.replace("&#125;", ">")
    $ContentHuman = $ContentHuman.replace("&gt;", ">")
    $ContentHuman = $ContentHuman.replace("&lt;", "<")

    Return $ContentHuman
}


Function Test-URL{
<#
		.SYNOPSIS
		Test if URL is reachable
		
		.PARAMETER URL
		The URL to be checked
		.EXAMPLE
		$URLStatus = Test-URL -URL https://MyTenant.sharepoint.com
#>
	
    Param
    (
        [Parameter(Mandatory=$true)]$URL
    )

    # we assume that there are NO links to file shares, that all links are http links 
    try{
        
        
			#write-host "Testing a URL"
			#write-host $URL
		# if this step goes wrong, try using Invoke-WebRequest $URL -UseBasicParsing
            #following stops the progress bar
			$progressPreference = 'silentlyContinue'
			#was a bit slow as it go whole page
			#$SiteStatus = Invoke-WebRequest $URL
			#following may work faster
			#ref for this is https://petri.com/testing-uris-urls-powershell/
			$SiteStatus = Invoke-WebRequest -DisableKeepAlive -UseBasicParsing -Method head -uri $URL
			#following enables the progress bar, in case other cmdlets need it 
			$progressPreference = 'Continue'

            if($SiteStatus.StatusCode -eq "200"){
                #write-host "Link OK" -f green
                $Status = "OK"
            }
        
    
    }
    catch{
        write-host "Dead Link"-f red 
        write-host $Error[0].Exception
        write-host $Error[0] -f red 
		if($error[0].Exception.Response.StatusCode -ne $null) {
			$status = $error[0].Exception.Response.StatusCode
		} else {
        $Status = "NOK"
		}
    }

    Return $Status
}

#>>> actual code that does stuff
 
 
 
    
 #Connect to the site using PNP Online
Connect-PnPOnline -Url $SiteURL -interactive

$ThisSitePages = Get-PnPListItem -List $Libraryname

$allSitePageTextBoxLinks = @()

#get the links out of every page.

foreach($page in $ThisSitePages) {
	
	$thisPageUrl = $page.FieldValues.FileRef
	$thisPageContent = $page.FieldValues.CanvasContent1

	$thisPageLinks = @()
	
	if($thisPageContent.length -lt 1) {
		#do nothing
	} else {
		$thisPageContentHumanReadable = Encode-HumanReadability $thisPageContent

		#$thisPageContentHumanReadableURLs = $thisPageContentHumanReadable | select-string -Pattern $URLPattern3 -AllMatches
		
		[regex]$reg = $URLPattern5
		$thisPageContentHumanReadableURLs = $reg.match($thisPageContentHumanReadable)

		while($thisPageContentHumanReadableURLs.success) {
				$thisLink = $thisPageContentHumanReadableURLs.Groups[1].Value
				$thisLinkText = $thisPageContentHumanReadableURLs.Groups[2].Value
				$objThisLinkInfo = [PSCustomObject]@{
					"url"=$thisPageUrl
					"link"=($thisLink) #href
					"linktext"=($thisLinkText) #linktext 
				}
				$thisPageLinks += $objThisLinkInfo
				$thisPageContentHumanReadableURLs = $thisPageContentHumanReadableURLs.nextMatch()
		}

		$allSitePageTextBoxLinks += $thisPageLinks
	}
	
	
}



Write-host "Total Links within the Canvas part of Modern ASPX pages within the "$Libraryname" i.e. excludes Mega menu links"
$allSitePageTextBoxLinks.count

<# testing code 
#filter just one page worth of links

$justOnePage = $allSitePageTextBoxLinks | ?{$_.url -like '*page-I-want-to-test-on-its-own.aspx'}

$allSitePageTextBoxLinks = $justOnePage

#>


#remove mailto links which we can't check
#could expand this to other types of link 
$allSitePageTextBoxLinksNotMailto = $allSitePageTextBoxLinks | where-object link -notlike mailto*

#get unique URLs and test each one, to find which one's dont work.
#then show which pages have links to the failing urls

#ref https://stackoverflow.com/questions/1391853/removing-duplicate-values-from-a-powershell-array comment by Omzig

Write-Host "Get list of unique links, to test (faster script speed not having to test the same link twice). Approx 40 tests/min in my environment" -ForegroundColor Yellow
$siteUniqueLinks = $allSitePageTextBoxLinksNotMailto | Sort-Object -Property link -Unique 
$testingCounter = 0
$testingTotal = $siteUniqueLinks.count 
$startTestingDateTime = get-date

$siteUniqueLinksNotOk = @()

#for each link in $siteUniqueLinks
#>> Make real URLs to test <<
#make sure it is a real url, then test it with test-link function
#In my environment there wer about 6 unique characters at the start of each hyperlink
#	this is because some links went to /sites/sitename/sitepages/pagename.aspx and some links went to https://tenantname.sharepoint.com/sites/sitename/sitepages/pagename.aspx
#	To test a link we need to make it a real URL not a relative one so we need to figure that out.
# if ht do nothing
# if /_ append $Tenanturl
# if /: append $Tenanturl
# if anything else append $siteUrl

Foreach($item in $siteUniqueLinks) {
	
	$thisLinkToTest = $item.link 
	$thisTestedLink =@()
	
	#figure out how to build a real URL 
	$thisLinkFirst2char = $item.link.Substring(0,2)
	if($thisLinkFirst2char -eq "ht") {
		#do nothing 
		} else {
			if($thisLinkFirst2char -eq "/_" -Or $thisLinkFirst2char -eq "/:" -Or $thisLinkFirst2char -eq "/s" -Or $thisLinkFirst2char -eq "/p") {
				#append $tenantUrl
				$thisLinkToTest = $tenantUrl+$thisLinkToTest
			} else {
				#append $siteUrl
				$thisLinkToTest = $siteUrl+$thisLinkToTest
			}
		}
	
	#write-host $thislinktotest
	#test the URL 
	$URLStatus = Test-URL $thisLinkToTest
	#write-host $URLStatus
	
	if($urlstatus -ne "OK") {
		$thisTestedLink = [PSCustomObject]@{
					"url"=$URLStatus
					"link"=$item.link 
				}
		
		
		$siteUniqueLinksNotOk +=$thisTestedLink
		}
	write-host "Testing "$testingcounter" of "$testingTotal" status: "$urlstatus
	$testingcounter = $testingcounter+1
}

#find the items that match the FileNotFound URLs in the $allSitePageTextBoxLinksNotMailto
$endTestingDateTime = get-date
write-host "Started testing : "$startTestingDateTime" Finished testing "$endTestingDateTime 
Write-Host "Total Unique links which are broken or script function Test-URL can't interpret the response code." -ForegroundColor Yellow
$siteUniqueLinksNotOk.count 

$allSitePageTextBoxLinksNotMailto404 = $allSitePageTextBoxLinksNotMailto | where-object -filterscript {$_.link -in $siteUniqueLinksNotOk.link}

Write-Host "Total pages that contain a link that is broken." -ForegroundColor Yellow
$allSitePageTextBoxLinksNotMailto404.count 

$allSitePageTextBoxLinksNotMailto404Clickable = @()

foreach($item in $allSitePageTextBoxLinksNotMailto404){
	#make clickable links to give to the user of the script, so it is easier for them to open the pages that need editing
	$thisUrlClickable = @()
	$thisLinkFirst2char = $item.url.Substring(0,2)
	if($thisLinkFirst2char -eq "ht") {
		$thisUrlClickable = $item.url
		} else {
			if($thisLinkFirst2char -eq "/_" -Or $thisLinkFirst2char -eq "/:" -Or $thisLinkFirst2char -eq "/s" -Or $thisLinkFirst2char -eq "/p") {
				#append $tenantUrl
				$thisUrlClickable = $tenantUrl+$item.url 
			} else {
				#append $siteUrl
				$thisUrlClickable = $siteUrl+$item.url 
			}
		}
	
	$allSitePageTextBoxLinksNotMailto404Clickable += [PSCustomObject]@{
					"url"=$thisUrlClickable
					"link"=$item.link
					"linktext"=$item.linktext
				}
}

Write-Host "Total rows that will be returned." -ForegroundColor Yellow
$allSitePageTextBoxLinksNotMailto404Clickable.count

$allSitePageTextBoxLinksNotMailto404Clickable = $allSitePageTextBoxLinksNotMailto404Clickable | sort-object -property url 

$allSitePageTextBoxLinksNotMailto404Clickable | out-gridview




