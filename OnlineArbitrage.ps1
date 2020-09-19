<#
.DESCRIPTION
    This script scrapes hotukdeals.com and argos.co.uk for items which are marked at a 60% or greater discount and send you an email with them in a list.

.NOTES
    Author: Ian Waters
    Last Edit: 2020-09-19
    Version 1.0 - Public Release to GitHub
#>



#master array used for collecting all products from different url's
[System.Collections.ArrayList]$productList = New-Object -TypeName "System.Collections.ArrayList"

Class Product
{
    [string]$image = ""
    [string]$title = ""
    [string]$url = ""
    [decimal]$oldPrice = ""
    [decimal]$newPrice = ""
    [decimal]$percentage = ""
}


Class ArgosClearanceFinder
{
    [string]GetImage([string]$markup)
    {
        #strip article image section
        $regex = "<picture>(.*?)</picture>"
        $articleImageSection = $matches = Select-String -InputObject $markup -Pattern $regex -AllMatches 

        #strip image
        $regex   = "<img src=""(.*?);"
        $matches = Select-String -InputObject $matches.Matches[0].Value -Pattern $regex -AllMatches 

        if($matches.Matches.Count -gt 0)
        {
            $image = "https:" + $matches.Matches[0].Value.Substring(10,$matches.Matches[0].Value.Length-11)
    
            return $image   
        }
   
        return ""
    }

    [string]GetTitle([string]$markup)
    {
        #strip thread title section
        $regex   = "component-product-card-title(.*?)</a>"
        $matches = Select-String -InputObject $markup -Pattern $regex -AllMatches 
        $threadTitleSection = $matches.Matches[0].Value

        #strip title
        $regex   = "/>(.*?)</a>"
        $matches = Select-String -InputObject $threadTitleSection -Pattern $regex -AllMatches 
        $title = $matches.Matches[0].Value
        $title = $title.Substring($title.IndexOf('>')+1,($title.LastIndexOf('<')-$title.IndexOf('"'))-3)
        
        return $title
    }

    [string]GetLink([string]$markup)
    {
        #strip thread title section
        $regex   = "component-product-card-title(.*?)</a>"
        $matches = Select-String -InputObject $markup -Pattern $regex -AllMatches 
        $threadTitleSection = $matches.Matches[0].Value

        #strip title
        $regex   = "content=""(.*?)"""
        $matches = Select-String -InputObject $threadTitleSection -Pattern $regex -AllMatches 
        $url = $matches.Matches[0].Value
        $url = $url.Substring($url.IndexOf('"')+1,($url.LastIndexOf('"')-$url.IndexOf('"'))-1)
        
        return "https://www.argos.co.uk" + $url
    }

    [decimal]GetOldprice([string]$markup)
    {
        #strip thread price section
        $regex   = "Original Price(.*?)</div>"
        $matches = Select-String -InputObject $markup -Pattern $regex -AllMatches 
                   
        if($matches.Matches.Count -gt 0)
        {
            $threadPriceSection = $matches.Matches[0].Value
            $threadPriceSection = $threadPriceSection.Replace("?","£")
            $threadPriceSection = $threadPriceSection.Replace("�","£")
            
            #strip price
            $price = $threadPriceSection.Substring($threadPriceSection.IndexOf('£')+1,($threadPriceSection.Length-($threadPriceSection.IndexOf('£')+1))-6)
                       
            return $price
        }

        return [decimal]0
    }

    [decimal]GetPercentage([decimal]$oldPrice, [decimal]$newPrice)
    {
        if($oldPrice -eq 0 -or $newPrice -eq 0)
        {
            return 0
        }

        try
        {
            $round = [math]::Round((($oldPrice-$newPrice)/$oldPrice)*100,2)
            return $round
        }
        catch
        {
        }
        
        return 0
    }

    [decimal]GetNewPrice([string]$markup)
    {
        #strip thread price section
        $regex   = "ProductCardstyles__PriceText(.*?)</div>"
        $matches = Select-String -InputObject $markup -Pattern $regex -AllMatches 
                   
        if($matches.Matches.Count -gt 0)
        {
            $threadPriceSection = $matches.Matches[0].Value
            $threadPriceSection = $threadPriceSection.Replace("</strong>","")

            #strip price
            $price = $threadPriceSection.Substring($threadPriceSection.IndexOf('£')+1,($threadPriceSection.Length-($threadPriceSection.IndexOf('£')+1))-6)

            return $price
        }

        return [decimal]0
    }

    GetProducts([string]$url,[System.Collections.ArrayList]$productList)
    {
        $request = Invoke-WebRequest -Uri $url -UseBasicParsing
        $page = $request.RawContent -replace "`r`n",''
        $page = $request.RawContent -replace "`n",''
        $regex = "<div data-test=""component-product-card(.*?)>(.*?)<div class=""ProductCardstyles__ButtonContainer(.*?)"
        $matches = Select-String -InputObject $page -Pattern $regex -AllMatches 

        ForEach($match in $matches.Matches)
        {
            [Product]$product   = New-Object Product
            $product.image      = $this.GetImage($match.Value)
            $product.title      = $this.GetTitle($match.Value)
            $product.url        = $this.GetLink($match.Value)
            $product.oldPrice   = $this.GetOldprice($match.Value)
            $product.newPrice   = $this.GetNewPrice($match.Value)
            $product.percentage = $this.GetPercentage($product.oldPrice,$product.newPrice)
            $index              = $productList.add($product)
        }
    }
}

Class HotDealsFinder
{
    [string]GetImage([string]$markup)
    {
        #A BIT HACKY, NEED TO REDO AT SOME POINT

        #strip article image section
        $regex = "src=""(.*?)"""
        $articleImageSection = $matches = Select-String -InputObject $markup -Pattern $regex -AllMatches 

        #strip image
        $regex   = "title=""(.*?)"""
        $matches = Select-String -InputObject $articleImageSection -Pattern $regex -AllMatches 
    
        if($matches.Matches.Count -gt 0)
        {
            $image = $matches.Matches[0].Value
    
            if($matches.Matches.count -eq 2)
            {
                $image = $image.Substring($image.IndexOf('"')+1,($image.LastIndexOf('"')-$image.IndexOf('"'))-1)
                return $image
            }
        }
   
        #if here then try second get image method

        #strip article image section
        $regex = "class=""cept-thread-image-link(.*?)<\/a>"
        $matches = Select-String -InputObject $markup -Pattern $regex -AllMatches 
    
        if($matches.Matches.Count -gt 0)
        {
            $image = $matches.Matches[0].value

            #strip image
            $regex   = "data-lazy-img=""(.*?)"""
            $matches = Select-String -InputObject $image -Pattern $regex -AllMatches 
    
            $image = $matches.Matches[0].Value

            $image = $image.Substring($image.IndexOf('"')+1,($image.LastIndexOf('"')-$image.IndexOf('"'))-1)
            $image = $image.Replace("\",'')
            $image = "https://" + $image.Substring($image.IndexOf('images'),($image.LastIndexOf(".jpg")-$image.IndexOf('images'))) + ".jpg"

            return $image
        }

        return ""
    }

    [string]GetTitle([string]$markup)
    {
        #strip thread title section
        $regex   = "cept-tt(.*?)</a>"
        $matches = Select-String -InputObject $markup -Pattern $regex -AllMatches 
        $threadTitleSection = $matches.Matches[0].Value

        #strip title
        $regex   = "title=""(.*?)"""
        $matches = Select-String -InputObject $threadTitleSection -Pattern $regex -AllMatches 
        $title = $matches.Matches[0].Value
        $title = $title.Substring($title.IndexOf('"')+1,($title.LastIndexOf('"')-$title.IndexOf('"'))-1)
        
        return $title
    }

    [string]GetLink([string]$markup)
    {
        #strip thread title section
        $regex   = "cept-tt(.*?)</a>"
        $matches = Select-String -InputObject $markup -Pattern $regex -AllMatches 
        $threadTitleSection = $matches.Matches[0].Value

        #strip title
        $regex   = "href=""(.*?)"""
        $matches = Select-String -InputObject $threadTitleSection -Pattern $regex -AllMatches 
        $url = $matches.Matches[0].Value
        $url = $url.Substring($url.IndexOf('"')+1,($url.LastIndexOf('"')-$url.IndexOf('"'))-1)
        
        return $url
    }

    [decimal]GetOldprice([string]$markup)
    {
        #strip thread price section
        $regex   = "<span class=""mute--text(.*?)</span>"
        $matches = Select-String -InputObject $markup -Pattern $regex -AllMatches 
    
        if($matches.Matches.Count -gt 0)
        {
            $threadPriceSection = $matches.Matches[0].Value
            #strip price
            $price = $threadPriceSection.Substring($threadPriceSection.IndexOf('>')+1,($threadPriceSection.LastIndexOf('<')-$threadPriceSection.IndexOf('>'))-1)
        
            try
            {
                $price = [decimal]::parse($price.Substring(1,$price.Length-1))
            }
            catch
            {
                $price = 0
            }
        
            return $price
        }

        return [decimal]0
    }

    [decimal]GetPercentage([string]$markup)
    {
        #strip thread price section
        $regex   = "<span class=""space--ml-1(.*?)</span>"
        $matches = Select-String -InputObject $markup -Pattern $regex -AllMatches 
    
        if($matches.Matches.Count -gt 0)
        {
            $threadPriceSection = $matches.Matches[0].Value
            #strip price
            $percentage = $threadPriceSection.Substring($threadPriceSection.IndexOf('>')+1,($threadPriceSection.LastIndexOf('<')-$threadPriceSection.IndexOf('>'))-1)
        
            try
            {
                $percentage = [decimal]::parse($percentage.Substring(0,$percentage.Length-1))
            }
            catch
            {
                 $percentage = [decimal]0
            }


            return [decimal]$percentage
        }

        return [decimal]0
    }

    [decimal]GetNewPrice([string]$markup)
    {
        #strip thread price section
        $regex   = "<span class=""thread-price(.*?)</span>"
        $matches = Select-String -InputObject $markup -Pattern $regex -AllMatches 
    
        if($matches.Matches.Count -gt 0)
        {
            $threadPriceSection = $matches.Matches[0].Value
            #strip price
            $price = $threadPriceSection.Substring($threadPriceSection.IndexOf('>')+1,($threadPriceSection.LastIndexOf('<')-$threadPriceSection.IndexOf('>'))-1)
        
            try
            {
                $price = [decimal]::parse($price.Substring(1,$price.Length-1))
            }
            catch{$price = 0}
        
            return [decimal]$price
        }

        return [decimal]0
    }

    GetProducts([string]$url,[System.Collections.ArrayList]$productList)
    {
        $request = Invoke-WebRequest -Uri $url -UseBasicParsing
        $page = $request.RawContent -replace "`r`n",''
        $page = $request.RawContent -replace "`n",''
        $regex = "<article(.*?)>(.*?)</article>"
        $matches = Select-String -InputObject $page -Pattern $regex -AllMatches 

        ForEach($match in $matches.Matches)
        {
            [Product]$product   = New-Object Product
            $product.image      = $this.GetImage($match.Value)
            $product.title      = $this.GetTitle($match.Value)
            $product.url        = $this.GetLink($match.Value)
            $product.oldPrice   = $this.GetOldprice($match.Value)
            $product.newPrice   = $this.GetNewPrice($match.Value)
            $product.percentage = $this.GetPercentage($match.Value)
            $index              = $productList.add($product)
        }
    }
}


#Hot Deals UK Finder Object
[HotDealsfinder]$hotDealsFinder = New-Object -TypeName HotDealsFinder
$hotDealsFinder.GetProducts("http://www.hotukdeals.com/", $productList)
$hotDealsFinder.GetProducts("https://www.hotukdeals.com/tag/gaming", $productList)
$hotDealsFinder.GetProducts("https://www.hotukdeals.com/tag/gaming?page=2", $productList)
$hotDealsFinder.GetProducts("https://www.hotukdeals.com/tag/gaming?page=3", $productList)
$hotDealsFinder.GetProducts("https://www.hotukdeals.com/tag/electronics", $productList)
$hotDealsFinder.GetProducts("https://www.hotukdeals.com/tag/electronics?page=2", $productList)
$hotDealsFinder.GetProducts("https://www.hotukdeals.com/tag/electronics?page=3", $productList)
$hotDealsFinder.GetProducts("https://www.hotukdeals.com/tag/beauty", $productList)
$hotDealsFinder.GetProducts("https://www.hotukdeals.com/tag/beauty?page=2", $productList)
$hotDealsFinder.GetProducts("https://www.hotukdeals.com/tag/beauty?page=3", $productList)
$hotDealsFinder.GetProducts("https://www.hotukdeals.com/tag/kids", $productList)
$hotDealsFinder.GetProducts("https://www.hotukdeals.com/tag/kids?page=2", $productList)
$hotDealsFinder.GetProducts("https://www.hotukdeals.com/tag/kids?page=3", $productList)
$hotDealsFinder.GetProducts("https://www.hotukdeals.com/tag/home", $productList)
$hotDealsFinder.GetProducts("https://www.hotukdeals.com/tag/home?page=2", $productList)
$hotDealsFinder.GetProducts("https://www.hotukdeals.com/tag/home?page=3", $productList)
$hotDealsFinder.GetProducts("https://www.hotukdeals.com/tag/sports-fitness", $productList)
$hotDealsFinder.GetProducts("https://www.hotukdeals.com/tag/sports-fitness?page=2", $productList)
$hotDealsFinder.GetProducts("https://www.hotukdeals.com/tag/sports-fitness?page=3", $productList)

#Argos Clearance Finder Object
[ArgosClearanceFinder]$argosClearanceFinder = New-Object -TypeName ArgosClearanceFinder
$argosClearanceFinder.GetProducts("https://www.argos.co.uk/clearance/technology/c:29949/?tag=ar:events:clearance:m035:technology", $productList)
$argosClearanceFinder.GetProducts("https://www.argos.co.uk/clearance/technology/c:29949/clearance:true/opt/page:2/", $productList)
$argosClearanceFinder.GetProducts("https://www.argos.co.uk/clearance/technology/c:29949/clearance:true/opt/page:3/", $productList)
$argosClearanceFinder.GetProducts("https://www.argos.co.uk/clearance/technology/c:29949/clearance:true/opt/page:4/", $productList)
$argosClearanceFinder.GetProducts("https://www.argos.co.uk/clearance/technology/c:29949/clearance:true/opt/page:5/", $productList)
$argosClearanceFinder.GetProducts("https://www.argos.co.uk/clearance/technology/c:29949/clearance:true/opt/page:6/", $productList)

function Format-HTML([System.Collections.ArrayList]$productList)
{
    [string]$formattedHTML = ""

    $formattedHTML += "<!DOCTYPE html>
                       `n`r<html>
                       `n`r<head>
                       `n`r<style>
                       `n`rtable, th, td {
                       `n`r   border: 1px solid black;
                       `n`r   border-collapse: collapse;
                       `n`r}
                       `n`rth, td {
                       `n`r   padding: 5px;
                       `n`r  text-align: left;    
                       `n`r}
                       `n`r</style>
                       `n`r</head>
                       `n`r<body>
                       
                       `n`r<h2>Latest Deals</h2>
                       
                       `n`r<table style=""width:100%"">
                       "

    $formattedHTML += "`n`r<tr>
                       `n`r<th>Image</th>
                       `n`r<th>Title</th>
                       `n`r<th>Old Price</th>
                       `n`r<th>New Price</th>
                       `n`r<th>Percentage</th>
                       `n`r</tr>"

    for($i = 0;$i-lt$productList.count;$i++)
    {
        [Product]$prod = $productList[$i]

        $formattedHTML += "`n`r<tr>"

        $formattedHTML += "`n`r<td><a href=""$($prod.url)""><img src=""$($prod.image)"" width=""200"" height=""200""></a></td>"
        $formattedHTML += "`n`r<td>$($prod.title)</td>"
        $formattedHTML += "`n`r<td>$($prod.oldPrice)</td>"
        $formattedHTML += "`n`r<td>$($prod.newPrice)</td>"
        
        $formattedHTML += "`n`r<td>$($prod.percentage)</td>"
        
        $formattedHTML += "`n`r</tr>"
    }

    $formattedHTML += "`n`r</table>
                       `n`r</body>
                       `n`r</html>"
    
    return $formattedHTML
}

function Send-Mail($body)
{
    $notificationToEmailAddress = "<to email address>"
    $notificationFromEmailAddress = "<from email address>"
    $smtpServer = "<your smtp server>"
    $smtpPort = "25"
    $subject = "Latest Deals"
    Send-MailMessage -From $notificationFromEmailAddress -to $notificationToEmailAddress -Subject $subject -Body $body -BodyAsHtml -SmtpServer $smtpServer -port $smtpPort
}

<#
for($i = 0;$i-le$productList.count;$i++)
{
    [Product]$prod = $productList[$i]
    write-host "=========" 
    Write-Host "Image:" $prod.image
    Write-Host "Title:" $prod.title
    Write-Host "Old Price:" $prod.oldPrice
    Write-Host "New Price:" $prod.newPrice
    Write-Host "Percentage:" $prod.percentage
}
#>

#remove any products from list with percentage of less than 40 and sort
$productList = $productList | where-object {$_.percentage -gt 60} | Sort-Object percentage -Descending
$productList = $productList | where-object {$_.title -notlike "*Amazon*" -and $_.title -notlike "*playstation store*" -and $_.title -notlike "*app store*" -and $_.title -notlike "*steam*" -and $_.title -notlike "*vodafone*"}


clear-host
$emailBody = Format-HTML $productList
$emailBody
Send-Mail $emailBody


