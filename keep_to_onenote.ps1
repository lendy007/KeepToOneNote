$NotebookPath = ""
$sourcePath = "" 
$NotebookName = "import"

# create new notebook
$OneNote = New-Object -ComObject OneNote.Application
$Scope = [Microsoft.Office.Interop.OneNote.HierarchyScope]::hsNotebooks
[ref]$xml = ""
$OneNote.OpenHierarchy($NotebookPath, "", $xml, "cftNotebook")

$SectionPath = $NotebookPath + $NotebookName + '.one'
  
# create new section
[ref]$xmlSection = ""
$OneNote.OpenHierarchy($SectionPath, "", $xmlSection, "cftSection")

Get-ChildItem $sourcePath -Filter *.html | Foreach-Object {
$source = Get-Content -Encoding UTF8 -Path $_.FullName -Raw;
Write-Host "Processing $($_.Name)"

[ref]$newpageID = ''
$OneNote.CreateNewPage($xmlSection.Value,[ref]$newpageID,[Microsoft.Office.Interop.OneNote.NewPageStyle]::npsBlankPageWithTitle)
      
[ref]$NewPageXML = ''
$OneNote.GetPageContent($newpageID.Value,[ref]$NewPageXML,[Microsoft.Office.Interop.OneNote.PageInfo]::piAll)
      
$null = [Reflection.Assembly]::LoadWithPartialName('System.Xml.Linq')
$xDoc = [System.Xml.Linq.XDocument]::Parse($NewPageXML.Value)
 
# Get OneNote XML namespace
$ns = $xDoc.Root.Name.Namespace

# first quickstyle = pagetitle
$quickstyledef = $xDoc.Descendants() | Where-Object -Property Name -Like -Value '*}QuickStyleDef'
$quickstyledef.SetAttributeValue('font','Source Sans Pro Black')
$quickstyledef.SetAttributeValue('fontColor','#80be6a')

$lastModifiedDate = ((Get-Item $_.FullName).LastWriteTimeUtc).ToString("s")
$xDoc.FirstNode.Attribute('dateTime').Value=$lastModifiedDate
$xDoc.FirstNode.Attribute('lastModifiedTime').Value=$lastModifiedDate

$title = $xDoc.Descendants() | Where-Object -Property Name -Like -Value '*}T'
if (-not $title)
{throw 'Error: can not find title element'}

# set site title
$title.Value = "$($_.Name)"

$x = $xDoc.Descendants() | Where-Object -Property Name -Like -Value '*}Title'


$OutlineNode = New-Object System.Xml.Linq.XElement( $ns + "Outline")
$OEChildrenNode = New-Object System.Xml.Linq.XElement( $ns + "OEChildren")



$html = New-Object -ComObject "HTMLFile";
$html.IHTMLDocument2_write($source);
$html.childNodes[1].childNodes[1].childNodes | ? { $_.className -like 'note*' } | % { $_.childnodes } | % {

    $node = $_

    switch ($_.className) {

    'title' {
        $title.Value = $node.innerText        
    }
    {($_ -eq "content") -or ($_ -eq "attachments")}
     {
        if ($node.innerHTML -ne $null) {
            $HTMLBlock = New-Object System.Xml.Linq.XElement( $ns + "HTMLBlock")
            $HTMLData = New-Object System.Xml.Linq.XElement( $ns + "Data")
            $CdataNode = New-Object System.Xml.Linq.XCData($node.innerHTML)
            
            $HTMLData.Add($CdataNode)
            $HTMLBlock.Add($HTMLData)
            $OEChildrenNode.Add($HTMLBlock)
        }
        else
        {
            Write-Warning "Empty content block"
        }

      }
    
    }
}

if ($OEChildrenNode.HasElements){
    $OutlineNode.Add($OEChildrenNode)

    $x.AddAfterSelf( $OutlineNode )

    
}
$onenote.UpdatePageContent($xDoc.ToString())
}