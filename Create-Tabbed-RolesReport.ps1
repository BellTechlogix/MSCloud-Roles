<#
	CreateTabbed-RolesReport.ps1
	Created By - Kristopher Roy
	Created On - Nov 10 2021
	Modified On - Nov 10 2021

	This Script Combines the other Role Reports into a single tabbed excel formatted report
#>

#Timestamp
$runtime = Get-Date -Format "yyyyMM"

#folder to store completed reports
$rptfolder = "C:\Reports\GTIL\"

#XMLFile for output
$XMLFile = $rptFolder+$runtime+"ConsolidatedRolesReport.xml"

#XLSXFile for output
$XLSXFile = $rptFolder+$runtime+"ConsolidatedRolesReport.xlsx"

#report1 O365 User Roles Report to import
$tab1 = $rptfolder+$runtime+"-MSO365Roles.csv"

#csv 1 O365 User Roles Report to import
$O365Rolesreport = import-csv $tab1
$O365reportcount = $O365Rolesreport.count

#report2 Azure User Roles Report to import
$tab2 = $rptfolder+$runtime+"-AADRoles.csv"

#csv 2 Azure User Roles Report to import
$AzureRolesreport = import-csv $tab2

#Lets create our XML File, this is the initial formatting that it will need to understand what it is, and what styles we are using.
(
 '<?xml version="1.0"?>
<?mso-application progid="Excel.Sheet"?>
<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"
xmlns:o="urn:schemas-microsoft-com:office:office"
xmlns:x="urn:schemas-microsoft-com:office:excel"
xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"
xmlns:html="http://www.w3.org/TR/REC-html40">
<DocumentProperties xmlns="urn:schemas-microsoft-com:office:office">
<Author>Kristopher Roy</Author>
<LastAuthor>'+$env:USERNAME+'</LastAuthor>
<Created>'+(get-date)+'</Created>
<Version>16.00</Version>
</DocumentProperties>
<OfficeDocumentSettings xmlns="urn:schemas-microsoft-com:office:office">
<AllowPNG/>
</OfficeDocumentSettings>
 <ExcelWorkbook xmlns="urn:schemas-microsoft-com:office:excel">
  <WindowHeight>7920</WindowHeight>
  <WindowWidth>25530</WindowWidth>
  <WindowTopX>32767</WindowTopX>
  <WindowTopY>32767</WindowTopY>
  <ProtectStructure>False</ProtectStructure>
  <ProtectWindows>False</ProtectWindows>
 </ExcelWorkbook>
 <Styles>
  <Style ss:ID="Default" ss:Name="Normal">
   <Alignment ss:Vertical="Bottom"/>
   <Borders/>
   <Font ss:FontName="Calibri" x:Family="Swiss" ss:Size="11" ss:Color="#000000"/>
   <Interior/>
   <NumberFormat/>
   <Protection/>
  </Style>
  <Style ss:ID="s62">
   <Font ss:FontName="Calibri" x:Family="Swiss" ss:Size="11" ss:Color="#FFFFFF"
    ss:Bold="1"/>
   <Interior ss:Color="#4472C4" ss:Pattern="Solid"/>
  </Style>
    <Style ss:ID="s63">
    <Alignment ss:Horizontal="Left" ss:Vertical="Bottom"/>
   <Font ss:FontName="Calibri" x:Family="Swiss" ss:Size="11" ss:Color="#000000"/>
  </Style>
    <Style ss:ID="s64">
   <Alignment ss:Horizontal="Left" ss:Vertical="Bottom"/>
   <Font ss:FontName="Calibri" x:Family="Swiss" ss:Size="11" ss:Color="#000000"/>
   <Interior ss:Color="#FF0000" ss:Pattern="Solid"/>
  </Style>
    <Style ss:ID="s65">
   <Alignment ss:Horizontal="Left" ss:Vertical="Bottom"/>
   <Font ss:FontName="Calibri" x:Family="Swiss" ss:Size="11" ss:Color="#000000"/>
   <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
  </Style>
  <Style ss:ID="s66">
   <Alignment ss:Horizontal="Left" ss:Vertical="Bottom"/>
   <Font ss:FontName="Calibri" x:Family="Swiss" ss:Size="11" ss:Color="#000000"/>
   <Interior ss:Color="#00B050" ss:Pattern="Solid"/>
  </Style>
 </Styles>')>$XMLFile

  #Tab1 Report
add-content $XMLFile (
 '<Worksheet ss:Name="'+($runtime)+'-O365Roles">
  <Table ss:ExpandedColumnCount="18" ss:ExpandedRowCount="'+($O365reportcount+1)+'" x:FullColumns="1"
   x:FullRows="1" ss:DefaultRowHeight="15">
   <Column ss:AutoFitWidth="0" ss:Width="119.25"/>
   <Column ss:Width="111.75"/>
   <Column ss:Width="77.25"/>
   <Column ss:Width="99"/>
   <Column ss:AutoFitWidth="0" ss:Width="111.75" ss:Span="1"/>
   <Column ss:Index="7" ss:Width="58.5"/>
   <Column ss:AutoFitWidth="0" ss:Width="109.5"/>
   <Column ss:Width="122.25"/>
   <Column ss:Width="141.75"/>
   <Column ss:Width="141.75"/>
   <Column ss:Width="141.75"/>
   <Column ss:Width="141.75"/>
   <Column ss:Width="141.75"/>
   <Column ss:Width="141.75"/>
   <Column ss:Width="141.75"/>
   <Column ss:Width="141.75"/>
   <Column ss:Width="141.75"/>
   <Column ss:Width="141.75"/>
   <Column ss:Width="141.75"/>
   <Row ss:AutoFitHeight="0">
    <Cell ss:StyleID="s62"><Data ss:Type="String">DisplayName</Data></Cell>
    <Cell ss:StyleID="s62"><Data ss:Type="String">EmailAddress</Data></Cell>
    <Cell ss:StyleID="s62"><Data ss:Type="String">Company Administrator</Data></Cell>
    <Cell ss:StyleID="s62"><Data ss:Type="String">Guest Inviter</Data></Cell>
    <Cell ss:StyleID="s62"><Data ss:Type="String">User Administrator</Data></Cell>
    <Cell ss:StyleID="s62"><Data ss:Type="String">Service Support Administrator</Data></Cell>
    <Cell ss:StyleID="s62"><Data ss:Type="String">Directory Readers</Data></Cell>
    <Cell ss:StyleID="s62"><Data ss:Type="String">Directory Writers</Data></Cell>
    <Cell ss:StyleID="s62"><Data ss:Type="String">Exchange Administrator</Data></Cell>
    <Cell ss:StyleID="s62"><Data ss:Type="String">SharePoint Administrator</Data></Cell>
	<Cell ss:StyleID="s62"><Data ss:Type="String">Skype for Business Administrator</Data></Cell>
	<Cell ss:StyleID="s62"><Data ss:Type="String">Directory Synchronization Accounts</Data></Cell>
	<Cell ss:StyleID="s62"><Data ss:Type="String">Intune Administrator</Data></Cell>
	<Cell ss:StyleID="s62"><Data ss:Type="String">Power BI Administrator</Data></Cell>
	<Cell ss:StyleID="s62"><Data ss:Type="String">Reports Reader</Data></Cell>
	<Cell ss:StyleID="s62"><Data ss:Type="String">Teams Administrator</Data></Cell>
	<Cell ss:StyleID="s62"><Data ss:Type="String">Global Reader</Data></Cell>
	<Cell ss:StyleID="s62"><Data ss:Type="String">Groups Administrator</Data></Cell>
   </Row>')
   FOREACH($user in $O365Rolesreport)
   {
   add-content $XMLFile ('
      <Row ss:AutoFitHeight="0">
    <Cell><Data ss:Type="String">'+($user.DisplayName)+'</Data></Cell>
    <Cell><Data ss:Type="String">'+($user.EmailAddress)+'</Data></Cell>
    <Cell><Data ss:Type="String">'+($user.'Company Administrator')+'</Data></Cell>
    <Cell><Data ss:Type="String">'+($user.'Guest Inviter')+'</Data></Cell>
    <Cell><Data ss:Type="String">'+($user.'User Administrator')+'</Data></Cell>
    <Cell><Data ss:Type="String">'+($user.'Service Support Administrator')+'</Data></Cell>
    <Cell><Data ss:Type="String">'+($user.'Directory Readers')+'</Data></Cell>
    <Cell><Data ss:Type="String">'+($user.'Directory Writers')+'</Data></Cell>
    <Cell><Data ss:Type="String">'+($user.'Exchange Administrator')+'</Data></Cell>
    <Cell><Data ss:Type="String">'+($user.'SharePoint Administrator')+'</Data></Cell>
    <Cell><Data ss:Type="String">'+($user.'Skype for Business Administrator')+'</Data></Cell>
    <Cell><Data ss:Type="String">'+($user.'Directory Synchronization Accounts')+'</Data></Cell>
    <Cell><Data ss:Type="String">'+($user.'Intune Administrator')+'</Data></Cell>
    <Cell><Data ss:Type="String">'+($user.'Power BI Administrator')+'</Data></Cell>
    <Cell><Data ss:Type="String">'+($user.'Reports Reader')+'</Data></Cell>
    <Cell><Data ss:Type="String">'+($user.'Teams Administrator')+'</Data></Cell>
    <Cell><Data ss:Type="String">'+($user.'Global Reader')+'</Data></Cell>
    <Cell><Data ss:Type="String">'+($user.'Groups Administrator')+'</Data></Cell>
   </Row>
   ')
   }
   $user = $null

   #Close out the XML
      add-content $XMLFile ('</Table>
  <WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel">
   <PageSetup>
    <Header x:Margin="0.3"/>
    <Footer x:Margin="0.3"/>
    <PageMargins x:Bottom="0.75" x:Left="0.7" x:Right="0.7" x:Top="0.75"/>
   </PageSetup>
   <Selected/>
   <ProtectObjects>False</ProtectObjects>
   <ProtectScenarios>False</ProtectScenarios>
  </WorksheetOptions>
 </Worksheet>
</Workbook>')

