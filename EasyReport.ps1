<#
  .NOTES

  Created with:  PowerShell ISE
  Creation Date: 03-10-2016
  Created by:    Christof Van Geendertaelen
  Filename:      EasyReport.ps1

  .SYNOPSIS
        This scripts creates an html report using only a few parameters.

  .DESCRIPTION
        When you need an HTML report fast, this script allows you to generate
        an HTML report with only a few parameters. The rest of the logic is
        handled by this script.

  .PARAMETER Logo
        The path to the logo. The logo is converted by the script to base 64
        encoding. This is the only way to show the logo in an Outlook email.

  .PARAMETER Title
        This is the title of the report.
        
  .PARAMETER Text
        This text is displayed directly under the header of the report and can
        be used to write a brief introduction for the upcoming report or data. 

  .PARAMETER Data
        The actual data that will be displayed is passed using the Data parameter.
        
        The data should be delivered as an html fragment which can be generated
        quite easily using the ConvertT-HTML -Fragment Commandlet which is
        natively available in PowerShell.

  .EXAMPLE
        EasyReport.ps1 -Logo 'logo.png' -Title 'ReportHeader' -Text $TextVariable
            -Data $DataVariable

#>

# Parameter section

Param(
    
    [string]$LogoPath,
    [string]$ReportTitle,
    [string]$ReportText,
    [string]$ReportData

)

Function ConvertLogoToBase64($ConvertLogoToBase64_LogoPath)
{
    # Convert the logo to a Base 64 encoded value that can be enclosed in the Report

    $Base64Logo = [convert]::ToBase64String((Get-Content $ConvertLogoToBase64_LogoPath -Encoding Byte))
    
    Return $Base64Logo
}

Function FindImageHeight($FindImageHeight_LogoPath)
{
    # Find the height of the logo

    $Image = [System.Drawing.Image]::FromFile((Get-ChildItem $FindImageHeight_LogoPath).FullName)

    $Height = ($Image | Select -ExpandProperty Size | Select Height).Height

    Return $Height
}

Function BuildHtml($Logo, $Height, $Title, $Text, $Data)
{
    # Build the actual HTML code

    $HTMLHeader = @"
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Frameset//EN" "http://www.w3.org/TR/html4/frameset.dtd">
<html>
<head>
<title>My Systems Report</title>
<style>
	body {
		font-family: "Segoe UI", Tahoma, Geneva, Verdana, sans-serif;
	}
	table, tr, td {
		border: 1px solid black;
		border-collapse: collapse;
	}
	th {
		background-color: #ffffff;
	}
	tr:hover {
		background-color: #f5f5f5;
	}
	#Logo {
		float: left;
        height: $($Height)px;
        vertical-align: middle;
	}
	#Title {
		position: relative;
		left: 10px;
        height: $($Height)px;
        line-height: $($Height)px;
	}
</style>
</head>
<div>
<img id="Logo" src="data:image/png;base64,$($Logo)" />
<h1 id="Title">$($Title)</h1>
</div>
<hr />
<p>$($Text)</p>
<br />
$($data)
"@

    Return $HTMLHeader

}

$ConvertLogoToBase64 = ConvertLogoToBase64 $LogoPath

$LogoHeight = FindImageHeight $LogoPath

$HTMLReport = BuildHtml $ConvertLogoToBase64 $LogoHeight $ReportTitle $ReportText $ReportData

$HTMLReport