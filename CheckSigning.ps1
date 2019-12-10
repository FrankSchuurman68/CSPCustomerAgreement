# KISS Script to use a Excel File with CSP customers, exported from 
# the Partner Portal and add a New Customer Agreement.
#
# Created and Maintained by FSC: 10-12-2010
#
# #####################################################################

Install-Module -Name PartnerCenter -AllowClobber -Scope CurrentUser
Connect-PartnerCenter

$file = "C:\Users\fsc\Documents\PartnerCenterCode\CSP Customers.xlsx"
$sheetName = "Customers"
$CloudTemplateID = '998b88de-aa99-4388-a42c-1b3517d49490'
$CustomerTemplateID = '117a77b0-9360-443b-8795-c6dedc750cf9'

# Public code used to access the Excel File
$objExcel = New-Object -ComObject Excel.Application
$workbook = $objExcel.Workbooks.Open($file)
$sheet = $workbook.Worksheets.Item($sheetName)
$objExcel.Visible=$false

$rowMax = ($sheet.UsedRange.Rows).count

$rowID,$colID = 1,1
$rowName,$colName = 1,2
$rowDomain,$colDomain = 1,3

for ($i=1; $i -le $rowMax-1; $i++)
{

# Retrieve the values from the Excel File (Public Code)
$MicrosoftID = $sheet.Cells.Item($rowID+$i,$colID).text
$CompanyName = $sheet.Cells.Item($rowName+$i,$colName).text
$DomainName = $sheet.Cells.Item($rowDomain+$i,$colDomain).text

# We only want to alter the Customer with this Partner Status. For Advisor Partner status fi, we are unable to set a New Agreement, resulting in an Error.
$PartnerStatus = Get-PartnerCustomer -CustomerId $MicrosoftID

    if ($Partnerstatus.RelationshipToPartner -eq 'Reseller') {

    # Write-Host('-----------------------------------------------------')
    # Check is there is a valid CloudAgreement. We need to use these data to create a new CustomerAgreement
    $Value = Get-PartnerCustomerAgreement -AgreementType 'MicrosoftCloudAgreement' -CustomerId $MicrosoftID

    # If the Length of the agreement is greater than zero, there is a agreement. So we can continue and use the agreement fields.
        if ($Value.Type.Length -ne 0) 
        {
            # place the contents of the Cloud Agreement in variables, so we can re-use them later.
            $CustomerFirstName = $Value.PrimaryContact.FirstName
            $CustomerLastName = $Value.PrimaryContact.LastName
            $CustomerEmail = $Value.PrimaryContact.Email

            # Now we check if there is a valid CustomerAgreement in place for this customer.
            $Value = Get-PartnerCustomerAgreement -AgreementType 'MicrosoftCustomerAgreement' -CustomerId $MicrosoftID
            
            # If the length of this second Agreement is equal to 0, then there is no agremeent, as expected.
            if ($Value.Type.Length -eq 0) 
            {
                # Now we enter a new CustomerAgreement for this client, and we move on to the Next line in Excel.
                New-PartnerCustomerAgreement -AgreementType 'MicrosoftCustomerAgreement' -ContactEmail $CustomerEmail -ContactFirstName $CustomerFirstName -ContactLastName $CustomerLastName -CustomerId $MicrosoftID -TemplateId $CustomerTemplateID
            }
            # Repeat
        }

    # Just here for reference code
    # Write-Host($CompanyName)
    # Write-Host($Value.Type)
    # Write-Host($Value.PrimaryContact.FirstName)
    # Write-Host($Value.PrimaryContact.LastName)
    # Write-Host($Value.PrimaryContact.Email)

    }

}

# Close Excel Object (Public code)
$objExcel.quit()