Set-StrictMode -Version 2 

# Identify Who Should Be Notified In Case This Program Does Not Run Successfully
$strErrorNotificaitonEmailTo = "bomf@ford.com;JADEEL@ford.com;psatyan1@ford.com;braokott@ford.com"

# Used To Identify When To Send Team Password Expiration Email
$intDaysSendNotification = 20
$intDaysSendNotificationStandard = $intDaysSendNotification

# Used To Identify Every How Many Days
$intNotificaitonInterval = 2
$intNotificaitonIntervalStandard = $intNotificaitonInterval

# Used To Identify When To Send Supervisor Password Expiration Email
$intDaysSendSupervisorNotification = 5
$intDaysSendSupervisorNotificationStandard = $intDaysSendSupervisorNotification

#Identify What Share Point Site The Credential Management Tool Was Created Under
$strSharePointSite = "https://it1.spt.ford.com/sites/FEDEBOM/CMT/"

# Share Point List Conneciton IDs
$strSharePointListName = "CredentialDetails"
$strSharePointViewID = "{EB34BF90-0CC6-4958-93B5-A690EA3EEC1C}"

# Summary Email
$strRecieveEmailSummary = "Yes"
$strEmailSummaryRecipients = "dmorasa@ford.com;akalayan@ford.com;ASENGUP5@ford.com;cmirall1@ford.com;braokott@ford.com"
$strEmailSummaryCadence = "Monthly"

# Email Method
$strEmailMethod = "Outlook"

# Send Notificaitons For All IDs Or Only Production
$strSendNotificationsFor = "All Credentials"

# Override Expiration Email Hash 
$hasOverride=@{}

# Share Point Web URI - WSDL Address
$strSharePointWSDLAddress = $strSharePointSite + "/_vti_bin/lists.asmx?WSDL"

# Current Date DD/MM/YYYY
$strCurrentDate = (Get-Date)

# Process Success Flag
$booReturn = "False"



######################################################################
# Function Main                                                       
#                                                                     
# Description:                                                        
# Controls Process Flow.                                              
######################################################################
function Main {
    # Loop Through SharePoint List Entries To Check Expiration Days And Send Notifications
    $booReturn = LoopThroughSharePointListEntries
    
    # Write Process Status To Host
    if($booReturn -eq "True"){
        write-host "Process Completed Successfully" $strCurrentDate.ToString("MM/dd/yyyy")
    }
    else{
        write-host "Process Did Not Complete Successfully" $strCurrentDate.ToString("MM/dd/yyyy")
    }
}



######################################################################
# Function LoopThroughSharePointListEntries                           
#                                                                     
# Description:                                                        
# Updates All Entries In A Share Point List With Passed Parameters    
######################################################################
function LoopThroughSharePointListEntries(){
    
    # Create xml query to retrieve list.
    $xmlDoc = new-object System.Xml.XmlDocument
    $query = $xmlDoc.CreateElement("Query")
    $viewFields = $xmlDoc.CreateElement("ViewFields")
    $queryOptions = $xmlDoc.CreateElement("QueryOptions")
    $query.set_InnerXml("FieldRef Name='Full Name'")
    $rowLimit = "2000"
    
    $list = $null
    $service = $null
    
    try{
        $service = New-WebServiceProxy -Uri $strSharePointWSDLAddress  -Namespace SpWs  -UseDefaultCredential -ErrorAction:'Stop'
        $booReturn = "True"
    }
    catch{
        ProcessError "Function: LoopThroughSharePointListEntries" "Try: Create New-WebServiceProxy" "Error: $_" $strErrorNotificaitonEmailTo
    }
    
    if($service -ne $null){
        try{ 
            $list = $service.GetListItems($strSharePointListName, "", $query, $viewFields, $rowLimit, $queryOptions, "")
            $booReturn = "True"
        }
        catch{
            ProcessError "Function: LoopThroughSharePointListEntries" "Try: GetListItems" "Error: $_" $strErrorNotificaitonEmailTo
        }
    }
    
    $arrSummaryCollection = New-Object System.Collections.ArrayList


    if($list.data.ItemCount -gt 0){

        $list.data.row | ForEach-Object {

            if($hasOverride.count -gt 0){
                # Look For Override Notificaitons For ID Type
                foreach ($arrOverrideCredType in $hasOverride.keys) {
                    if($_.("ows_ID_x0020_Type") -eq $arrOverrideCredType){
                        $intDaysSendNotification = $($hasOverride.Item($arrOverrideCredType))[0]
                	       $intNotificaitonInterval = $($hasOverride.Item($arrOverrideCredType))[1]
                        $intDaysSendSupervisorNotification = $($hasOverride.Item($arrOverrideCredType))[2]
                        break
                    }
                }
            }

            $intRowID = $_.("ows_ID")
            $datExpDate = [datetime]$_.("ows_Expiration_x0020_Date")
            $intExpDays = (New-TimeSpan -Start $strCurrentDate.ToString("MM/dd/yyyy") -End $datExpDate.ToString("MM/dd/yyyy")).Days

            if((($strRecieveEmailSummary -eq "Yes") -and ($strEmailSummaryCadence -eq "Monthly") -and ($strCurrentDate.ToString("dd") -eq "01") -and ($strCurrentDate.ToString("yyyy") -eq $datExpDate.ToString("yyyy")) -and ($datExpDate.ToString("MM") -eq $strCurrentDate.ToString("MM"))) -or ($intExpDays -lt 1)){
                if(($strSendNotificationsFor -eq "All Credentials") -or ($_.("ows_Environment") -eq "Production")){
                    $arrSummaryCollection.add(($_.("ows_Applicaiton_x0020__x0028_ITMS_x0"), $_.("ows_Credential"), $datExpDate, $intExpDays, $_.("ows_Link_x0020_To_x0020_Process_x002"), $intRowID,$_.("ows_Environment")))
                }
            }
            elseif((($strRecieveEmailSummary -eq "Yes") -and ($strEmailSummaryCadence -eq "Weekly") -and ($strCurrentDate.ToString("ddd") -eq "Mon") -and ($datExpDate -ge $strCurrentDate) -and ($datExpDate -le $strCurrentDate.AddDays(7))) -or ($intExpDays -lt 1)){
                if(($strSendNotificationsFor -eq "All Credentials") -or ($_.("ows_Environment") -eq "Production")){
                    $arrSummaryCollection.add(($_.("ows_Applicaiton_x0020__x0028_ITMS_x0"), $_.("ows_Credential"), $datExpDate, $intExpDays, $_.("ows_Link_x0020_To_x0020_Process_x002"), $intRowID,$_.("ows_Environment")))
                }
            }

            $booReturn = UpdateSharePointListEntries $intRowID $intExpDays

            if($booReturn -eq "False"){
                return $booReturn
                exit
            }
        
            if(($intExpDays -le ($intDaysSendNotification))){
                if(($intExpDays % $intNotificaitonInterval -eq 0) -or ($intExpDays -le ($intDaysSendSupervisorNotification))){
                    if($intExpDays -eq 0){
                        $strMessageExpirationText = '</b> expired <b>today</b>.'
                        $strMailSubject = 'Action Required - ' + $_.("ows_Applicaiton_x0020__x0028_ITMS_x0") + ' Credential "' + $_.("ows_Credential") + '" Expired Today'
                    }
                    elseif($intExpDays -lt 0){
                        $strMessageExpirationText = '</b> expired <b>' + (-1 * $intExpDays) + '</b> day(s) ago.'
                        $strMailSubject = 'Action Required - ' + $_.("ows_Applicaiton_x0020__x0028_ITMS_x0") + ' Credential "' + $_.("ows_Credential") + '" Expired ' + (-1 * $intExpDays) + ' Day(s) Ago'
                    }
                    else{
                        $strMessageExpirationText = '</b> will expire in <b>' + $intExpDays + '</b> day(s).'
                        $strMailSubject = 'Action Required - ' + $_.("ows_Applicaiton_x0020__x0028_ITMS_x0") + ' Credential "' + $_.("ows_Credential") + '" Will Expire In ' + $intExpDays + ' Day(s)'
                    }
                
                    $strProcessLInk = $_.("ows_Link_x0020_To_x0020_Process_x002").substring(0,$_.("ows_Link_x0020_To_x0020_Process_x002").indexof(","))
                
                    $strDirectIDLink = $strSharePointSite + "/Lists/$strSharePointListName/EditForm.aspx?ID=" + $_.("ows_id") + "&Source=" + $strSharePointSite + "CredentialManagementCatalog/CredentialDetails.aspx"
                
                    $strPDOManagedMessage = '<span style="font-family: Arial;font-size: medium">'`
                                         + '   <b>' + $_.("ows_Managed_x0020_By") + ' Managed - Credential Expiration Notification</b><br>'`
                                         + '</span>'`
                                         + '<span style="font-family: Arial;font-size: x-small">'`
                                         + '	The ' + $_.("ows_Applicaiton_x0020__x0028_ITMS_x0") + ' ' + $_.("ows_ID_x0020_Type") + ' credential <b>' +  $_.("ows_Credential") + $strMessageExpirationText + ' &nbsp;Please use the following procedure to manage the credential.<br>'`
                                         + '	<br>'`
                                         + '</span>'`
                                         + '<span style="font-family: Arial;font-size: x-small">'`
                                         + '	<b>Credential Management Procedure:</b><br>'`
                                         + '	<a href="' + $strProcessLInk + '">' + $strProcessLInk + '</a><br>'`
                                         + '	<br>'`
                                         + '	After the credential is managed, use the following link to update the credentials “Expiration Date” field or you will continue to receive credential expiration emails.<br>'`
                                         + '   <br>'`
                                         + '</span>'`
                                         + '<span style="font-family: Arial;font-size: x-small">'`
                                         + '	<b>Credential Management Tool:</b><br>'`
                                         + '	<a href="' + $strDirectIDLink + '">'`
                                         + '	' + $strDirectIDLink + '</a></span>'`
                                         + ' </span>'`
                                         + '<br>'`
                                         + '<br>'`
                                         + '<span style="font-family: Arial;font-size: xx-small">'`
                                         + '   Innovate & Go Further <br>'`
                                         + '   Credential Managment Tool<br>'`
                                         + '   Ford Motor Company<br>'`
                                         + '</span>'
                                     
                    $strNonPDOManagedMessage = '<span style="font-family: Arial;font-size: medium">'`
                                            + '   <b>' + $_.("ows_Managed_x0020_By") + ' Managed - Credential Expiration Notification</b><br>'`
                                            + '</span>'`
                                            + '<span style="font-family: Arial;font-size: x-small">'`
                                            + '	The ' + $_.("ows_Applicaiton_x0020__x0028_ITMS_x0") + ' ' + $_.("ows_ID_x0020_Type") + ' credential <b>' +  $_.("ows_Credential") + $strMessageExpirationText + ' &nbsp;Please '`
                                            + '   ensure that the organization responsible for managing this credential (' + $_.("ows_Managed_x0020_By") + ') is aware of the credential expiration date and that the credential renewal is before the expiration date.<br>'`
                                            + '   <br>'`
                                            + '</span>'`
                                            + '<span style="font-family: Arial;font-size: x-small">'`
                                            + '   <b>Credential Management Procedure:</b><br>'`
                                            + '   <a href="' + $strProcessLInk + '">' + $strProcessLInk + '</a><br>'`
                                            + '   <br>'`
                                            + '   After the credential is managed, use the following link to update the credentials “Expiration Date” field or you will continue to receive reminder emails.<br><br>'`
                                            + '</span>'`
                                            + '<span style="font-family: Arial;font-size: x-small">'`
                                            + '   <b>Credential Management Tool:</b><br>'`
                                            + '   <a href="' + $strDirectIDLink + '">'`
                                            + '   ' + $strDirectIDLink + '</a></span>'`
                                            + '</span>'`
                                            + '<br>'`
                                            + '<br>'`
                                            + '<span style="font-family: Arial;font-size: xx-small">'`
                                            + '   Innovate & Go Further <br>'`
                                            + '   Credential Managment Tool<br>'`
                                            + '   Ford Motor Company<br>'`
                                            + '</span>'
                                        
                    if($_.("ows_Managed_x0020_By") -eq "PDO"){
                        $strMailBody = $strPDOManagedMessage
                    }
                    else{
                        $strMailBody = $strNonPDOManagedMessage
                    }
                


                    if($intExpDays -le ($intDaysSendSupervisorNotification)){
                        $strMailTo = $_.("ows_Manager_x0020__x002f__x0020_Supe") + ";" + $_.("ows_Individuals_x0020_To_x0020_Notif")
                    }
                    else{
                        $strMailTo = $_.("ows_Individuals_x0020_To_x0020_Notif")
                    }
                
                    $arrMailTo = $strMailTo.split(";")
                    $arrMailTo = $arrMailTo | select -uniq
                    $strMailTo = ""
                    foreach ($strEntry in $arrMailTo) {
                        if($strEntry -like '*@*'){
                            $strMailTo += $strEntry.Trim("#"," ") + ";"
                        }
                    }
                    $strMailTo = $strMailTo.Trim(";")

                    if(($strSendNotificationsFor -eq "All Credentials") -or ($_.("ows_Environment") -eq "Production")){
                        if($strEmailMethod -eq "Outlook"){
                            $booReturn = SendOutlookEmail $strMailTo $strMailSubject $strMailBody
                        }
                        else{
                            $booReturn = SendSMTPEmail $strMailTo $strMailSubject $strMailBody
                        }
                    }
                }
            }

            # Reset Email Cadence To Standard Incase It Was Overriden
            $intDaysSendNotification = $intDaysSendNotificationStandard
            $intNotificaitonInterval = $intNotificaitonIntervalStandard
            $intDaysSendSupervisorNotification = $intDaysSendSupervisorNotificationStandard

        }



        if((($strRecieveEmailSummary -eq "Yes") -and ($strEmailSummaryCadence -eq "Monthly") -and ($strCurrentDate.ToString("dd") -eq "01" )) -or (($strRecieveEmailSummary -eq "Yes") -and ($strEmailSummaryCadence -eq "Weekly") -and ($strCurrentDate.ToString("ddd") -eq "Mon"))){

            if($strEmailSummaryCadence -eq "Monthly"){
                $strSummaryContext = "Month";
            }
            else{
                $strSummaryContext = "Week";
            }


            $strFollowUpText = "After Credentials are managed, make sure to update the credential's “"Expiration Date”" field in the Credential Management Tool.<br><br>"
            $strTableFormat = ""

            if($arrSummaryCollection.Count -lt 1){
                $strMessageSummaryText = "There are no Credentials that have expired nor will expire this $strSummaryContext."
                $strFollowUpText =""
            }
            elseif($arrSummaryCollection.Count -eq 1){
                $strMessageSummaryText = "There is 1 Credential that has expired or will expire this $strSummaryContext."
            }
            else{
                $strMessageSummaryText = "There are " + $arrSummaryCollection.Count + " Credentials that have expired or will expire this $strSummaryContext."
            }

            if($arrSummaryCollection.Count -ge 1){

                $strTableFormat = '<b>Credential Details:</b><br>'`
                                + '<table style="font-family: Arial;font-size:12px;border: 1px solid black">'`
                                + '   <tr>'`
                                + '      <th>Application / Service</th>'`
                                + '      <th>Credential</th>'`
                                + '      <th>Environment</th>'`
                                + '      <th>Expiration Date</th>'`
                                + '      <th>Expiration Days</th>'`
                                + '      <th>Change Procedure</th>'`
                                + '      <th>Credential Managemnt Tool Link</th>'`
                                + '   </tr>'

                foreach ($arrCredentialDetail in $arrSummaryCollection) {

                       $strTableFormat += '   <tr>'`
                                        + '      <td>' + $arrCredentialDetail[0] + '</td>'`
                                        + '      <td>' + $arrCredentialDetail[1] + '</td>'`
                                        + '      <td>' + $arrCredentialDetail[6] + '</td>'`
                                        + '      <td>' + $arrCredentialDetail[2].ToString("MM/dd/yyyy") + '</td>'`
                                        + '      <td>' + $arrCredentialDetail[3] + ' Days</td>'`
                                        + '      <td style="text-align: center;"><a href="' + $arrCredentialDetail[4].substring(0,$arrCredentialDetail[4].indexof(",")) + '">Procedure Link</a></td>'`
                                        + '      <td style="text-align: center;"><a href="' + $strSharePointSite + '/Lists/' + $strSharePointListName + '/EditForm.aspx?ID=' + $arrCredentialDetail[5] + '&Source=' + $strSharePointSite + 'CredentialManagementCatalog/CredentialDetails.aspx">Credential Link</a></td>'`
                                        + '   </tr>'
                }
            }
            $strTableFormat += '</table><br>'
    
            $strSummaryMailSubject = "Informational - Credential Management Tool $strEmailSummaryCadence Expiration Summary"
            if($strEmailSummaryCadence -eq "Monthly"){ 
                $strSummaryMailSubject += " (" + $strCurrentDate.ToString("MMMM") + " " + $strCurrentDate.ToString("yyyy") + ")"
            }

            $strSummaryMailBody     = '<style>'`
                                    + '   table {'`
                                    + '        border-collapse: collapse;'`
                                    + '   }'`
                                    + '   table, th, td {'`
                                    + '        border: 1px solid black;'`
                                    + '   }'`
                                    + '   td {'`
                                    + '        padding: 8px;'`
                                    + '       text-align: left;'`
                                    + '   }'`
                                    + '   th {'`
                                    + '        padding: 4px;'`
                                    + '        background-color:#bdbdbd;'`
                                    + '   }'`
                                    + '</style>'`
                                    + '<span style="font-family: Arial;font-size: medium">'`
                                    + '   <b>' + $strEmailSummaryCadence + ' Credential Expiration Summary'
                                
            if($strEmailSummaryCadence -eq "Monthly"){ 
                $strSummaryMailBody += ' For ' + $strCurrentDate.ToString("MMMM") + ' ' + $strCurrentDate.ToString("yyyy")
            }       
                                
            $strSummaryMailBody +=    '</b><br>'`
                                    + '</span>'`
                                    + '<span style="font-family: Arial;font-size: x-small">'`
                                    + $strMessageSummaryText + '<br>'`
                                    + '	<br>'`
                                    + '</span>'`
                                    + '<span style="font-family: Arial;font-size: x-small">'`
                                    + $strTableFormat`
                                    + $strFollowUpText`
                                    + '</span>'`
                                    + '<span style="font-family: Arial;font-size: x-small">'`
                                    + '	<b>Credential Management Tool:</b><br>'`
                                    + '	<a href="' + $strSharePointSite + 'CredentialManagementCatalog/CredentialDetails.aspx">'`
                                    + '	' + $strSharePointSite + 'CredentialManagementCatalog/CredentialDetails.aspx</a></span>'`
                                    + ' </span>'`
                                    + '<br>'`
                                    + '<br>'`
                                    + '<span style="font-family: Arial;font-size: xx-small">'`
                                    + '   Innovate & Go Further <br>'`
                                    + '   Credential Managment Tool<br>'`
                                    + '   Ford Motor Company<br>'`
                                    + '</span>'

            if($strEmailMethod -eq "Outlook"){
                $booReturn = SendOutlookEmail $strEmailSummaryRecipients $strSummaryMailSubject $strSummaryMailBody
            }
            else{
                $booReturn = SendSMTPEmail $strEmailSummaryRecipients $strSummaryMailSubject $strSummaryMailBody
            }
        }
    }

    return $booReturn
}        



######################################################################
# Function UpdateSharePointListEntries
#
# Description: 
# Updates All Entries In A Share Point List With Passed Parameters
######################################################################
function UpdateSharePointListEntries([int] $intListEntryID, [String] $strUpdateValue){
    
    # Create an xmldocument object and construct a batch element and its attributes.
    $xmldoc = new-object system.xml.xmldocument
    
    # note that an empty viewname parameter causes the method to use the default view
    $batchelement = $xmldoc.createelement("Batch")
    $batchelement.setattribute("onerror", "continue")
    $batchelement.setattribute("listversion", "1")
    $batchelement.setattribute("viewname", $strSharePointViewID)
    
    $xml = ""
    $xml += "<Method ID='1' Cmd='Update'>" +
            "<Field Name='ID'>$intListEntryID</Field>" +
            "<Field Name='Expiration_x0020_Days'>$strUpdateValue</Field>" +
            "</Method>"
    
    $batchelement.innerxml = $xml
    $ndreturn = $null
    
    try {            
        $ndreturn = $service.updatelistitems($strSharePointListName, $batchelement)
        $booReturn = "True"
    }            
    catch {             
        ProcessError "Function: LoopThroughSharePointListEntries" "Try: New-WebServiceProxy First $strSharePointWSDLAddress" "Error: $_" $strErrorNotificaitonEmailTo
    } 
    
    return $booReturn
}



######################################################################
# Function SendOutlookEmail                                           
#                                                                     
# Description:                                                        
# Sends Outlook Email Notifications                                   
######################################################################
function SendOutlookEmail([String] $strMailTo, [String] $strMailSubject, [String] $strMailBody, [String] $strWasError){
    
    try{
        $Outlook = New-Object -ComObject Outlook.Application -ErrorAction:'Stop'
        $Mail = $Outlook.CreateItem(0)
        $Mail.To = $strMailTo
        $Mail.Subject = $strMailSubject
        $Mail.HTMLBody = $strMailBody
        $Mail.Save()
#        $Mail.Send()
        
        $booReturn = "True"
    }
    catch{
        if($strWasError -ne ""){
            $strAlosError = " Also: " + ($strWasError -join '-')
        }
        write-host "Process Did Not Complete Successfully" ($strCurrentDate.ToString("MM/dd/yyyy")) " - In Function: SendSMTPEmail, Error: " $_ $strAlosError
        exit
    }
    return $booReturn
}



######################################################################
# Function SendSMTPEmail                                           
#                                                                     
# Description:                                                        
# Sends SMTP Email Notifications                                   
######################################################################
function SendSMTPEmail([String] $strMailTo, [String] $strMailSubject, [String] $strMailBody, [String] $strWasError){

    $arrTo = $strMailTo.split(";")
    $arrFrom = $strErrorNotificaitonEmailTo.split(";")
    $strAlosError = ""

    try{

        foreach ($strEmailAddress in $arrTo) {

            $strSMTPServer = "apprl.azell.com"
            $objMSG = new-object Net.Mail.MailMessage
            $obJSMTP = new-object Net.Mail.SmtpClient($strSMTPServer)

            $objMSG.From = $arrFrom[0]
            $objMSG.ReplyTo = $arrFrom[0]
            $objMSG.To.Add($strEmailAddress)
            $objMSG.Subject = $strMailSubject
            $objMSG.IsBodyHTML = $true
            $objMSG.Body = $strMailBody

            $obJSMTP.Send($objMSG)
        }

        $booReturn = "True"
    }
    catch{
        if($strWasError -ne ""){
            $strAlosError = " Also: " + ($strWasError -join '-')
        }
        write-host "Process Did Not Complete Successfully" ($strCurrentDate.ToString("MM/dd/yyyy")) " - In Function: SendSMTPEmail, Error: " $_ $strAlosError
        exit
    }

    return $booReturn
              
}



######################################################################
# Function ProcessError                                               
#                                                                     
# Description:                                                        
# Sends Outlook Email Notifications If An Error Occurs                
######################################################################
function ProcessError([String] $strFunction, [String] $strTry, [String] $strSysError, [String] $strMailTo){
    
    $strLongProgName = $MyInvocation.ScriptName
    $strShortProgName = $MyInvocation.ScriptName.substring($MyInvocation.ScriptName.LastIndexOf("\")+1)
    
    $strMailSubject       = 'Action Required - An Error Has Occured While Processing ' + $strShortProgName + ' for the Credential Management Tool'
    
    $strErrorEmailMessage = '<span style="font-family: Arial;font-size: medium">'`
                         + '   <b>PowerShell Error Notificiation - ' + $strShortProgName + '</b>'`
                         + '</span>'`
                         + '<span style="font-family: Arial;font-size: x-small"><br>'`
                         + '   An Error occured while processing "' + $strLongProgName + '". &nbsp;Please resolve this issue and reprocess the program run manually.<br>'`
                         + '   <br>'`
                         + '   <b>Error:</b><br> ' + $strFunction + '<br>' + $strTry + '<br>' + $strSysError + '<br>'`
                         + '</span>'`
                         + '<br>'`
                         + ' <span style="font-family: Arial;font-size: xx-small">'`
                         + '	Innovate & Go Further <br>'`
                         + '	Credential Managment Tool<br>'`
                         + '	Ford Motor Company<br>'`
                         + '</span>'

    $arrErrorDetails = ($strFunction,$strTry,$strSysError)

    if($strEmailMethod -eq "Outlook"){
        $booReturn = SendOutlookEmail $strErrorNotificaitonEmailTo $strMailSubject $strErrorEmailMessage $arrErrorDetails
    }
    else{
        $booReturn = SendSMTPEmail $strErrorNotificaitonEmailTo $strMailSubject $strErrorEmailMessage $arrErrorDetails
    }

    write-host "Process Did Not Complete Successfully" $strCurrentDate.ToString("MM/dd/yyyy") " - " $strFunction ", " $strTry ", " $strSysError
    exit
}



# Call Main To Kick Off Program
Main



