Param(
 
	[Parameter(Mandatory=$True)]
     [string]$VeracodeUsersExcelFile,

	[Parameter(Mandatory=$True)]
     [string]$SheetName,

    [Parameter(Mandatory=$True)]
     [string]$UsersXmlFile,
	 
	[Parameter(Mandatory=$True)]
     [string]$VeraCodeJarFile,
	 
	[Parameter(Mandatory=$True)]
     [string]$VeraCodeID,
	 
	[Parameter(Mandatory=$True)]
     [string]$VeraCodeSecretKey,
	 
	[Parameter(Mandatory=$True)]
     [string]$VeraCodeAction
 
)

# Input Files Existance Validation
Function Validate-InputFilesExist($VeracodeUsersExcelFile, $UsersXmlFile) {

    # The below args usef this function only!!!
    $VUEF = Test-Path $VeracodeUsersExcelFile
    $UXF = Test-Path $UsersXmlFile
    $VCF = Test-Path $VeraCodeJarFile

	if ( (Test-Path $VeracodeUsersExcelFile) -and (Test-Path $UsersXmlFile) -and (Test-Path $VeraCodeJarFile)  ) {
	
		Write-Output "Veracode Users Excel File Validation $VUEF"
        Write-Output "Users Xml File Validation $UXF"
        Write-Output "Vera Code Jar File Validation $VCF"

	
	} else {

		Write-Output "Invalid Input, Check Below :False are Validation Error(s)"
		Write-Output "Veracode Users Excel File Validation $VUEF"
        Write-Output "Users Xml File Validation $UXF"
        Write-Output "Vera Code Jar File Validation $VCF"
        exit 0
		
	}

}

# Open Excel File
Function Open-ExcelFile($ExcelFile) {

    $invokeExcel = New-Object -ComObject Excel.Application
    $ExcelWorkBook = $invokeExcel.Workbooks.Open("$ExcelFile")
    Return $ExcelWorkBook

}

# Count All Users From Excel File in Username Column 
Function Get-UsersCountFromExcel($ExcelFile, $ExcelSheet) {

    $ExcelWorkBook = Open-ExcelFile $ExcelFile
    $ExcelSheet = $ExcelWorkBook.sheets.item("$ExcelSheet")
    $RowsWithData = ($ExcelSheet.UsedRange.Rows).count
	# Close Excel File
    $ExcelWorkBook.Close()
    Return $RowsWithData

}

# Get Veracode Usernames From Xml File
Function Get-UserNamesFromXmlFile($XmlFile) {

    # Get Data From Xml File
	$XmlFile = resolve-path("$XmlFile")
    $Xmldoc = new-object System.Xml.XmlDocument
    $Xmldoc.load($XmlFile)
    $Xmldoc = [xml] (get-content “$XmlFile”)

    # Create array with all UserNames or Emails
    $UserNames = $Xmldoc.userlist.users.usernames
    $UserNames_Array = $UserNames.split(",")

    # Initialize Array For Modified Emails
    $Modified_UserNames_Array = New-Object System.Collections.ArrayList

    $count = $UserNames_Array.Count
    
    # Modifing Emails And Adding To Modified Emails Array
    while ( $count -gt 0 ) {

        $count--
        $UserName = $UserNames_Array[$count].Replace("http://www.okta.com/exk540niaapnezpdt1t7", "")
        [void] $Modified_UserNames_Array.Add("$UserName")
        
    }
	
    # Return Modified Emils Array
    Return $Modified_UserNames_Array

}

# Get New Users From Excel File
Function Get-UsersDetailsFromExcelFile($ExcelFile, $ExcelSheet) {

    #Load Column Numbers to variable
    $UserName_Col = 1
	$FirstName_Col = 2
    $LastName_Col = 3
    $Team_Col = 4
    $AccountType_Col = 5
    $LoginEnabled_Col = 6
    $Role_Col = 7

    # Get rows count
	$Rows = Get-UsersCountFromExcel $ExcelFile $ExcelSheet
    
	# Create an hashtable variable for all values
    [hashtable]$ReturnUserDetails = @{}
    
    # Open Excel File load the sheet with data
	$ExcelWorkBook = Open-ExcelFile $ExcelFile $ExcelSheet
    $WorkSheet = $ExcelWorkBook.sheets.item($ExcelSheet)

    # Initialize arrays for each column item
    $UserNames_Array = New-Object System.Collections.ArrayList
	$FirstNames_Array = New-Object System.Collections.ArrayList
	$LastNames_Array = New-Object System.Collections.ArrayList
    $Teams_Array = New-Object System.Collections.ArrayList
    $AccountTypes_Array = New-Object System.Collections.ArrayList
    $LoginEnabled_Array = New-Object System.Collections.ArrayList
    $Roles_Array = New-Object System.Collections.ArrayList

	# Retrive and Add Values to Arrays
    while($Rows -gt 1) {

       $UserName = $WorkSheet.Rows.Item($Rows).Columns.Item($UserName_Col).Text
	   $FirstName = $WorkSheet.Rows.Item($Rows).Columns.Item($FirstName_Col).Text
       $LastName = $WorkSheet.Rows.Item($Rows).Columns.Item($LastName_Col).Text
       $Team = $WorkSheet.Rows.Item($Rows).Columns.Item($Team_Col).Text
       $AccountType = $WorkSheet.Rows.Item($Rows).Columns.Item($AccountType_Col).Text
       $LoginEnabled = $WorkSheet.Rows.Item($Rows).Columns.Item($LoginEnabled_Col).Text
       $Role = $WorkSheet.Rows.Item($Rows).Columns.Item($Role_Col).Text
       [void] $UserNames_Array.Add("$UserName")
       [void] $FirstNames_Array.Add("$FirstName")
       [void] $LastNames_Array.Add("$LastName")
       [void] $Teams_Array.Add("$Team")
       [void] $AccountTypes_Array.Add("$AccountType")
       [void] $LoginEnabled_Array.Add("$LoginEnabled")
       [void] $Roles_Array.Add("$Role")
       $Rows--

    }  
   
    #Assign all return Arrays in to hashtable
    $ReturnUserDetails.UserNames_Array = $UserNames_Array
    $ReturnUserDetails.FirstNames_Array = $FirstNames_Array
    $ReturnUserDetails.LastNames_Array = $LastNames_Array
    $ReturnUserDetails.Teams_Array = $Teams_Array
    $ReturnUserDetails.AccountTypes_Array = $AccountTypes_Array
    $ReturnUserDetails.LoginEnabled_Array = $LoginEnabled_Array
    $ReturnUserDetails.Roles_Array = $Roles_Array
	
	# Close Excel File
	$ExcelWorkBook.Close()
	
    #Return the hashtable
    Return $ReturnUserDetails
}

# Compare Value Functions of Excel and Xml files, for getting New Users to create!
Function Create-NewUsers($ExcelFile, $ExcelSheet, $XmlFile) {

	# Get Users From Excel Function
    $NewUsersDetailsFromExcel = Get-UsersDetailsFromExcelFile $ExcelFile $ExcelSheet
	# Get Users From Xml Function
    $UsersFromXml = Get-UserNamesFromXmlFile $XmlFile
    # Initialize an Array For New Users
	$NewUsers_Array = New-Object System.Collections.ArrayList
	
	Write-Output "*************************************************************"
	$count = $NewUsersDetailsFromExcel.UserNames_Array.Count
    while ( $count -gt -1 ) {
        		
		$UserName = $NewUsersDetailsFromExcel.UserNames_Array[$count]
        $FirstName = $NewUsersDetailsFromExcel.FirstNames_Array[$count]
        $LastName = $NewUsersDetailsFromExcel.LastNames_Array[$count]
        $Team = $NewUsersDetailsFromExcel.Teams_Array[$count]
        $AccountType = $NewUsersDetailsFromExcel.AccountTypes_Array[$count]
        $LoginEnabled = $NewUsersDetailsFromExcel.LoginEnabled_Array[$count]
        $Role = $NewUsersDetailsFromExcel.Roles_Array[$count]

        $CreateUserCommand = "java -jar $VeraCodeJarFile -vid $VeraCodeID -vkey $VeraCodeSecretKey -action $VeraCodeAction -firstname $FirstName -lastname $LastName -emailaddress $UserName -loginenabled $LoginEnabled -loginaccounttype $AccountType -teams $Team -roles $Role"
        #$CreateUserCommand = "echo $CreateUserCommand"
        if ($UsersFromXml -notcontains $UserName) { 
            
            if ( $UserName ) {
                
				IEX $CreateUserCommand
				[void] $NewUsers_Array.Add("$UserName")                               
                DD "FirstName	  :	" $FirstName
                DD "LastName	  :	" $LastName
				DD "Emails		  :	" $UserName 
                DD "Team		  :	" $Team
                DD "AccountType   :	" $AccountType
                DD "LoginEnabled  :	" $LoginEnabled
                DD "Role		  :	" $Role
				
                Write-Output "==============================> NEW USER <=============================="
                Write-Output "========================================================================="
        
            } 

        } 
		
        $count--
		
    }	
	
	# Collect and display new users
    if ( $NewUsers_Array.Count -eq 0 ) {

        Write-Output "========================================================================="
        Write-Output "No New Users Found! Created Users Count " $NewUsers_Array.Count
        Write-Output "========================================================================="

    } else {

        $NewUsersCount = $NewUsers_Array.Count
        Write-Output "New Users Created! Count is $NewUsersCount"
        Write-Output "########## New Users ##########"
        Write-Output $NewUsers_Array
        Write-Output "###############################"
        
    }

    Write-Output "*************************************************************"
    [System.GC]::Collect()

}

# Display Method
Function DD($Text, $Val) {

	Write-Host $Text $Val
	
}
#Clear-Host
Validate-InputFilesExist $VeracodeUsersExcelFile $UsersXmlFile
Create-NewUsers $VeracodeUsersExcelFile $SheetName $UsersXmlFile
