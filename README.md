# Export-ADComputerPropertyFrequencies_To-Html.ps1

--- 

### SYNOPSIS
Analyze all Active Directory computers by frequency of assigned properties. <br>
Export Analysis results to HTML report. <br>
Find out how many AD Computers are assigned a specific property. <br>
Find out how many unique property assignments exist for a specific property. 

### DESCRIPTION
Analyze all Active Directory computers by frequency of assigned properties. <br>
Export Analysis results to HTML report. 

### PARAMETER Server
Specifies the Active Directory server domain to query.

### PARAMETER Enabled
Specifies to only return results from enabled computers.

### EXAMPLE
PS C:\> Export-ADComputerPropertyFrequencies_To-Html.ps1 -Server CyberCondor.local -Enabled $true

## Input
Template.html
- Template for HTML layout 
styles.css
- Styles to display HTML layout properly
Computers
- Contains list of all computers and their properties

## Output
~\_FrequenyAnalysisReports\ADComputerPropertiesFrequenyAnalysisReport_$Date
- Computers.html
- styles.css

---

# Details

This PowerShell script analyzes the properties of Active Directory (AD) computers and generates a report in HTML format.

The script starts by specifying two mandatory parameters: `$Server` and `$Enabled`. The `$Server` parameter specifies the domain of the Active Directory server to be queried, while the `$Enabled` parameter specifies whether to only return results from enabled computers.

The script first performs a check to see if the required AD module is installed and that the specified server can be reached. If either of these conditions is not met, the script will output a warning message indicating the cause of the issue.

The script then defines several functions to carry out various tasks:

-   `Get-ExistingUsers_AD` retrieves the list of existing AD users and returns it.
    
-   `Get-UserRunningThisProgram` finds the user who is currently running the script, based on the `$env:UserName` environment variable.
    
-   `Get-ExistingComputers_AD` retrieves the list of existing AD computers and returns it.
    
-   `Create-NewHtmlReportFromTemplate` creates a new HTML report based on a specified template file. The template file, a CSS stylesheet, and a list of computers with their properties, are the inputs to this function. The function outputs a new HTML report file with the specified name in a specified folder.

The script concludes by calling the `Create-NewHtmlReportFromTemplate` function to generate the final report.

---

# Get-PropertyFrequencies

This function, named "Get-PropertyFrequencies", is a PowerShell script that takes two input arguments: `$Property` and `$Object`. The purpose of the function is to determine the frequency of each unique value of a specified property within the given object.

The first steps of the function involve initializing variables and obtaining a list of all unique values for the specified property within the object. If the property is of type "DateTime", the function sets a flag `$isDate` to `$true`.

Next, the function loops through each unique value, creating a new object for each value with properties "Property" (the unique value), "Count" (initialized to 0), and "Frequency" (initialized to "100%").

The function then determines if the property is of type "DateTime". If it is, the function formats the unique values as "yyyy-MM" strings.

Afterwards, the function loops through each value of the specified property in the object. If the value is equal to a unique value, the corresponding "Count" property of the object is incremented. If the property is of type "DateTime", the function also checks for equality after converting both the value and the unique value to strings in the "yyyy-MM" format.

At the end, the function computes the frequency for each unique value as the count divided by the total number of values and returns an array of objects sorted by count and property.

Note: The function also includes a "Write-Progress" command which outputs a progress bar during the execution of the script.
