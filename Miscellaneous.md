---
inclusion: auto
name: Miscellaneous
description: Use this file if no information is found in any other steering files
---
<!------------------------------------------------------------------------------------
   Add rules to this file or a short description and have Kiro refine them for you.
   
   Learn about inclusion modes: https://kiro.dev/docs/steering/#inclusion-modes
-------------------------------------------------------------------------------------> 

# Other Scripting Functions

## RemoveCustomTable

You can use this function for Optiva Workflows and Copy Methods.

### Purpose
Removes all rows from an extension table for an object.

### Syntax
```vbscript
Dim variable As String = RemoveCustomTable(DataTableName[, Symbol, Object])
```

### Arguments
<table>
  <thead>
    <tr>
      <th>Part</th>
      <th>Description</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <td>DataTableName</td>
      <td>Name of table in the <b>Extension Table Definition</b> form.</td>
    </tr>
    <tr>
      <td>Symbol</td>
      <td>Type of object. Use empty quotation marks to indicate the symbol for the workflow or the copy method.</td>
    </tr>
    <tr>
      <td>Object</td>
      <td>Object for which to remove the extension table rows. Use empty quotation marks to indicate the current object for the workflow, or the new object for the copy method.</td>
    </tr>
  </tbody>
</table>

### Description
Use this function to delete all rows from an extension table for an object. Use for tables where users can add rows.

### Examples
This example removes all rows from CustomTable1 for the item 10001.
```vbscript
Dim sRemoveCustomTable As String = RemoveCustomTable("CustomTable1", "ITEM", "10001")
```
This example removes all rows from CustomTable1 for a new formula. The new formula is being created by a copy method.
```vbscript
Dim sRemoveCustomTable As String = RemoveCustomTable("CustomTable1")
```

## RemoveDocuments

You can use this function for Optiva workflows.

### Purpose

When RemoveNotification is called from any action, it will remove the identified notification from the field identified.

**Note:** There are optional parameters for symbol, object and detailCode. rowKey is reserved for future use.

### Syntax

```vbnet
Public Function RemoveNotification(propertyName As String, Optional objSymbol As String = "", Optional objectKey As String = "", Optional detailCode as String = "", Optional rowkey As String = "") as Long
```

### Arguments

<table>
  <tr>
    <th>Argument</th>
    <th>Description</th>
  </tr>
  <tr>
    <td>propertyName</td>
    <td>Holds the property name or field name</td>
  </tr>
  <tr>
    <td>Optional objSymbol</td>
    <td>The data object type or the current object type if blank.</td>
  </tr>
  <tr>
    <td>Optional objectKey</td>
    <td>The data object key or the current object key if blank.</td>
  </tr>
  <tr>
    <td>detailCode</td>
    <td>The system detail code that is used to return the new row.</td>
  </tr>
  <tr>
    <td>Optional rowKey</td>
    <td>This is reserved for future use.</td>
  </tr>
</table>

### Remove Test Example

The following line of code will remove this specific Notification from the current object.

```vbnet
RemoveNotification("DESCRIPTION")
Return 111
```


## RenameFormula

You can use this function for Optiva Workflows.

### Purpose

Assigns a new name and version to an object.

### Syntax

```vbscript
Dim variable As Object = RenameFormula(NewCode, NewVersion)
```

### Arguments

<table>
<thead>
<tr>
<th>Part</th>
<th>Description</th>
</tr>
</thead>
<tbody>
<tr>
<td>NewCode</td>
<td>New name for the formula.</td>
</tr>
<tr>
<td>NewVersion</td>
<td>New version for the formula.</td>
</tr>
</tbody>
</table>

### Description

RenameFormula assigns a new name and version to a formula. It replaces the name of the current formula. It is not a copy of the formula. This function refreshes on-screen data.

### Examples

In this example, a new name and version number are assigned to the current workflow formula.

```vbscript
Dim oNewname As Object = RenameFormula ("PIZZA_SAUCE", "004")
```

```vbnet
Dim oIncrementformula As Object = RenameFormula("","+")
```


## RenameObject

You can use this function for Optiva Workflows.

### Purpose

Assigns a new name and version to an object. Currently, this function applies to formulas, label content, or specifications.

### Syntax

```vbnet
Dim variable As Long = RenameObject(Symbol, OldCode, NewCode)
```

### Arguments

<table>
  <thead>
    <tr>
      <th>Part</th>
      <th>Description</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <td>Symbol</td>
      <td>Type of object. Use empty quotation marks to indicate the symbol of the object for the workflow.</td>
    </tr>
    <tr>
      <td>OldCode</td>
      <td>Current object code.</td>
    </tr>
    <tr>
      <td>NewCode</td>
      <td>New code.</td>
    </tr>
  </tbody>
</table>

### Description

RenameObject assigns a new name and version to a formula, label content, or specification. It replaces the name of the object, but does not make a copy of the object. This function refreshes on-screen data.

### Examples

In this example, a new name and version number are assigned to the current workflow specification.

```vbscript
Dim lNewname As Long = RenameObject("", "", "LOWFAT\0001")
```

In this example, a specification is identified.

```vbscript
Dim lNewname As Long
RenameObject("SPECIFICATION", "FAT\0001", "LOWFAT\0001")
```



## SaveReportToIDM

You can use this function for Optiva Workflows.

### Purpose

Use this function to generate and save a report document in IDM for any Optiva symbol.

### Syntax

```vba
Dim variable As Boolean = SaveReportToIDM(reportUrl, jobFileName,
objectSymbol, reportType, reportPath, outputPath, debugMode,
IDMdocCode, fileName, IDMattributes)
```

### Arguments

<table>
  <thead>
    <tr>
      <th>Parameter</th>
      <th>Description</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <td>reportUrl</td>
      <td>Optiva WebReport Url. For example: http://<machine_name>/FsWebReports/Launch.aspx<br>The machine_name is the web server that is hosting WebReports.</td>
    </tr>
    <tr>
      <td>jobFileName</td>
      <td>The path and the XML file that contains the data to report.</td>
    </tr>
    <tr>
      <td>objectSymbol</td>
      <td>Optiva object symbol. For example, Formula or Item.</td>
    </tr>
    <tr>
      <td>reportType</td>
      <td>Either PDF for Adobe Acrobat, Excel, MS Word or HTML.</td>
    </tr>
    <tr>
      <td>reportPath</td>
      <td>Crystal Reports template; the path and file name of the Crystal Reports RPT file to use.</td>
    </tr>
    <tr>
      <td>outputFilePath</td>
      <td>The path and file name to which the report file will copy from a temporary report path.</td>
    </tr>
    <tr>
      <td>debugMode</td>
      <td>To write a trace log. If “TRACE” is provided, a log file will be created. For example, “TRACE” or empty string.</td>
    </tr>
    <tr>
      <td>IDMdocCode</td>
      <td>An IDM Doc Code to which the report will be attached.</td>
    </tr>
    <tr>
      <td>fileName</td>
      <td>Name of the report file for the File_Name IDM attribute. If an empty string is provided, the temporary report file name that is generated will be set to the “File_Name” IDM attribute.</td>
    </tr>
    <tr>
      <td>IDMattributes</td>
      <td>All required attributes (i.e., OPTIVA_SYMBOL, OPTIVA_SYMBOL_ID, OPTIVA_DOC_TITLE) and any additional attributes like expiry_date for an IDM doc type.</td>
    </tr>
  </tbody>
</table>

### Example

This example saves the report file in IDM for a formula object symbol.

```vbnet
Option Strict Off
Imports System
Imports System.Collections.Generic
Imports System.Diagnostics
Imports DocICP = Infor.DocumentManagement.ICP

Class ActionScript
    Inherits FcProcFuncSetEventWF

    Function wf_start() As Long

        Dim reportURL As String = ""
        Dim jobfileName As String = ""
        Dim objectSymbol As String = ""
        Dim reporttype As String = ""
        Dim docfilepath As String = ""
        Dim IDMdocCode As String
        Dim fileName As String
        Dim reportPath As String
        dim outputPath as String
        dim debugMode as string
        Dim IDMattributes As Dictionary(Of String, Object)

        Try

            reportURL = "http://localhost/FsWebReports/launch.aspx"
            jobFileName = "\SERVERNAME\FsWebReports\Source\ReportTemplate.xml"
            objectSymbol = "FORMULA"
            reporttype = "pdf"
            reportPath = "\SERVERNAME\FsWebReports\Source\FsCRMSDS01.rpt"
            filename = "Formula_Detail_RPT." & reporttype
            outputPath = ""
            debugMode = "TRACE"
            IDMdocCode = "Report_Attach_Test"

            ' IDM Attributes
            IDMattributes = New Dictionary(Of String, Object)
            IDMattributes.Add("OPTIVA_SYMBOL", objectSymbol)
            IDMattributes.Add("OPTIVA_SYMBOL_ID", "01010")
            IDMattributes.Add("OPTIVA_SYMBOL_VERSION", "0001")
            IDMattributes.Add("OPTIVA_DOC_TITLE", "A Test-" & DateTime.Now.ToString().Trim())

            Me.SaveReportToIDM(reportUrl, jobFileName, objectSymbol, reporttype,
                reportPath, outputPath, debugMode, IDMdocCode, fileName, IDMattributes)

        Catch ex As exception
            Messagelist("error : " & ex.Message)
        End Try
        Return 111

    End Function

End Class
```

## Serialize

You can use this function to convert object into JSON format.

### Syntax

```csharp
Dim keyValuePairs As Dictionary(Of String, object) = New Dictionary(Of String, object)
Dim jsonString as String = Serialize(keyValuePairs)
```

### Arguments

<table>
<thead>
<tr>
<th>Part</th>
<th>Description</th>
</tr>
</thead>
<tbody>
<tr>
<td>String</td>
<td>The value of the data type.</td>
</tr>
<tr>
<td>object</td>
<td>The object type.</td>
</tr>
</tbody>
</table>

### Description
Serialize function converts the specified object into JSON format.

### Example
This is an example to serialize object into JSON string.
A dictionary is created to store the object data.

```vbnet
Dim keyValuePairs As Dictionary(Of String, object) = New Dictionary(Of String, object)
```

The data is stored into the newly created dictionary.

```vbnet
keyValuePairs.Add("Project Code", "P0001")
keyValuePairs.Add("Description", "Project Description")
keyValuePairs.Add("Customer Name", "Customer Name")
keyValuePairs.Add("Total Cost", 1234.52)
```

Using the serialize function, the object data is converted into a JSON string.

```vbnet
Dim jsonString as String = Serialize(keyValuePairs)
```

## ShellAPI

You can use this function for Optiva Workflows.

### Purpose

Performs an action for a registered application. For example, opening a Microsoft WORD file or displaying a file in a browser.

### Syntax

```vbnet
Dim variable As Integer = ShellAPI(Target)
```

### Arguments

<table>
<thead>
<tr>
<th>Part</th>
<th>Type</th>
<th>Description</th>
</tr>
</thead>
<tbody>
<tr>
<td>Target</td>
<td>String</td>
<td>Path to the file or URL. If the target is an executable, then the parameters are passed to the executable.<br>Parameters vary according to the program.</td>
</tr>
</tbody>
</table>

### Examples

In this example, the Yahoo home page opens.

```vbnet
Dim iOpen As Integer = ShellAPI("http://www.yahoo.com")
```

This example opens a Microsoft Word file; Open is the default action and no other action is specified.

```vbnet
Dim iReturncode As Integer
ShellAPI ("C:\Database Files\DBChanges.doc")
```

You can use this function for Optiva Workflows.

### Purpose

ShellAPI2 performs an action for a registered application. For example, opening a Microsoft WORD file or displaying a file in a browser.

### Syntax

```vbnet
Dim variable As Integer = ShellAPI2(Target)
```

### Arguments

<table>
  <thead>
    <tr>
      <th>Part</th>
      <th>Type</th>
      <th>Description</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <td>Target</td>
      <td>String</td>
      <td>Path to the file or URL. If the target is an executable, then the parameters are passed to the executable.<br>Parameters vary according to the program.</td>
    </tr>
  </tbody>
</table>

### Example

The below is the example of the script used to create ShellAPI2.

```vbnet
Option Strict Off
imports System
imports System.Diagnostics
Class ActionScript
    Inherits FcProcFuncSetEventWF
    Function wf_start() As Long
        Workflow.ShellAPI("http://usfrvqas19c/FsWebReports/launch.aspx?p1=BoatInfo2008.xml&p2=asdf&p3=PDF&p4=C:\inetpub\wwwroot\FsWebReports\Source\BoatInfo2008.rpt")
    End Function
End Class
```

## StartForm

You can use this function for Optiva Workflows and Equations.

### Purpose

Opens an Optiva form.

### Syntax

```vbnet
Dim variable As Long = StartForm(FormName[, Object, Arguments ])
```

### Arguments

<table>
  <thead>
    <tr>
      <th>Part</th>
      <th>Description</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <td>FormName</td>
      <td>Optiva name for a form. In Optiva, the form name begins with <b>frm</b>.</td>
    </tr>
    <tr>
      <td>Object</td>
      <td>Optional. Object key to be opened in the form. Do not include an object code here, if you want the object for the workflow to open in the form.</td>
    </tr>
    <tr>
      <td>&lt;arguments&gt;</td>
      <td>Optional. String.<br>Depends on the form being opened. Most forms do not require additional arguments.<br>When you open the Compare form, add a semicolon-delimited list of string values. The order of the values is:<ul><li>Symbol of the base compare object</li><li>Object (key) of the base compare object</li><li>Symbol of the objects in the compare list</li><li>One or more objects (keys) to include in the compare list</li></ul>Use the <b>VERIFY.FORMULA_BTN</b> profile attribute to specify a workflow to use for opening the <b>Compare</b> form from the <b>Verify</b> button in the <b>Formula</b> form. See the <i>Infor Optiva Application Configuration Guide</i>.<br>When you open the <b>Project Management</b> form, add a semi colon-delimited list of string values. The order of the values is:<ul><li>WipId of the action wip that you want to manage.</li><li>The symbol of the object that you want to manage.</li></ul>If you do not specify a symbol, PROJECT is the default symbol.</td>
    </tr>
  </tbody>
</table>

### Description

StartForm opens a form. Use it in the **VIEW** and **EDIT** events. Use multiple statements to open more than one form at a time.

### Item and formula examples

In this example, the Item form for TOMATOES opens.

```vbnet
Dim lOpen As Long = StartForm("frmitem", "TOMATOES")
```

In the next example, the **Formula** form for the current formula opens.

```vbnet
Dim lOpen As Long = StartForm("frmformula")
```

In this example, formula FS-0008\0002 is the base formula that is compared to these specifications: FS-0008-FOR\0002 and FS-0008-PKG\0002.

```vbscript
Dim lOpen As Long = StartForm("FRMCOMPARE", "", "FORMULA;FS-0008\0002;SPECIFICATION;FS-0008-FOR\0002;FS-0008-PKG\0002")
```

Two action sets are available in the Optiva seed database in the **Lookup by Sets** lookup. You can use them to open the **Compare** form.

You can open the form using the **Verify** button in the **Formula** toolbar.

*   **VERIFY_FMLA_SPEC** (i.e., the default).
    Any specifications on the **References** tab of the open formula are automatically populated into the Compare List of the **Compare** form.
*   **VERIFY_FMLA_PARENT**
    Determines whether the open formula has a parent formula. If it does, the parent is set as the base formula and the open formula as the first line in the Compare List. If the open formula has no parent, then this action set behaves the same as the **VERIFY_FMLA_SPEC** action set.

### Advanced Workflow Management

You can open the **Project Management** form for any business object that supports stage gate action set workflows.

*   To display the **Project Management** form for the Project symbol, specify the name of the object. You can also specify the name of the object and the WIP ID. For these examples, the name of the object is Project1 and the WIP ID is 2237.

```vbscript
Dim retVal as Long = StartForm("FRMPROJECTMGMT", "PROJECT1")
Dim retVal as Long = StartForm("FRMPROJECTMGMT", "PROJECT1", "2237")
```

*   To display the **Project Management** form for a symbol other than Project, you must specify the symbol itself as the final argument. Add a semi-colon before the symbol. The WIP ID is optional.
    In this example, the **Project Management** form displays the Item that is named Item 1, along with WIP ID 2242.

```vbscript
Dim retVal as Long = StartForm("FRMPROJECTMGMT", "ITEM1", "2242;ITEM")
```

To show the most recent WIP ID, omit the WIP ID from the argument. Remember to add the semi-colon before the symbol.

```vbscript
Dim retVal as Long = StartForm("FRMPROJECTMGMT", "ITEM1", ";ITEM")
```

# Scripting functions for IDM Reports

Infor Document Management (IDM) Output Reports are available to multi-tenant PLM for Process customers. See the *Infor PLM for Process Reports Administration Guide* for general information about report setup and libraries available to multi-tenant customers to support IDM Output Reports.

Two new scripting classes and a function have been added to PLM for Process to assist in generating IDM Output Reports. These require Secure Scripting to be enabled on the server, which is standard in SAAS environments.

## CreateIDMReport

This function is used to generate or create reports and to add the reports to IDM for a staging doc code. It does not launch the report for the user to view. See the Infor PLM for Process Reports Administration guide for information about script libraries available to perform that function.

It takes IDMReportParams object as Input and gives IDMReportResult Object as Output.

Here is an example script:

```vbscript
Dim reportParams as New IDMReportParams
Dim reportResult as New IDMReportResult

reportParams.xmlData = sXML
reportParams.templateFileName = "Label Content Summary Report"
reportParams.targetFileName = "Label Content Summary Report.pdf"
reportResult = CreateIDMReport(reportParams)
```

## IDMReportParams

IDMReportParams is passed as input to CreateIDMReport function. The IDMReportParams properties are as shown in table below. XmlData, templateFileName, targetFileName are required properties that cannot be left null.

<table>
<thead>
   <tr>
      <th>Properties</th>
      <th>Description</th>
   </tr>
</thead>
<tbody>
   <tr>
      <td>xmlData</td>
      <td>XML Data of Object as String</td>
   </tr>
   <tr>
      <td>templateFileName</td>
      <td>The Report Template name in IDM.</td>
   </tr>
   <tr>
      <td>targetFileName</td>
      <td>Denotes the generated report name.</td>
   </tr>
    <tr>
      <td>templateDocCode (Optional)</td>
      <td>IDM Document Type which contains Template Word File.<br><br>By default, a profile attribute IDM.REPORT_TEMPLATE DOCODE is added, the function will read value from profile. User can configure any IDM Doc Type as Template Doc Code and assign it to a profile attribute in Optiva or they can give as argument as shown below. A PLM for Process Function Code of the same name is optional. Templates are not attached to PLM for Process objects.<br><br>Example:<br><br>```reportParams.templateDocCode = "OPTIVA_REPORT_TEMPLATES"```</td>
    </tr>
    <tr>
      <td>sendEmail (Optional)</td>
      <td>It is a Boolean type. If user set sendEmail to true, the report generated will be sent to the given email address.<br><br>Example:<br><br>```reportParams.sendEmail = True```</td>
    </tr>
    <tr>
      <td>targetEmailAddresses (Optional)</td>
      <td>Email address where report should be sent, if sendEmail is set to true.<br><br>It will support multiple addresses by passing it as a comma separated.<br><br>Example:<br><br>```reportParams.targetEmailAddresses = "Firstname1.Lastname1@infor.com; Firstname2.Lastname2@infor.com"```</td>
    </tr>
    <tr>
      <td>ReturnFileContent (Optional)</td>
      <td>It is Boolean type, and you can specify whether you want the file content as byte array in the response or not.<br><br>Example:<br><br>```reportParams.ReturnFileContent = True```</td>
    </tr>
  <tr>
    <td>TargetDocCode (Optional)</td>
    <td>IDM Doc Type to which generated report should be attached.<br>
By default, a profile attribute IDM.REPORT_STAGINGDOC_CODE is added, and the function will read the value from the profile. However, you can configure any IDM Doc Type as Target Doc Code and decide to use different Doc Types for different objects or reports. In most cases, you will also create a PLM for Process function code with the same name for visibility inside PLM for Process. Use of this function will also require the next section of code, creating a CMItem.<br>
Example:<br>
reportParams.targetDocCode = "OPTIVA_STAGING_DOC"</td>
  </tr>
  <tr>
    <td>targetCMItem<br>(Optional, used with TargetDocCode)</td>
    <td>If you specify targetDocCode, then the user should create a CMItem and should pass it to function with all required attributes.<br>
Example:<br>
Dim targetCMItem As DocICP.CMItem = New DocICP.CMItem()<br>
targetCMItem.EntityName = targetDocCode<br>
targetCMItem.SetAttributeValue("OPTIVA_SYMBOL", _OBJECTSYMBOL, DocICP.DataType.String)<br>
objKey = _OBJECTKEY.Split("\")<br>
targetCMItem.SetAttributeValue("OPTIVA_SYMBOL_ID", objKey(0), DocICP.DataType.String)<br>
If objKey.Length > 1 Then targetCMItem.SetAttributeValue("OPTIVA_SYMBOL_VERSION", objKey(1), DocICP.DataType.String)<br>
reportParams.targetCMItem = targetCMItem</td>
  </tr>
</table>

## IDMReportResult

CreateIDMReport function returns the result as an IDMReportResult class. The IDMReportResult properties are as shown in the below table, and can be used in alert messages for debugging issues.

<table>
  <tr>
    <th>Properties</th>
    <th>Description</th>
  </tr>
  <tr>
    <td>IDMDoc_ITEM_ID</td>
    <td>ITEMID of IDM Document which contains generated report.<br>
User can search in IDM in the Target Doc Code using ITEMID Attribute to get generated report.</td>
  </tr>
    <tr>
      <td>IDMDoc_URL</td>
      <td>Doc URL of IDM Document which contains generated report.<br>User can use that URL to redirect to IDM Document directly.</td>
    </tr>
    <tr>
      <td>IDMDoc_JobId</td>
      <td>Job ID is used to see the status of the job. User can see the status of the job, error messages and data related to job in Job Management in IDM.</td>
    </tr>
    <tr>
      <td>FileData</td>
      <td>FileData will have Report File as bytes array. It only applicable if ReturnFileContent is set to true.</td>
    </tr>
    <tr>
      <td>ErrorMessage</td>
      <td>ErrorMessage will have an error message in case of failure.</td>
    </tr>
  </tbody>
</table>

Example debugging messages:

```javascript
messagelist(reportResult.ErrorMessage)

messagelist("ITEMID - " & reportResult.IDMDoc_ITEM_ID)

messagelist("IDM Document URL - " & reportResult.IDMDoc_URL)

messagelist("Job ID - " & reportResult.IDMDoc_JobId)
```


