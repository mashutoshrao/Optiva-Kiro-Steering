---
inclusion: always
---
<!------------------------------------------------------------------------------------
   Add rules to this file or a short description and have Kiro refine them for you.
   
   Learn about inclusion modes: https://kiro.dev/docs/steering/#inclusion-modes
-------------------------------------------------------------------------------------> 
# Script libraries

This chapter briefly describes some of the script libraries that are used in Optiva.

## Prerequisite

Before you begin working with the script libraries, review these topics:

*   Creating the Script Library on page 41
*   Calling the Script Library on page 43

## DELAYEDWORKFLOW

You can use the DELAYEDWORKFLOW script library to schedule, delay, and repeat workflows on a regular basis. For example, you can schedule Global Calc to run nightly. This library is installed along with the other ION scripts.

**Note:** You must add the supervisor field for the valid Security user, AUTOIMPORT. Any delayed workflow ran by the AUTOIMPORT user must add the supervisor field to get the Security user. Any connection to IDM will be made with that user. The selected user must have IDM access and most of the time, this will be the Optiva service account user. This is only applicable for MT users. See *Obtaining IDM user for interacting IDM in workflow* in the *Infor PLM for Process (Optiva) Application Configuration Guide* for more information.

## Functions

When you use the functions in this script library, an ActionSetStart.xml file is placed in the COR_INBOX_ENTRY table. The COR_INBOX_ENTRY table is in the central OPTIVAIOBOX database. This file has a time entry that causes the XML Import Listeners to retrieve the workflow at a future time.

Three functions are available. Choose the function that is best suited for your business environment. The second parameter defines the time period that is used by the Script Library.

In the first code block, the second parameter is a TimeSpan object.

```vbnet
Public Sub StartDelayedWorkflow(actionSetCode As String, delay As TimeSpan, symbol As String,
    key As String, LabCode As String, ParamArray params As String())
```

In the second code block, the second parameter is an integer that represents a number of days; the third parameter is an integer that represents a number of hours.

```vbnet
Public Sub StartDelayedWorkflow(actionSetCode As String, delayDays As Integer, delayHours As Integer, symbol As String, key As String, LabCode As String, ParamArray params As String())
```

In the third code block, the second parameter is a DateTime Object.

```vbnet
Public Sub StartDelayedWorkflow(actionSetCode As String, startDate As DateTime, symbol As String, key As String, LabCode As String, ParamArray params As String())
```

<table>
  <thead>
    <tr>
      <th>Input Parameters</th>
      <th>Description</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <td>actionSetCode</td>
      <td>Specify the action set that you intend to schedule.</td>
    </tr>
    <tr>
      <td>TimeSpan</td>
      <td>Specify the interval in which the workflow is scheduled to run. For example, (0,1,0,0) represents days, hours, minutes and seconds.</td>
    </tr>
    <tr>
      <td>symbol</td>
      <td>Optional. The workflow is scheduled for the current Optiva symbol.</td>
    </tr>
    <tr>
      <td>key</td>
      <td>Optional. The workflow is scheduled for the current object.</td>
    </tr>
    <tr>
      <td>LabCode</td>
      <td>Optional. Specify the lab that runs the Delayed actionSetCode when you schedule a workflow.</td>
    </tr>
    <tr>
      <td>ParamArray params</td>
      <td>Optional. Specify the input parameters in the sequence in which they are displayed in the **Action Set &gt; Input** tab. You can specify array variables for the input parameters.</td>
    </tr>
    <tr>
      <td>delayDays</td>
      <td>Specify the number of days to delay running the workflow.</td>
    </tr>
    <tr>
      <td>delayHours</td>
      <td>Specify the number of hours to delay running the workflow.</td>
    </tr>
    <tr>
      <td>startDate</td>
      <td>Specify the date and time to begin running the workflow.</td>
    </tr>
  </tbody>
</table>

### Examples for delaying a workflow by hours or days

This workflow is scheduled to run on March 10, 2019. When scheduling a workflow for a specific date, you must schedule it using UTC datetime.

```vbnet
Dim DW As DELAYEDWORKFLOW = new DELAYEDWORKFLOW(Me)
Dim theDate as DateTime = New DateTime(2019, 3, 10)
DW.StartDelayedWorkflow("MYWORKFLOW", theDate, "", "", "", "")
```

This workflow is delayed by 1 day.

```vbnet
Dim DW As DELAYEDWORKFLOW = new DELAYEDWORKFLOW(Me)
Dim delay as New TimeSpan(1, 0, 0, 0)
DW.StartDelayedWorkflow("MYWORKFLOW", delay, "", "", "", "")
```

This workflow is delayed by 1 hour.

```vbnet
Dim DW As DELAYEDWORKFLOW = new DELAYEDWORKFLOW(Me)
Dim delay as New TimeSpan(0, 1, 0, 0)
DW.StartDelayedWorkflow("MYWORKFLOW", delay, "", "", "", "")
```

This workflow is also delayed by 1 hour, using an integer.

```vbnet
Dim DW As DELAYEDWORKFLOW = new DELAYEDWORKFLOW(Me)
DW.StartDelayedWorkflow("MYWORKFLOW", 0, 1, "", "", "", "")
```

## How can you repeat a workflow?

You set the `actionSetCode` parameter to be the same workflow that you are currently running. So, it launches itself after a delay, every time it runs.

This example shows the workflow `FORMULA_UPDATE` being delayed for an hour. This script must be contained with the Action that is run by the workflow `FORMULA_UPDATE` so that it repeats once every hour.

```vbnet
Dim DW As DELAYEDWORKFLOW = new DELAYEDWORKFLOW(Me)
DW.StartDelayedWorkflow("FORMULA_UPDATE", 0, 1, "", "", "", "")
```

This example shows the workflow `MYWORKFLOW` being delayed for an hour. This script must be contained with the Action that is run by the workflow `MYWORKFLOW` so that it repeats once every hour.

```vbnet
Dim DW As DELAYEDWORKFLOW = new DELAYEDWORKFLOW(Me)
Dim delay as New TimeSpan(0,1,0,0)
Dim pArray() as String = {"One", "Two"}
DW.StartDelayedWorkflow("MYWORKFLOW",delay,"", "", "", pArray)
```

The `StartDelayedWorkflow` library is not designed to execute a fixed number of times.

# GDSNXMLHELPER

This script library holds functions which will be assembled to generate GDSN XML. It provides methods for adding a basic GSDN header tag and other common GDSN tags.

# GENERATEGSDNXML

This script library is given as an example on how the functions in the GDSNXMLHELPER script library can be used to generate GDSN XML. The GetXmlForFOODTemplate function uses the methods from the GDSNXMLHELPER script library to generate the XML.

# TPDATAVALIDATIONSEC

Using workflow scripts, you can validate parameter values that are saved in the Optiva database. Then, you can use the TPDATAVALIDATIONSEC script library to send an email to the current user instead of writing an HTML file to disk. The email alerts the user that data for the technical parameters has been saved to the database with an improper format.

You can use this library for unsecured and secured scripting.

### Functions

The TPDATAVALIDATION workflow script is provided in the Optiva seed database and upgraded databases. This script is used by the TPDATAVALIDATIONSEC script library to perform the data validation for the technical parameters.

# Using public functions from the FSLIBIDMHELPER script library

The **FSLIBIDMHELPER** script library has generic and sample scripting functions.

Do not make a change to the **FSLIBIDMHELPER** script library. You can change a function by copying another script library and use the function with that script library instead.

## Pre-requisites

To use the **FSLIBIDMHELPER** script library, you must:

*   Understand IDM functionality.
*   Understand the IDM Xquery syntax. You can use the IDM search in OS to understand the syntax.
*   Access IDM API documentation from the OS installation. Open https://yourOSServerName:9543/ca/index.html to access the IDM API documentation.
*   Understand and read documentation on the **FSLIBIDMHELPER** script library.
*   Understand the IDM .Net library. Open the IDM API documentation to learn about the .NET library. When you import a script to .NET in IDM, the script must include the code Imports DocICP = Infor.DocumentManagement.ICP.

This example shows how to use a public function from the **FSLIBLIDMHELPER** script library:

Imports DocICP = Infor.DocumentManagement.ICP

Dim IDMCHK As FSLIBIDMHELPER = New FSLIBIDMHELPER(Me)
Dim IDMdocCode As String = "PRODUCT_DATA_SHEET_NA"
Dim chk As String = IDMCHK.GetDocCount(IDMdocCode, _OBJECTSYMBOL, _OBJECTKEY)
If chk = 0 Then
    MessageList("IDMCHK = Fail", chk)
Else
    MessageList("IDMCHK = Pass", chk)
End If

If chk > 0 Then
    MessageList("Vendor has completed task - update security access from WRITE to READ on specific Fields and Parameters")
    MessageList("Vendor can add COMMENT to object Doc Code")
    Dim rc1 As Integer = SetSecurityACL("SPECIFICATION", _OBJECTKEY, "@", "USER", "VENDOR", -1)
    Dim rc As Integer = SetSecurityACL("SPECIFICATION", _OBJECTKEY, "@", "USER", "VENDOR", 3)
    Return 111
Else
    MessageList("There are no technical data sheet documents. Please attach and save your work.")
</code>

# Generic functions

These generic functions are in the FSLIBIDMHELPER script library.

## getIDMXQuery

This generic function creates an Xquery for the user. The input parameters are specified in the scripting function.

```vbscript
getIDMXQuery(DocCode As String, objSymbol As String, objKey As String, filename As String) As String
```

<table>
<thead>
<tr>
<th>Input Parameter</th>
<th>Description</th>
</tr>
</thead>
<tbody>
<tr>
<td>DocCode</td>
<td>Document Code</td>
</tr>
<tr>
<td>objSymbol</td>
<td>Optiva Symbol</td>
</tr>
<tr>
<td>objKey</td>
<td>Optiva Object ID</td>
</tr>
<tr>
<td>filename</td>
<td>File Name (optional)</td>
</tr>
</tbody>
</table>

## GetDocCount

Use this function to get the IDM document count from a given Document Code, Object Symbol and Object Key.

```vbscript
GetDocCount(DocCode As String, objSymbol As String, objKey As String) As Long
```

<table>
<thead>
<tr>
<th>Input Parameter</th>
<th>Description</th>
</tr>
</thead>
<tbody>
<tr>
<td>DocCode</td>
<td>The IDM Doc Code to count documents from.</td>
</tr>
<tr>
<td>objSymbol</td>
<td>The Optiva Object Symbol that to count documents from.<br>For example, the symbol for the source object can be a **Formula** or an **Item**.</td>
</tr>
<tr>
<td>objKey</td>
<td>The Optiva Object to count documents from.<br>For example, **ITEM001** or **FORMULA1\0001**.</td>
</tr>
</tbody>
</table>

## GetDocuments

Use this function to get the IDM Documents from a given Doc Code, Object Symbol and Object Key.

```vbnet
GetDocuments(DocCode As String, objSymbol As String, objKey As String, filename As String) As DocICP.CMIItems
```

<table>
<thead>
<tr>
<th>Input Parameter</th>
<th>Description</th>
</tr>
</thead>
<tbody>
<tr>
<td>DocCode</td>
<td>The IDM Doc Code to get documents from.</td>
</tr>
<tr>
<td>objSymbol</td>
<td>The Optiva Object Symbol to get documents from. For example, the symbol for the source object can be a <strong>Formula</strong> or an <strong>Item</strong>.</td>
</tr>
<tr>
<td>objKey</td>
<td>Optiva Object ID to get documents from. For example, ITEM001 or FORMULA1\0001.</td>
</tr>
<tr>
<td>filename</td>
<td>The file name of the IDM documents.</td>
</tr>
</tbody>
</table>

## GetDocuments

Use this function to get all IDM Documents for a given Document Code, Object Symbol and Object Key.

```vbnet
GetDocuments(idmQuery As String) As DocICP.CMIItems
```

<table>
<thead>
<tr>
<th>Input Parameter</th>
<th>Description</th>
</tr>
</thead>
<tbody>
<tr>
<td>idmQuery</td>
<td>An IDM Query to search for IDM Documents in IDM and retrieve the documents.</td>
</tr>
</tbody>
</table>

## GetDocument
Use this function to get the IDM Document from a Doc Code, Object Symbol and Object Key.

```vbnet
GetDocument(DocCode As String, objSymbol As String, objKey As String, filename As String) As DocICP.CMIItem
```

<table>
<thead>
<tr>
<th>Input Parameter</th>
<th>Description</th>
</tr>
</thead>
<tbody>
<tr>
<td>DocCode</td>
<td>The IDM Doc Code to get the document from.</td>
</tr>
<tr>
<td>objSymbol</td>
<td>The Optiva Object Symbol to get the document from.</td>
</tr>
<tr>
<td>objKey</td>
<td>The Optiva Object to get the document for given Doc Code from.<br>For example, ITEM001 or FORMULA1\0001.</td>
</tr>
<tr>
<td>filename</td>
<td>The file name of the IDM Document.</td>
</tr>
</tbody>
</table>

Using public functions from the FSLIBIDMHELPER script library

## DeleteDocument
Use this function to delete the IDM Document for a Document Code associated to an object.

```vba
DeleteDocument(DocCode As String, objSymbol As String, objKey As String, filename As String)
```

<table>
  <thead>
    <tr>
      <th>Input Parameter</th>
      <th>Description</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <td>DocCode</td>
      <td>The IDM Doc Code to delete the document from.</td>
    </tr>
    <tr>
      <td>objSymbol</td>
      <td>The Optiva Object Symbol to delete the document from. For example, the symbol for the source object can be a **Formula** or an **Item**.</td>
    </tr>
    <tr>
      <td>objKey</td>
      <td>The Optiva Object to delete the document from.<br>For example, ITEM001 or FORMULA1\0001.</td>
    </tr>
    <tr>
      <td>filename</td>
      <td>The file name of the IDM Document.</td>
    </tr>
  </tbody>
</table>

## MakeItemCopy
Use this function to copy the specified IDM Document. It will return the new copy of IDM Document.

```vba
MakeItemCopy (cmItemToCopy As DocICP.CMItem) As DocICP.CMItem
```

<table>
  <thead>
    <tr>
      <th>Input Parameter</th>
      <th>Description</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <td>DocItem</td>
      <td>The IDM document (CMItem) to be copied.</td>
    </tr>
  </tbody>
</table>

## MakeItemsCopy
Use this function to copy the specified IDM Documents. It will return the new copy of IDM Documents.

```vba
MakeItemsCopy (cmItemsToCopy As DocICP.CMItems) As DocICP.CMItems
```

<table>
  <thead>
    <tr>
      <th>Input Parameter</th>
      <th>Description</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <td>DocItems</td>
      <td>The IDM document (CMItems) to be copied.</td>
    </tr>
  </tbody>
</table>

## ImportIDMConfig
Use this function to import the API Gateway file from a specific file path. It will configure IDM in the present database.

```vba
ImportIDMConfig (ionAPIFileName As String) As Boolean
```
<table>
  <thead>
    <tr>
      <th>Input Parameter</th>
      <th>Description</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <td>ionAPIFilename</td>
      <td>This is a file path that contains the API Gateway file and the file name.<br><br>This is an example of a file path with the API Gateway file:<br><br>APIFileName = "D:\MyProject\Infor_Home pagesWidgetsSDK\Idm_inhyvwminglexil.ionapi"</td>
    </tr>
  </tbody>
</table>

## UpdateDocument

You can use this function to update IDM documents. To use this function, you must copy this function to the script library.

```javascript
UpdateDocument (DocCode As String, objSymbol As String, objKey As String, filename As String)
```

<table>
  <thead>
    <tr>
      <th>Input Parameter</th>
      <th>Description</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <td>DocCode</td>
      <td>IDM Document Code for the document to be updated.</td>
    </tr>
    <tr>
      <td>objSymbol</td>
      <td>Optiva Symbol. For example, the symbol for the source object can be a **Formula** or an **Item**.</td>
    </tr>
    <tr>
      <td>objKey</td>
      <td>Optiva Object to update the Document for the given Doc Code.<br><br>For example, ITEM001 or FORMULA1\0001.</td>
    </tr>
    <tr>
      <td>filename</td>
      <td>The file name of the IDM Document.</td>
    </tr>
  </tbody>
</table>

## CopyDocument

You can use this function to copy a IDM document from a source object to another target object for the specified document code. To use this function, you must copy this function to the script library.

```javascript
CopyDocument (DocCode As String, sourceSymbol As String, sourceKey As String, targetSymbol As String, targetKey As String, filename As String)
```

<table>
  <thead>
    <tr>
      <th>Input Parameter</th>
      <th>Description</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <td>DocCode</td>
      <td>The IDM Doc Code attached to the source object and to be attached to the target object.</td>
    </tr>
    <tr>
      <td>sourceSymbol</td>
      <td>The Optiva Object Symbol with the attachment to be copied to the target object.<br>For example, the symbol for the source object can be a **Formula** or an **Item**.</td>
    </tr>
    <tr>
      <td>sourceKey</td>
      <td>The Optiva Object to be copied and attached to the target object.<br>For example, ITEM001 or FORMULA1\0001.</td>
    </tr>
    <tr>
      <td>targetSymbol</td>
      <td>Optiva Object Symbol to attach to the copied attachment.<br>For example, the symbol for the source object can be a **Formula** or an **Item**.</td>
    </tr>
    <tr>
      <td>targetKey</td>
      <td>The Optiva Object to attach the copied attachment to.<br>For example, ITEM001 or FORMULA1\0001.</td>
    </tr>
    <tr>
      <td>filename</td>
      <td>The name of the file to copy from source object to target object.</td>
    </tr>
  </tbody>
</table>

## AttachDocument
You can use this function to create and attach an IDM document to the target object.
```
AttachDocument (DocCode As String, objSymbol As String, objKey As String, filename As String)
```

<table>
  <thead>
    <tr>
      <th>Input Parameter</th>
      <th>Description</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <td>DocCode</td>
      <td>The IDM Doc Code for the file to be attached to.</td>
    </tr>
    <tr>
      <td>objSymbol</td>
      <td>The Optiva Object Symbol to attach the document to.<br>For example, the symbol for the source object can be a **Formula** or an **Item**.</td>
    </tr>
    <tr>
      <td>objKey</td>
      <td>The Optiva Object to attach the document to.<br>For example, ITEM001 or FORMULA1\0001.</td>
    </tr>
    <tr>
      <td>FileName</td>
      <td>The file path of the file to attach to the object to.</td>
    </tr>
  </tbody>
</table>

SaveDocumentToFolder
You can use this function to save IDM documents to a specified folder.

Using public functions from the FSLIBIDMHELPER script library

The document object is retrieved from IDM with the GetDocument function. Then the function saves the IDM document object to the folder.

SaveDocumentToFolder (DocItem As String DocICP.CMItem, fileFolder As String)

<table>
  <thead>
    <tr>
      <th>Input Parameter</th>
      <th>Description</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <td>DocItem</td>
      <td>The IDM document to save to the folder.</td>
    </tr>
    <tr>
      <td>fileFolder</td>
      <td>The file folder that document is saved to.</td>
    </tr>
  </tbody>
</table>

RetrieveDocument

Use this function to retrieve the file stream of the specified IDM Document.

RetrieveDocument (DocItem As DocICP.CMItem) As Stream

<table>
  <thead>
    <tr>
      <th>Input Parameter</th>
      <th>Description</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <td>DocItem</td>
      <td>The IDM document from the file stream to be retrieved.</td>
    </tr>
  </tbody>
</table>

# Scripting functions for IDM wizard attachments

GetWizardAttachments, SetWizardAttachmentsToObject, AddAttachmentstoObject, AddAttachmentCopytoObject and AddSpecificAttachmentToObject are the functions used for IDM wizard attachments.
These functions are written in the FSLIBIDMHELPER Script Library.
Use the GetWizardAttachments function to get the attachments based upon the WIPID and Object Symbol. Then, use the SetWizardAttachmentsToObject function to set those attachments to the Object.

## GetWizardAttachments

This function returns the Attachments (CMItems) for a particular Object Symbol. Workflow ID and Object Symbol are inputs to the function.

```vba
Dim idmHelper as new FsLibIDMHelper(me)
Dim attachments As CMItems = idmHelper.GetWizardAttachments( wipID ,ObjectSymbol )
```

<table>
  <thead>
    <tr>
      <th>Input Parameter</th>
      <th>Description</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <td>wipID</td>
      <td>The Workflow ID of wizard where the document are attached to.</td>
    </tr>
  </tbody>
</table>

Using public functions from the FSLIBIDMHELPER script library

<table>
  <thead>
    <tr>
      <th>Input Parameter</th>
      <th>Description</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <td>objSymbol</td>
      <td>The Optiva Object Symbol to retrieve IDM Documents from.<br>For example, the symbol for the source object can be a **Formula** or an **Item**.</td>
    </tr>
  </tbody>
</table>

## SetWizardAttachmentsToObject

This function sets the attachments to the Object. You can pass a single attachment or multiple attachments to the function. You can even modify Attachments (CMItems) and pass the modified attachments to the function.

Workflow ID, Object Symbol, Object Code and Attachments (CMItems) are inputs to the function.

```vba
Dim status As Integer = idmHelper.SetWizardAttachmentsToObject( wipID ,ObjectSymbol, ObjectCode , attachments )
```

<table>
  <thead>
    <tr>
      <th>Input Parameter</th>
      <th>Description</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <td>wipID</td>
      <td>Workflow Id of the wizard that documents are attached to.</td>
    </tr>
    <tr>
      <td>objSymbol</td>
      <td>The Optiva Object Symbol to attach the Documents (CMItems) to.<br>For example, the symbol for the source object can be a **Formula** or an **Item**.</td>
    </tr>
    <tr>
      <td>objKey</td>
      <td>The Optiva Object to attach the Documents (CMItems) to.<br>For example, ITEM001 or FORMULA1\0001.</td>
    </tr>
    <tr>
      <td>cmItems</td>
      <td>The attachments (CMItems) received from the GetWizardAttachments function and to attach to the target symbol.</td>
    </tr>
  </tbody>
</table>

## AddAttachmentstoObject

This function adds all attachments from the Wizard to the business object.

```vba
Public Function AddAttachmentstoObject(wipID as String, ByVal sObjectSymbol as String, ByVal sObjectCode as String)
```

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
      <td>wipID</td>
      <td>String</td>
      <td>Workflows in Progress ID.</td>
    </tr>
    <tr>
      <td>sObjectSymbol</td>
      <td>String</td>
      <td>Optiva object symbol that the IDM attachment will be attached to.</td>
    </tr>
    <tr>
      <td>sObjectCode</td>
      <td>String</td>
      <td>Optiva object code that the IDM attachment will be attached to.</td>
    </tr>
  </tbody>
</table>

This example adds all IDM attachments from the Wizard to a business object.

```vbnet
Public Function AddAttachmentsto0bject(wipID as String, ByVal sObjectSymbol as String, ByVal sObjectCode as String)
    Dim idmHelper as new FsLibIDMHelper(me)
    Dim attachments As CMItems = idmHelper.GetWizardAttachments(wipID ,s0bjectSymbol)
    Dim rc As Integer = idmHelper.SetWizardAttachmentsTo0bject(wipID ,s0bjectSymbol, s0bjectCode , attachments )
    return rc
End Function
```

## AddAttachmentCopytoObject

This function copies IDM attachments and attaches the new attachment to a business object.

```vbnet
Public Function AddAttachmentCopytoObject(wipID as String, ByVal sObjectSymbol as String, ByVal sObjectCode as String)
```

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
      <td>wipID</td>
      <td>String</td>
      <td>Workflows in Progress ID.</td>
    </tr>
    <tr>
      <td>s0bjectSymbol</td>
      <td>String</td>
      <td>Optiva object symbol that the IDM attachment will be attached to.</td>
    </tr>
    <tr>
      <td>s0bjectCode</td>
      <td>String</td>
      <td>Optiva object code that the IDM attachment will be attached to.</td>
    </tr>
  </tbody>
</table>

This example shows the item (CMItem) is copied and MakeItemCopy copies the attachment before the copy of the new attachment is attached to the business object. The method is used to attach the same document to multiple objects. You can copy the document to pass it to the SetWizardAttachmentsTo0bject method.

```vbnet
Public Function AddAttachmentCopytoObject(wipID as String, ByVal s0bjectSymbol as String, ByVal s0bjectCode as String)
    Dim idmHelper as new FsLibIDMHelper(me)
    Dim CopiedItem As CMItem
    Dim attachments As CmItems = idmHelper.GetWizardAttachments(wipID ,s0bjectSymbol)
    CopiedItem = idmHelper.MakeItemCopy(attachments(0))
    Dim rc As Integer = idmHelper.SetWizardAttachmentsTo0bject(wipID ,s0bjectSymbol, s0bjectCode , CopiedItem )
    return rc
End Function
End Class
```

## AddSpecificAttachmenttoObject

This function adds specific IDM attachments from the Wizard to a business object.

```vbnet
Public Function AddSpecificAttachmentto0bject(wipID as String, ByVal s0bjectSymbol as String, ByVal s0bjectCode as String)
```

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
      <td>wipID</td>
      <td>String</td>
      <td>Workflows in Progress ID.</td>
    </tr>
    <tr>
      <td>s0bjectSymbol</td>
      <td>String</td>
      <td>Optiva object symbol that the IDM attachment will be attached to.</td>
    </tr>
    <tr>
      <td>sObjectCode</td>
      <td>String</td>
      <td>Optiva object code that the IDM attachment will be attached to.</td>
    </tr>
  </tbody>
</table>

This example shows how a specific attachment (attachments(0)) is first attached (CMItems) to the business object. This method is used to add only one attachment from the list of attachments.

```vbnet
Public Function AddSpecificAttachmenttoObject(wipID as String, ByVal sObjectSymbol as String, ByVal sObjectCode as String)
    Dim idmHelper as new FsLibIDMHelper(me)
    Dim attachments As CMItems
    attachments = idmHelper.GetWizardAttachments(wipID ,sObjectSymbol)
    Dim rc As Integer = idmHelper.SetWizardAttachmentsToObject(wipID ,sObjectSymbol, sObjectCode , attachments(0))
    return rc
End Function
```
