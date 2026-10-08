---
inclusion: auto
name: Copy Method
description: Use this file if asked for copy method
---
<!------------------------------------------------------------------------------------
   Add rules to this file or a short description and have Kiro refine them for you.
   
   Learn about inclusion modes: https://kiro.dev/docs/steering/#inclusion-modes
-------------------------------------------------------------------------------------> 
## Scripting in copy methods
Users can create a unique object or copy an existing object. They can apply a rule code to generate the nameof the new object. A copy method can be used to determine the attributes of the new object.
The rule code (i.e., create rule) determines the auto code segments in the section below it. Users specifysegments to generate a name for the new object according to the rule. They also select a copy method todefine the attributes of the new object.
Optionally, select the Enable Tracing check box to record an entry to the related tracing table when theparticular call to that script has completed. See the Infor PLM for Process Application Configuration Guide formore information about tracing records.
The users selects a rule and copy method and clicks OK. Then, the script that is associated with the copymethod is run

## Script Function
```
Function execute() As Long
End Function
```

Class used for copy method script = CopyScript

### Examples of Copy Methods


## Copy Method to assign sets and classes

This example shows how the Copy Method works during a formula copy.

*   All informational parameters that belong to the FOOD set code are cleared.
*   The item code that is assigned to the formula is removed.
*   The formula is assigned to a class. The class is chosen by the user.
*   The formula is assigned a security level of 15.
*   The formula is assigned to the COLOR set code.
*   The user is notified that the new formula does not have "FORM" in its name.
*   The yield value is converted to an integer and examined as to whether it is a whole number.

```vbnet
Option Strict Off
Imports System
Imports System.Diagnostics

Class CopyScript
    Inherits FcProcFuncSetEventCopyMethod

    Function execute() As Long
        Dim lRemove, lRemove2, lClass1, lSecurity As Long
        Dim lNotify, lHasset, lSet1 As Long
        Dim oExist1 As Object
        Dim lMessage1, lMessage3, lMessage4 As Long
        Dim iInteger1, iName As Integer

        lRemove = ObjPropertyRemove(0, "Value.TP1", "", "", "SET=FOOD", 2)
        lRemove2 = ObjPropertyRemove(0, "ITEMCODE", "", "")
        lClass1 = ObjPropertySet("CLASS D", 1, "CLASS", "", "")
        lSecurity = SetSecurity("", "", "", "", 15)
        lNotify = Notify("JOE_SMITH", "FORMULACOPY", 0, 0, CopyMethod.Context._OBJECTKEY,
            CopyMethod.Context._TASKUSER)
        lHasset = HasSetCodes("", "", "COLOR")
        If lHasset = 0 Then
            lSet1 = AddSetCodes("", "", "COLOR")
        End If
        Dim strObjectKey As String = _OBJECTKEY
        If Not strObjectKey.Contains("FORM") Then
            lMessage1 = MessageList("This is not a FORM Formula")
        End If
        iInteger1 = 0 'set initial value
        If IsNumeric(cStr(iName)) Then
            If CInt(iName) = iName Then
                iInteger1 = 1
            End If
        End If
        If iInteger1 = 1 Then
            lMessage3 = MessageList("The yield is a whole number.")
        Else
            lMessage4 = MessageList("The yield is not a whole number.")
        End If
    End Function
End Class
```

## Copy methods and segment data

Use GetSegData(n) to retrieve data from a create rule segment. GetSegData can also retrieve user input supplied for a create rule.

In this example, the first segment is an Input segment. This segment is where a user inputs the name of a company. The SCRIPT segment for the create rule calls a copy method in which GetSegData retrieves the

company name input from the first segment. Then objProperty can retrieve other detail from that object, such as the description.

```vbnet
Option Strict Off
imports System
imports System.Diagnostics

Class CopyScript
    Inherits FcProcFuncSetEventCopyMethod

    Function execute() As Long
        Dim iBlank As Integer
        Dim lMessage As Long
        Dim Company as String = Context.GetSegData(1)
        Dim Desc As String = ObjProperty("DESCRIPTION", "COMPANY", Company)
        Context.ReturnValue = Desc
    End Function
End Class
```

## Clearing multi-language descriptions when copying an object

The user can clear the descriptions when a user copies an object to a new code, the multi-language descriptions are copied.

The RemoveMultiLanguageValues() script function can be used in copy methods scripts to clear out translated free text that is not relevant to the new object.

The syntax for the new function:

```vbnet
RemoveMultiLanguageValues(objectSymbol As String, objectKey As String) As Integer
RemoveMultiLanguageValues(objectSymbol As String, objectKey As String, fieldNames As List(Of String)) As Integer
RemoveMultiLanguageValues(objectSymbol As String, objectKey As String, fieldNames As List(Of String), languageCodes As List(Of String)) As Integer
```

The objectSymbol and objectKey arguments can be sent a blank to indicate that the current object must be used.

The fieldNames and languageCodes arguments can be sent an empty list to indicate that all fields and all languages must be removed (except for the current language). A null list for either argument results an error.

The fieldNames argument can be sent a mixed list of FIELD_NAME values or DB_FIELD_NAME values from FsValidationField.

This returns the number of multi-language rows removed during this call. Returns -1 if the target object is not found.

If the fieldNames or languageCodes lists contain either fields that don't support multi-language or language codes not defined in Optiva's configuration an error message is displayed.

*   Remove all multi-language rows for all supported fields, in all supported languages:

Dim rc As Long = RemoveMultiLanguageValues("", "")

*   Remove all multi-language rows for the Description and Comment fields, in all supported languages:
    `RemoveMultiLanguageValues("", "", New List(Of String) From {"DESCRIPTION", "COMMENT"})`
*   Remove all multi-language rows for the Description field, for the Greek language:
    `RemoveMultiLanguageValues("", "", New List(Of String) From {"DESCRIPTION"}, New List(Of String) From {"EL-GR"})`
*   Remove all multi-language rows for all supported fields, for the French language:
    `RemoveMultiLanguageValues("", "", New List(Of String), New List(Of String) From {"FR-FR"})`

## GetSegData

### Purpose

Retrieves data from the Sort Order, formerly Line #, column of a Create Rule segment.

### Syntax

```vbnet
Dim Company as String = Context.GetSegData(n)
```

### Argument

<table>
  <thead>
    <tr>
      <th>Part</th>
      <th>Description</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <td>n</td>
      <td>Sort Order #. This may change if you choose to re-order or insert rows.</td>
    </tr>
  </tbody>
</table>

### Example

In this example, the first segment is an input segment. This segment is where a user inputs the name of a company. The script segment for the create rule calls a copy method in which GetSegData retrieves the company name input from the first segment. Then ObjProperty can retrieve other details from that object, such as the description.

## GetSegDataByLineID

### Purpose

Retrieves values by the Row ID, which does not change when reordering or inserting rows.

### Syntax

```vbnet
Dim segVal As Object = CopyMethod.Context.GetSegDataByLineId(n)
```

### Argument

<table>
  <thead>
    <tr>
      <th>Part</th>
      <th>Description</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <td>n</td>
      <td>Row ID</td>
    </tr>
  </tbody>
</table>

### Example

In this example, the copy methods retrieves the segment value with Row ID 3.

```vbnet
Dim segVal As Object = CopyMethod.Context.GetSegDataByLineId(3)
```

