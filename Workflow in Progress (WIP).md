---
inclusion: auto
name: Copy Method
description: Use this file if asked for Getting/ Setting information in WIP task(Instructions of task, Description of task), Getting and Setting Inputs of the workflow.
---
<!------------------------------------------------------------------------------------
   Add rules to this file or a short description and have Kiro refine them for you.
   
   Learn about inclusion modes: https://kiro.dev/docs/steering/#inclusion-modes
-------------------------------------------------------------------------------------> 

## WIPInfoGet

You can use this function for Optiva Workflows.

### Purpose

Provides the ability to retrieve any column of information from the FsActionWipSteps table.

### Syntax (without LanguageCode)

```vba
Dim variable As String = WIPInfoGet(Parameter, [LineID, ActionWIPId])
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
      <td>Parameter</td>
      <td>This matches a column name in the FsActionWipSteps table.</td>
    </tr>
    <tr>
      <td>LineID</td>
      <td>Optional. This is the LineID of a particular task in the FsActionWipSteps table.<br>If no value is entered, then the current LineID is used from the context.</td>
    </tr>
    <tr>
      <td>ActionWIPId</td>
      <td>Optional. This is the ActionWIPId that you are retrieving.<br>This is used when you are using the WipInfoGet function from an object other than from an existing ActionWip (e.g., Pending Tasks or Gantt Chart).</td>
    </tr>
  </tbody>
</table>

### Description

WIPInfoGet is used to retrieve the Description column in the FsActionWipSteps table.

```vba
WipInfoGet("DESCRIPTION")
```

### Syntax (with LanguageCode)

```vba
Dim variable As String = WIPInfoGet(Parameter, LanguageCode, [LineID, ActionWIPId])
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
      <td>Parameter</td>
      <td>This matches a column name in the FsActionWipSteps table.</td>
    </tr>
    <tr>
      <td>Language Code</td>
      <td>The language code of the column which is retrieved.</td>
    </tr>
  </tbody>
</table>

| Part | Description |
|---|---|
| LineID | Optional. This is the **LineID** of a particular task in the FsActionWipSteps table. If no value is entered, then the current **LineID** is used from the context. |
| ActionWIPId | Optional. This is the **ActionWIPId** that you are retrieving. This is used when you are using the WipInfoGet function from an object other than from an existing ActionWip (e.g., **Pending Tasks** or **Gantt Chart**). |

### Description

WIPInfoGet is used to retrieve the Language label for the Description column in the FsActionWipSteps table.

```plaintext
WipInfoGet("DESCRIPTION","EN-US")
```

## WIPInfoSet

You can use this function for Optiva Workflows.

### Purpose

Provides the ability to change the value of a column in the FsActionWipSteps table.

### Syntax (without LanguageCode)

```plaintext
Dim variable As Boolean = WIPInfoSet(Parameter, ParameterValue, [LineID,  ActionWIPId])
```

### Arguments

| Part | Description |
|---|---|
| Parameter | This matches a column name in the FsActionWipSteps table. |
| ParameterValue | This is the value which you insert into the FsActionWipSteps table. A failure occurs for the parameters that are read only in the FsValidationField. |
| LineID | Optional. This is the **LineID** of a particular task in the FsActionWipSteps table. If no value is entered, then the current **LineID** is used from the context. |
| ActionWIPId | Optional. This is the **ActionWIPId** that you are retrieving. This is used when you are using the WipInfoSet function from an object other than from an existing ActionWip (e.g., **Pending Tasks** or **Gantt Chart**). |

### Description

WIPInfoSet is used to change the value of the Description column in the FsActionWipSteps table.

```plaintext
WipInfoSet("DESCRIPTION","updated description").
```

### Syntax (with LanguageCode)

```plaintext
Dim variable As Boolean = WIPInfoSet(Parameter, ParameterValue, LanguageCode, [LineID, ActionWIPId])
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
<td>Parameter</td>
<td>This matches a column name in the FsActionWipSteps table.</td>
</tr>
<tr>
<td>ParameterValue</td>
<td>This is the value which you insert into the FsActionWipSteps table. A failure occurs for the parameters that are read only in the FsValidationField.</td>
</tr>
<tr>
<td>Language Code</td>
<td>The language code of the column which is updated.</td>
</tr>
<tr>
<td>LineID</td>
<td>Optional. This is the LineID of a particular task in the FsActionWipSteps table. If no value is entered, then the current LineID is used from the context.</td>
</tr>
<tr>
<td>ActionWIPId</td>
<td>Optional. This is the ActionWIPId that you are retrieving. This is used when you are using the WipInfoSet function from an object other than from an existing ActionWip (e.g., Pending Tasks or Gantt Chart).</td>
</tr>
</tbody>
</table>

### Description

WIPInfoSet is used to change the Language label for the Description column in the FsActionWipSteps table.

```plaintext
WipInfoSet("DESCRIPTION","updated description","En-US").
```

### Syntax

```plaintext
Dim variable As Boolean = WIPInfoSet(LineID, ActionWIPId, MultiLangData)
```
You can use the syntax for updating the language labels for an array of fields.

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
      <td>LineID</td>
      <td>Optional. This is the LineID of a particular task in the FsActionWipSteps table. If no value is entered, then the current LineID is used from the context.</td>
    </tr>
    <tr>
      <td>ActionWIPId</td>
      <td>Optional. This is the ActionWIPId that you are retrieving. This is used when you are using the WipInfoSet function from an object other than from an existing ActionWip (e.g., Pending Tasks or Gantt Chart).</td>
    </tr>
    <tr>
      <td>MultiLangdata</td>
      <td>This is an array of multi-language fields in the FsActionWipSteps table.<br>MultiLangData contains these fields:<br>LanguageCode.<br>MultiLangField.<br>Description.</td>
    </tr>
  </tbody>
</table>

### Description

WIPInfoSet is used to change the Language labels for the Description column in the FsActionWipSteps table.

```vbscript
dim fieldsandValues(1) As Optiva.ScriptProxy.MultiLangData
fieldsandValues(0) = new Optiva.ScriptProxy.MultiLangData()
fieldsandValues(0).languageCode = "EN-US"
fieldsandValues(0).multiLangField = "DESCRIPTION"
fieldsandValues(0).description = "EN-US description updated using WIPINFOSET"

fieldsandValues(1)= new Optiva.ScriptProxy.MultiLangData()
fieldsandValues(1).languageCode = "DE-DE"
fieldsandValues(1).multiLangField = "DESCRIPTION"
fieldsandValues(1).description = "DE-DE description updated using WIPINFOSET"
WIPInfoSet (0,0,fieldsandValues)
```

### Example of WIPInfoGet and WIPInfoSet

This example shows how WIPInfoGet and WIPInfoSet work together. In this scenario, a description is retrieved from the Wizard and added to the **Instructions** field in the **Pending Tasks** grid.

```vbscript
'Use WipParamGet to retrieve the Instruction input from the Workflow or Wizard
Dim sInstr As String = WipParamGet("WIPINSTRUCTION")
'Use WipInfoSet to write the Instruction text to the Instructions column in FSACTONWIPSTEPS
Dim sSetInstr As Boolean = WIPInfoSet("INSTRUCTIONS", sInstr)
```

## WipParamGet

You can use this function for Optiva Workflows.

### Purpose

Retrieves a parameter that was set by a previous action or by user input.

### Syntax

```vbnet
Dim variable As String = WipParamGet("parameter")
```

### Description

WipParamGet depends on a parameter value set by a previous action with WipParamSet, or by user input. For more information about user input parameters, see the descriptions for the **Action Set** form in the *Infor PLM for Process Workflow Administration Guide*.

---

### Examples

In this example, WipParamSet assigns the value FSI to the Input Parameter, NOTIFY.

```vbnet
Dim sValue As String = ("FSI")
MessageList(sValue)
Dim sReturncode As String = WIPParamSet("NOTIFY", sValue)
```

WipParamGet retrieves the value from the Input Parameter, NOTIFY, and stores it in the variable value.

```vbnet
Dim sValue As String = ("FSI")
sValue = WIPParamGet("NOTIFY")
MessageList(sValue)
Return 111
```

This sample action starts an action set where a report name is passed in by the user as an input parameter.

```vbnet
Dim sReportName As String = WipParamGet("REPORTNAME")
```

This action is for an MSDS report. It first retrieves an itemAlias input parameter for the action set. If this parameter is not defined for the action set, then the value is blank. The explode and combine are performed on the item codes.

```vbnet
Dim sItemAlias As String = WipParamGet("ITEMALIAS")
```

## WipParamSet

You can use this function for Optiva Workflows.

### Purpose

Sets a parameter value.

### Syntax

```vbscript
Dim variable As String = WipParamSet("parameter", value)
```

### Description

WipParamSet sets a value in an action step. The value can be retrieved in a later step by WipParamGet.

---

### Examples

This example assigns the value FSI to the Input Parameter, NOTIFY.

```vbscript
Dim sValue as String = MessageList(sValue)
Dim sReturncode As String = WIPParamSet("NOTIFY", sValue)
```

WipParamGet retrieves the value from the Input Parameter, NOTIFY, and stores it in the variable named `sValue`.

```vbscript
Dim sValue As String = WIPParamGet("NOTIFY")
MessageList(sValue)
Return 111
```
