---
inclusion: auto
name: Validations
description: Different types of functions to perform validations. Use this file only to perform validations
---

<!------------------------------------------------------------------------------------
   Add rules to this file or a short description and have Kiro refine them for you.
   
   Learn about inclusion modes: https://kiro.dev/docs/steering/#inclusion-modes
-------------------------------------------------------------------------------------> 
## CheckIn

You can use this function for Optiva Workflows.

**Purpose**

Unlocks an object.

### Syntax

```vbscript
Dim variable As Integer = CheckIn(symbol, keycode)
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
      <td>Object symbol. Leave blank for the current object of the workflow.</td>
    </tr>
    <tr>
      <td>KeyCode</td>
      <td>Object to unlock. Leave blank for the current object of the workflow.</td>
    </tr>
  </tbody>
</table>

### Description

Check-in unlocks an object. It returns a value of **0** or **1**.

<table>
  <thead>
    <tr>
      <th>Value</th>
      <th>Description</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <td>0</td>
      <td>Did not check in the object. May already be checked in.</td>
    </tr>
    <tr>
      <td>1</td>
      <td>Checked in the object.</td>
    </tr>
  </tbody>
</table>

### Examples

This example checks in the current object from which the Workflow was launched

```vbscript
Dim rc As Integer = CheckIn("", "")
If rc = 1 Then
    MessageList(_OBJECTSYMBOL & " " & _OBJECTKEY & " checked in.")
Else
    MessageList(_OBJECTSYMBOL & " " & _OBJECTKEY & " not checked in.")
End if
```

This example checks in a specific object

```vbscript
Dim rc As Integer = CheckIn("FORMULA", "FS-0010\0001")
If rc = 1 Then
    MessageList("Formula FS-0010\0001 has been checked in.")
Else
    MessageList("Formula FS-0010\0001 could not be checked in.")
End if
```

## CheckLoopCount

You can use this function for Optiva Workflows.

### Purpose
Allows looping, or a repeat of action steps, during a workflow.

### Syntax
```vbnet
Dim variable As Integer = CheckLoopCount()
```

### Arguments
None.
Typically, at least one argument for a function statement is required by the system. The argument for this function is not used.

For this function, use a conditional statement to indicate the number of times the workflow can loop.

Indicate on the **Action Set** form the action to which the workflow loops. In an Action Set, you can choose any previous line as the **Loop Line**, regardless of how complex the Action Set is. There are no restrictions on branched or nested steps.

Action sets can contain multiple loops. You can loop backward, but never forward.

If a task within a looping group has been skipped, then the task will remain in a skipped state.

**Note:** Include only one loop per action.

### Description
Use CheckLoopCount and its return code, **811**, to repeat actions in a workflow for a specified number of times. Or, use it until a satisfactory response is reached. This function is most appropriate for the Return event.

### Examples
In this example, the lab manager returns the workflow to the technical director up to three times. This script is assigned to the RETURN event. Each time the lab manager clicks **Return**, the workflow returns to the action that is assigned to the technical director. After three returns, if the lab manager clicks **Return**, the workflow cancels.

```vbnet
Dim iLoopcount As Integer = CheckLoopCount()
if (iLoopcount <= 3) then
    MessageList("Returning to the technical director.")
    Return 811
else
    MessageList("Workflow cancelled.")
    Return 9111
end if
```

Open the **Action Set** form. In the line that contains the CheckLoopCount function, add the **Line ID** to the destination line in the **Loop Line** column.

In this example, the CheckLoopCount function is in the second action. The value of **1** instructs the workflow to loop to the action in **Line ID 1**.

**Steps**

<table>
  <thead>
    <tr>
      <th>Row ID</th>
      <th>Description</th>
      <th>Level</th>
      <th>Action Code</th>
      <th>Loop Line</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <td>1</td>
      <td></td>
      <td>0</td>
      <td>CHECK ITEM</td>
      <td>0</td>
    </tr>
    <tr>
      <td>2</td>
      <td></td>
      <td>0</td>
      <td>REGULATORY REVIEW</td>
      <td>1</td>
    </tr>
  </tbody>
</table>

## CheckOut

You can use this function for Optiva Workflows.

### Purpose

Locks an object.

### Syntax

```vba
Dim variable As Integer = CheckOut(symbol, keycode)
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
      <td>Object symbol. Leave blank for the current object of the workflow.</td>
    </tr>
    <tr>
      <td>KeyCode</td>
      <td>Object to lock. Leave blank for the current object of the workflow.</td>
    </tr>
  </tbody>
</table>

### Description

Checkout locks an object. It returns a value of **0** or **1**.

<table>
  <thead>
    <tr>
      <th>Value</th>
      <th>Description</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <td>0</td>
      <td>Did not check out the object. May already be checked out.</td>
    </tr>
    <tr>
      <td>1</td>
      <td>Checked out the object.</td>
    </tr>
  </tbody>
</table>

If the object is already checked out and the workflow returns **0**, you can use IsLockedBy to verify by whom.

### Examples

This example checks out the current object from which the Workflow was launched

```vbscript
Dim rc As Integer = CheckOut("", "")
If rc = 1 Then
    MessageList(_OBJECTSYMBOL & " " & _OBJECTKEY & " checked out.")
Else
    MessageList(_OBJECTSYMBOL & " " & _OBJECTKEY & " not checked out.")
End if
```

This example checks out a specific object

```vbscript
Dim rc As Integer = CheckOut("FORMULA", "FS-0010\0001")
If rc = 1 Then
    MessageList("Formula FS-0010\0001 has been checked out.")
Else
    MessageList("Formula FS-0010\0001 could not be checked out.")
End if
```

## DocExist

You can use this function for Optiva Workflows and Copy Methods.

### Purpose

Checks for the existence of text and attachments to a function code.

---

### Syntax

```vba
Dim variable As Integer = DocExist(Symbol, Object, FunctionCode)
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
      <td>Type of object. Use empty quotation marks to indicate the current symbol for the workflow.</td>
    </tr>
    <tr>
      <td>Object</td>
      <td>Object code to check for text and attachments. Use empty quotation marks to indicate the current object for the workflow.</td>
    </tr>
    <tr>
      <td>FunctionCode</td>
      <td>A valid function code for the indicated symbol.</td>
    </tr>
  </tbody>
</table>

### Description

DocExist checks for the presence of text and attachments to a particular function code for the indicated object. It returns a number:

<table>
  <thead>
    <tr>
      <th>Return</th>
      <th>Text</th>
      <th>Attachments</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <td>-1</td>
      <td>Invalid function code</td>
      <td>Invalid function code</td>
    </tr>
    <tr>
      <td>0</td>
      <td>No</td>
      <td>No</td>
    </tr>
    <tr>
      <td>1</td>
      <td>Yes</td>
      <td>No</td>
    </tr>
    <tr>
      <td>2</td>
      <td>No</td>
      <td>Yes</td>
    </tr>
    <tr>
      <td>3</td>
      <td>Yes</td>
      <td>Yes</td>
    </tr>
  </tbody>
</table>

### Examples

This example checks for the presence of text and attachments to a workflow formula. The formula code is PIZZA_SAUCE\003. If none are detected, a message is displayed to the user.

```vba
Dim iTex As Integer = DocExist("FORMULA", "PIZZASAUCE\003", "REPORT")
if (iText < 1) then
    MessageList("You must attach a report.")
end if
```

This example checks for the presence of text in the SAFETY function code for the workflow object or new object. If there is none, a message is displayed to the user.

```vba
Dim iTex As Integer = DocExist("", "", "SAFETY")
if (iText = 1 OR iText = 3) then
    Return 111
else
    MessageList("There are no Safety documents. Include some.")
    Return 1
end if
```

```vbnet
Dim iAttach As Integer = DocExist("", "", "WEB")
if (iAttach < 2) then
    MessageList("You must attach the intranet site.")
end if
```
## DocExist

You can use this function for Optiva Workflows and Copy Methods.

### Purpose

Checks for the existence of text and attachments to a function code.

---

### Syntax

```vba
Dim variable As Integer = DocExist(Symbol, Object, FunctionCode)
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
      <td>Type of object. Use empty quotation marks to indicate the current symbol for the workflow.</td>
    </tr>
    <tr>
      <td>Object</td>
      <td>Object code to check for text and attachments. Use empty quotation marks to indicate the current object for the workflow.</td>
    </tr>
    <tr>
      <td>FunctionCode</td>
      <td>A valid function code for the indicated symbol.</td>
    </tr>
  </tbody>
</table>

### Description

DocExist checks for the presence of text and attachments to a particular function code for the indicated object. It returns a number:

<table>
  <thead>
    <tr>
      <th>Return</th>
      <th>Text</th>
      <th>Attachments</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <td>-1</td>
      <td>Invalid function code</td>
      <td>Invalid function code</td>
    </tr>
    <tr>
      <td>0</td>
      <td>No</td>
      <td>No</td>
    </tr>
    <tr>
      <td>1</td>
      <td>Yes</td>
      <td>No</td>
    </tr>
    <tr>
      <td>2</td>
      <td>No</td>
      <td>Yes</td>
    </tr>
    <tr>
      <td>3</td>
      <td>Yes</td>
      <td>Yes</td>
    </tr>
  </tbody>
</table>

### Examples

This example checks for the presence of text and attachments to a workflow formula. The formula code is PIZZA_SAUCE\003. If none are detected, a message is displayed to the user.

```vba
Dim iTex As Integer = DocExist("FORMULA", "PIZZASAUCE\003", "REPORT")
if (iText < 1) then
    MessageList("You must attach a report.")
end if
```

This example checks for the presence of text in the SAFETY function code for the workflow object or new object. If there is none, a message is displayed to the user.

```vba
Dim iTex As Integer = DocExist("", "", "SAFETY")
if (iText = 1 OR iText = 3) then
    Return 111
else
    MessageList("There are no Safety documents. Include some.")
    Return 1
end if
```

```vbnet
Dim iAttach As Integer = DocExist("", "", "WEB")
if (iAttach < 2) then
    MessageList("You must attach the intranet site.")
end if
```

## IsBlank

You can use this function for any Optiva script type: workflow, equations, copy methods, script hooks, etc.

**Purpose**

Determines if a value from a field is blank or contains characters. Null is tested as well.

**Syntax**

```vbnet
Dim variable As Integer = IsBlank(Value)
```

**Arguments**

<table>
<thead>
<tr>
<th>Part</th>
<th>Description</th>
</tr>
</thead>
<tbody>
<tr>
<td>Value</td>
<td>The field that is checked.</td>
</tr>
</tbody>
</table>

---

IsBlank determines if a variable has a non-null, non-blank value. For example, you can use this function to determine if a formula has an item code associated with it.

The function returns a value of **0** or **1**.

<table>
  <thead>
    <tr>
      <th>Value</th>
      <th>Description</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <td>0</td>
      <td>Field is not blank.</td>
    </tr>
    <tr>
      <td>1</td>
      <td>Field is blank.</td>
    </tr>
  </tbody>
</table>

### Examples

In this example, the first line places the ITEMCODE string of the PIZZASAUCE formula into the local variable oItemcode.

The second line checks whether there is a string in oItemcode. The script determines if the formula has a MFG item code.

```vbscript
Dim oItemcode As Object = ObjProperty("ITEMCODE", "FORMULA", "PIZZASAUCE\003")
Dim iBlank As Integer = IsBlank(oItemcode)
if (iBlank = 1) then
    MessageList("This formula does not have a MFG item code.")
end if
```

The next example checks if a formula has a work code.

```vbscript
Dim oWorkcode As Object = ObjProperty("WORKCODE", "FORMULA", "PIZZA_SAUCE\003")
Dim iBlank As Integer = IsBlank(oWorkcode)
if (iBlank = 1) then
    MessageList("This formula does not have a work code.")
end if
```

In the previous examples, you can leave the last two arguments, representing Symbol and Object, respectively, as empty quotation marks. Empty quotation marks indicate the new object for a copy method.

## IsLocked

You can use this function for Optiva Workflows.

### Purpose

Determines if an object is locked (i.e., checked out).

**Syntax**

```vbnet
Dim variable As Integer = IsLocked(Symbol, ObjectKey)
```

**Description**

IsLocked checks the locked (i.e., checked out) status of a formula.

<table>
  <tr>
    <th>Part</th>
    <th>Description</th>
  </tr>
  <tr>
    <td>Symbol</td>
    <td>Object symbol. Leave out for the current object of the workflow.</td>
  </tr>
  <tr>
    <td>KeyCode</td>
    <td>Object code. Leave out for the current object of the workflow.</td>
  </tr>
</table>

It returns a value of **0** or **1**.

<table>
  <tr>
    <th>Value</th>
    <th>Description</th>
  </tr>
  <tr>
    <td><strong>0</strong></td>
    <td>Returns **0** if the object is not locked.</td>
  </tr>
  <tr>
    <td><strong>1</strong></td>
    <td>Returns **1** if the object is locked.</td>
  </tr>
</table>

Use CheckIn or CheckOut depending on the return value.

**Note:** You can override security and change a locked formula using Workflow. Use discretion when overriding security.

**Examples**

This example determines if the current object from which the Workflow was launched has been locked (i.e., checked out).

```vbnet
Dim rc As Integer = IsLocked()
If rc = 1 Then
    MessageList(_OBJECTSYMBOL & " " & _OBJECTKEY & " is locked.")
Else
    MessageList(_OBJECTSYMBOL & " " & _OBJECTKEY & " not locked.")
End if
```

This example determines if a specific object has been locked (i.e., checked out).

```vbnet
Dim rc As Integer = IsLocked("FORMULA", "FS-0010\0001")
If rc = 1 Then
    MessageList( "Formula FS-0010\0001 is currently checked out.")
Else
    MessageList("Formula FS-0010\0001 is not checked out.")
End if
```

## IsLockedBy

You can use this function for Optiva Workflows.

### Purpose

Checks if an object is locked and returns the user who locked it.

### Syntax

```vbscript
Dim variable As String = IsLockedBy(Symbol, ObjectKey)
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
      <td>Object symbol. Leave out for the current object of the workflow.</td>
    </tr>
    <tr>
      <td>KeyCode</td>
      <td>Object code. Leave out for the current object of the workflow.</td>
    </tr>
  </tbody>
</table>

It returns a value of `0` or user ID.

<table>
  <thead>
    <tr>
      <th>Value</th>
      <th>Description</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <td>0</td>
      <td>Returns `0` if the object is not locked.</td>
    </tr>
    <tr>
      <td>user ID</td>
      <td>Returns the ID of the user who checked it out if it is locked.</td>
    </tr>
  </tbody>
</table>

### Examples

This example determines the user who has checked out the current object from which the Workflow was launched.

```vbscript
Dim rc As String = IsLockedBy()
If rc Is Nothing Then
    MessageList(_OBJECTSYMBOL & " " & _OBJECTKEY & " not checked out.")
Else
    MessageList(_OBJECTSYMBOL & " " & _OBJECTKEY & " checked out by" & rc)
End if
```

This example determines if a specific object has been checked out and if so by whom.

```vbscript
Dim rc As String = IsLockedBy("FORMULA", "FS-0010\0001")
If rc Is Nothing Then
    MessageList("Formula FS-0010\0001 is not checked out.")
Else
    MessageList("Formula FS-0010\0001 is checked out by" & rc)
End if
```
## ObjectExists

You can use this function for Optiva Workflows and Copy Methods.

### Purpose

Checks for the existence of an object key for a symbol.

### Syntax

```vba
Dim variable As Integer = ObjectExists(Object, Symbol)
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
<td>Object</td>
<td>Object to check for existence. Use empty quotation marks for the current object, for the workflow, or the new object for the copy method.</td>
</tr>
<tr>
<td>Symbol</td>
<td>Type of object. Use empty quotation marks for the current symbol for the workflow, or the new object for the copy method.</td>
</tr>
</tbody>
</table>

### Description

ObjectExists determines if an object key exists for a symbol. You can use ObjectExists to check if an object key, such as a particular formula or item, already exists.

*   If the object does not exist, then a value of **0** is returned.
*   If an object does exist, then a value of **1** is returned.

### Examples

In these examples, ObjectExists determines if the formula or item with the specified key exists. For formulas and specifications, you must include a backslash (\) followed by a version number.

```vbscript
Dim iExists As Integer = ObjectExists("PIZZASAUCE\003", "FORMULA")
iExists = ObjectExists("OLIVE_OIL", "ITEM")
```