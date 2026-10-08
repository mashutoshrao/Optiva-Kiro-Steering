---
inclusion: always
---
<!------------------------------------------------------------------------------------
   Add rules to this file or a short description and have Kiro refine them for you.
   
   Learn about inclusion modes: https://kiro.dev/docs/steering/#inclusion-modes
-------------------------------------------------------------------------------------> 
## GenerateCodeNumber

You can use this function for Optiva Workflows and Copy Methods.

### Purpose

Generates a unique code number that can be used to create new objects.

### Syntax

```vbnet
Dim codeNumber As String = GenerateCodeNumber(createRule)
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
      <td>createRule</td>
      <td>The create rule to be used for the object code.<br>To create a unique key, include a SEQUENCE segment in the rule instead of NEXTKEY.</td>
    </tr>
  </tbody>
</table>

### Description

GenerateCodeNumber uses a pre-defined create rule to generate a code number for an object.

### Example

This example uses the EXP FORMULA create rule to generate a code for a new formula based on PIZZA SAUCE\0001. The code is then applied to a new formula that is saved from an existing formula PIZZA SAUCE\0001.

```vba
Dim sCodeNumber As String = GenerateCodeNumber("EXP FORMULA")
MessageList("The new formula will be:", sCodeNumber)
Dim rc As Long = ObjectSaveAs("FORMULA", "PIZZA SAUCE\0001", sCodeNumber)
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

## ObjectSaveAs

You can use this function for Optiva Workflows and Copy Methods.

### Purpose

Copies an object to a new key code that does not exist in the database.

### Syntax

```vbnet
Dim variable As Long = ObjectSaveAs(Symbol, Source0bject, Target0bject)
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
<td>Type of object. Use empty quotation marks for the current symbol for the workflow, or the new object.</td>
</tr>
<tr>
<td>Source0bject</td>
<td>Object to be copied. Use empty quotation marks for the current object for the workflow or the new object for the copy method.</td>
</tr>
<tr>
<td>Target0bject</td>
<td>Key code of the new object that is copied from Source0bject. This object must not exist before you run the workflow or copy method.</td>
</tr>
</tbody>
</table>

### Description

ObjectSaveAs copies an existing object keycode to a new object keycode. The new object keycode does not exist in the database before you run the workflow or copy method.

### Example

In this example, formula PIZZASAUCE\003.002 and item OLIVE_OIL_BULK do not exist in the Optiva database before you run the workflow or copy method.

```
Dim lSave As Long = ObjectSaveAs("FORMULA","SAUCE\003","SAUCE\003.002")
Dim lSaveItem As Long = ObjectSaveAs("ITEM","OLIVE_OIL","OLIVE_OIL_ BULK")
```

## GetSequence

You can use this function for Optiva Workflows.

### Purpose

Generates the next sequence number.

### Syntax

```vbscript
Dim getSequence as Long = GetSequence("SetType", integer)
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
      <td>SetType as string</td>
      <td>A grouping identifier that enables you to maintain multiple sequences within the system.</td>
    </tr>
    <tr>
      <td>Increment as Integer</td>
      <td>Controls what the sequence is incremented by. Typically this value is 1, but it can be other positive integer values.</td>
    </tr>
  </tbody>
</table>

### Description

This function is used to generate the next sequence number. The SetType does not have to be created in the system. It is an arbitrary string.

Suppose two departments want to use sequence numbers. You can create unique numbers for each department. For example, you can use the name of the department.

### Example

This syntax generates the next number in the "aSET" sequence, incrementing it by 1.

```vbscript
Dim getSequence as Long = GetSequence("aSET", 1)
```

You can keep an arbitrary number of sequences by this method. For example, "Sequence1", "Sequence2". You can increment it by any amount (1 or greater; use only positive integers).
