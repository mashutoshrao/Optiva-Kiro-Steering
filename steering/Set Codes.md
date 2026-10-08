---
inclusion: auto
name: Set Codes
description: Use this file if a ask is anything related to set code
---
<!------------------------------------------------------------------------------------
   Add rules to this file or a short description and have Kiro refine them for you.
   
   Learn about inclusion modes: https://kiro.dev/docs/steering/#inclusion-modes
-------------------------------------------------------------------------------------> 
## AddSetCodes

You can use this function for Optiva Workflows and Copy Methods.

### Purpose

Assigns an object to a set code.

### Syntax

```vbscript
Dim variable As Long = AddSetCodes(Symbol, Object, Code1, [Code2, Code3...])
```

### Arguments

Part | Description
---|---
|Symbol|Type of object. Use empty quotation marks to indicate the current symbol for the workflow or the new object for a copy method.|
Object | Object code to be added to set codes. Use empty quotation marks to indicate the current object for the workflow or the new object for a copy method.
Code1 | First set code to assign the object.
Code2 | Optional. Additional set codes to assign the object.

### Description

AddSetCodes assigns a system object to a set. For example, a formula can be assigned to the FDA set, which indicates formulas with FDA approval. Sets are used to enhance lookups.

A set code can be added manually to an object by opening the object and using the **Set Classification** dialog.

See the *Infor PLM for Process Application Configuration Guide*.

### Examples

In this example, the PIZZASAUCE\003 formula is added to the SAUCE and PIZZA set codes. The symbol and objectKey can be left as empty quotation marks to indicate the formula of the workflow. For formulas and specifications, you must include a backslash (\) followed by a version number if you include the object code.

```vba
Dim lSet As Long = AddSetCodes("FORMULA", "PIZZASAUCE\003", "SAUCE", "PIZZA")
```

In this example, the workflow or new object is added to the SAUCE and TOP set codes.

```vba
Dim lSet As Long = AddSetCodes("", "", "SAUCE", "TOP")
```

In this example, the TOMATOES item is added to the SAUCE_INGRED set code.

```vba
Dim lSet As Long = AddSetCodes("ITEM", "TOMATOES", "SAUCE_INGRED")
```

### Copy Method example

You can construct a copy method where the user selects one or more sets to assign the new object. This example assumes that the fourth and fifth Auto Code segments allow the user to select from a list of sets.

```vba
Dim lSet4 As Long
Dim lSet5 As Long
Dim rc As Long
lSet4 = CopyMethod.Context.GetSegData(4)
lSet5 = CopyMethod.Context.GetSegData(5)
rc = AddSetCodes("", "", lSet4, lSet5)
```

## HasSetCodes

You can use this function for Optiva Workflows, Copy Methods, and Equations.

### Purpose

Checks that an object is assigned to one or more set codes.

### Syntax

```vbnet
Dim variable As Long = HasSetCodes(Symbol, Object, Code1, [Code2,...])
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
<td>Type of object. Use empty quotation marks for the current symbol for the workflow.</td>
</tr>
<tr>
<td>Object</td>
<td>Object for which to check set codes. Use empty quotation marks for the current object for the workflow.</td>
</tr>
<tr>
<td>Code1</td>
<td>The first set code to check.</td>
</tr>
<tr>
<td>Code2</td>
<td>Optional. Additional set codes to check.</td>
</tr>
</tbody>
</table>


### Description

HasSetCodes checks whether the Symbol and Object belong to a list of set codes. The function returns the number of set codes to which the object belongs.

The function returns the number zero (0) to the variable if it does not belong to any of the sets; or it returns a number equal to the number of set code arguments it does belong to.

This function is used to ensure that a correct formula or item is in the workflow, copy method, or equation. For example, a formula for sun screen must be tested for its ability to perform after water contact. Establish a set code for sun screen formulas that have been tested. In a workflow for approving a new sun screen, use HasSetCodes. This function can check that the sun screen has been tested for water contact and cancel the workflow if it has not.

### Examples

In this example, the Symbol and Object can be replaced by empty quotation marks to indicate the current values. If the Object is not a member of the set, then use a conditional statement and the AddSetCodes function to add it to the set.

For formulas and specifications, you must include a backslash (\) followed by a version number.

```vbnet
Dim lSet As Long = HasSetCodes ("FORMULA","PIZZASAUCE\003", "SAUCE","ICE_CREAM")
```

lSet = 1 PIZZASAUCE belongs to the SAUCE set, but not ICE_CREAM.

```vbnet
Dim lSet As Long = HasSetCodes("ITEM","TOMATOES", "NON_PERISHABLES")
```

lSet = 0 TOMATOES does not belong to a NON_PERISHABLE set.

In the next example, the formula for sun screen SPF45\0001.006 has not been tested. It does not belong to the WATERTEST set. The conditional statement ensures that the workflow is cancelled and the user is notified.

```vbnet
Dim lSet As Long = HasSetCodes("FORMULA","SPF45\0001.006", "WATERTEST")
if (lSet = 0) then
    MessageList("SPF45\0001.006 not water tested. Workflow cancelled.")
    Return 9111
else
    Return 111
end if
```