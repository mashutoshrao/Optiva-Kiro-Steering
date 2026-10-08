---
inclusion: auto
name: References
description: Use this file if a ask is anything related to References
---
<!------------------------------------------------------------------------------------
   Add rules to this file or a short description and have Kiro refine them for you.
   
   Learn about inclusion modes: https://kiro.dev/docs/steering/#inclusion-modes
-------------------------------------------------------------------------------------> 
## AddReferences

You can use this function for Optiva Workflows and Copy Methods.

### Purpose

References one object to another.

### Syntax

```vba
Dim variable As Integer = AddReferences(Symbol, Object, ReferenceSymbol, Reference0bject1 [, Reference0bject2, Reference0bject3...])
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
      <td>Type of object. Use empty quotation marks to indicate the current symbol for the workflow or copy method.</td>
    </tr>
    <tr>
      <td>Object</td>
      <td>Object code for which references are added. Use empty quotation marks to indicate the current object for the workflow or copy method.</td>
    </tr>
    <tr>
      <td>ReferenceSymbol</td>
      <td>Type of object that is referenced.</td>
    </tr>
    <tr>
      <td>Reference0bject1</td>
      <td>First object code to be referenced (shown on the **References** tab) of the Object. You can use the object ID in place of the object code.</td>
    </tr>
    <tr>
      <td>Reference0bject2, Reference Object3</td>
      <td>Optional. Additional object codes to be referenced (shown on the **References** tab) of the Object. You can use object IDs in place of the object code.</td>
    </tr>
  </tbody>
</table>

### Description

AddReferences adds objects, as references, to another object. References are listed on the **References** tab of objects. The system must be configured to allow the reference before using this function. For example, to use AddReferences to reference specifications to formulas, formulas must already be configured to reference specifications.

See the *Infor PLM for Process Application Configuration Guide*.

### Examples

In this example, the PIZZASAUCE\199 and PIZZASAUCE\202 specifications are referenced to the PIZZASAUCE\002 formula.

```vba
Dim iAddref As Integer = AddReferences("FORMULA", "PIZZASAUCE\002", "SPECIFICATION", "PIZZASAUCE\199", "PIZZASAUCE\202")
```

If the formula is the object of the workflow or copy method, then specify empty quotation marks for the symbol and object.

```vba
Dim iAddref As Integer = AddReferences("", "", "SPECIFICATION", "PIZZASAUCE\199", "PIZZASAUCE\202")
```

In this example, the first statement retrieves the manufacturing item code of the formula. The next line retrieves the object code of the specification. The specification is referenced on the **References** tab of the item.


The specification is added to the **References** tab of the formula.

```vbscript
Dim oItem, oSpec As Object
Dim iAddref As Integer
oItem = ObjProperty("ITEMCODE")
oSpec = ObjProperty("OBJECTCODE.REF.V\REFITEM3","ITEM",oItem,
"SPECIFICATION",2)
iAddref = AddReferences("", "", "SPECIFICATION", oSpec)
```

### Copy Method example

This copy method adds the customer name as a reference to the new object.

```vbscript
Dim rc As Long
rc = AddReferences("", "", "CUSTOMER", CopyMethod.Context.GetSegData(2),
"", "")
```

### Duplicate References example

The **LINE_ID** column, in the FSOBJECTREFERENCE table, supports duplicate references. In this example, SP-0001\0001 is added to the SPECIFICATION object twice.

```vbscript
Dim iAddref As Integer = AddReferences("", "",
"SPECIFICATION", "SP-0001\0001", "SP-0002\0003", "SP-0001\0001")
```

## RemoveReferences

You can use this function for Optiva Workflows and Copy Methods.

### Purpose

Removes references from an object.

### Syntax

```vba
Dim variable as Long = RemoveReferences(Symbol, ObjectKey[, ObjectType, ObjectCode])
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
      <td>Type of object. Use empty quotation marks to indicate the symbol of the object for the workflow, or the new object for the copy method.</td>
    </tr>
    <tr>
      <td>Object Key</td>
      <td>Object whose references are to be removed. Use empty quotation marks to indicate the object for the workflow, or the new object for the copy method.</td>
    </tr>
    <tr>
      <td>Object Type</td>
      <td>Optional. Remove all references of a single object type, for example VENDO R.</td>
    </tr>
    <tr>
      <td>Object Code</td>
      <td>Optional. Remove only a single reference that matches the type and code.</td>
    </tr>
  </tbody>
</table>

### Description

RemoveReferences removes some or all references from the **References** tab of an object.

### Examples

In this example, all of the references from the **References** tab of the PIZZA_SAUCE\003 formula are removed.

```vba
Dim lRemove As Long
RemoveReferences("FORMULA", "PIZZA_SAUCE\003")
```

In the next example, all of the references are removed from the **References** tab of the workflow object or newly created object.

```vba
Dim lRemove As Long = RemoveReferences("", "")
```

In this example, all of the VENDOR references are removed.

```vba
Dim lRemove As Long = RemoveReferences("", "", "VENDOR")
```

Only the specification reference SP-0001\0001 is removed.

```vba
Dim lRemove As Long = RemoveReferences("", "", "SPECIFICATION", "SP-0001\0001")```

## RemoveSetCodes

You can use this function for Optiva Workflows and Copy Methods.

### Purpose

Removes an object from one, several, or all set codes.

### Syntax

```vba
Dim variable As Long = RemoveSetCodes(Symbol, Object, Code1[, Code2,...])
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
      <td>Type of object. Use empty quotation marks to indicate the symbol of the object for the workflow, or the new object for a copy method.</td>
    </tr>
    <tr>
      <td>Object</td>
      <td>Object to be removed from the set codes. Use empty quotation marks to indicate the object for the workflow, or the new object for a copy method.</td>
    </tr>
    <tr>
      <td>Code1</td>
      <td>First Set code from which you want to remove the object. To remove the object from all set codes, use empty quotation marks.</td>
    </tr>
    <tr>
      <td>Code2</td>
      <td>Optional. Additional set codes from which to remove the object.</td>
    </tr>
  </tbody>
</table>

### Description

RemoveSetCodes removes a system object from a set. For example, a formula can be assigned to the FDA set, which indicates formulas with FDA approval. If a formula has to be retested under new guidelines, then run a workflow to remove the formula from the set.

A newly created object is automatically added to the set code for the lab. You cannot remove this set code.

### Examples

This example removes the PIZZASAUCE\003 formula from the SUB set.

```vba
Dim lRemoveset As Long
RemoveSetCodes("FORMULA", "PIZZASAUCE\003", "SUB")
```

The next example removes the workflow object or newly-created object from the TOP_COAT and SEALANT sets.

```vba
Dim lRemoveset As Long
RemoveSetCodes("", "", "TOP_COAT", "SEALANT")
```

In this example, the script checks if the workflow object or newly-created object belongs to a program. If not, the script removes the object from all sets.

```vbscript
Dim oProgram As Object = ObjProperty("PROJECTCODE")
Dim iBlank As Integer = IsBlank(oProgram)
if (iBlank = 1) then
    Dim lRemoveset As Long = RemoveSetCodes("", "", "")
end if
```