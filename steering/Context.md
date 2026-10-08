---
inclusion: auto
name: Context
description: Use this file if a ask is anything related to Context, addition of mfg. location, selling location to an object.
---
<!------------------------------------------------------------------------------------
   Add rules to this file or a short description and have Kiro refine them for you.
   
   Learn about inclusion modes: https://kiro.dev/docs/steering/#inclusion-modes
-------------------------------------------------------------------------------------> 
## AddContextAttrib

You can use this function for Optiva Workflows and Copy Methods.

### Purpose

Assigns an object to one or more context attributes.

**Syntax**

```vbnet
Dim variable As Long = AddContextAttrib(Symbol, Object, AttribCode, Value1 [, Value2, Value3,...])
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
<td>Symbol</td>
<td>Type of object, such as formula. Use empty quotation marks to indicate the current symbol for the workflow or the new object for the copy method.</td>
</tr>
<tr>
<td>Object</td>
<td>Object code for which to add context attributes. Use empty quotation marks to indicate the current object for the workflow or the new object for the copy method.</td>
</tr>
<tr>
<td>AttribCode</td>
<td>The type of attribute to which context is assigned. Use these codes:</td>
</tr>
<tr>
<td></td>
<td><ul><li>C_BRAND - brand</li><li>C_PRODTYPE - product type</li><li>SELLOC - selling location</li><li>MFGLOC - manufacturing location</li><li>C_ENDUSE - end use</li><li>C_ENDUSER - end user</li></ul></td>
</tr>
<tr>
<td>Value1</td>
<td>The first attribute to which the object is assigned.</td>
</tr>
<tr>
<td>Value2</td>
<td>Optional. Additional attributes to which the object is assigned.</td>
</tr>
</tbody>
</table>

**Description**

Use **AddContextAttrib** to assign context information to an object. Returns the number of attributes to which the object is assigned or zero for an error.

**Examples**

This example assigns selling locations CA, MA, and TN to formula PIZZASAUCE\003. These locations must already exist in the database.

```vbnet
Dim lSell As Long = AddContextAttrib("FORMULA", "PIZZASAUCE\003", "SELLOC","CA", "MA","TN")
```

This example uses **HasContextAttrib** to check if the object of the workflow or new object is assigned to the FOOD end use. If not, it is assigned the context.

```vbnet
Dim lUse As Long = HasContextAttrib("","","C_ENDUSE","FOOD")
    if (lUse = 0) then
        Dim lAdduse As Long = AddContextAttrib("","","C_ENDUSE","FOOD")
        MessageList("This formula is for enduse: Food.")
    end if
```

This example determines if the object of the workflow or new object is assigned to one of the pizza sets. If it is, the object is assigned to the end-use context attribute of the pizza.

```vbnet
Dim lSet As Long = HasSetCodes("", "", "MEAT", "SAUCE", "DOUGH")
if (lSet > 0) then
    AddContextAttrib("", "", "C_ENDUSE", "PIZZA")
end if
```

### Copy Method example

You can construct a copy method to add context attributes to the new object. In this example, the selling location is added using the second Auto Code segment. The brand is added using the third Auto Code segment.

```vbnet
Dim rc As Long
rc = AddContextAttrib("", "", "SELLLOC", CopyMethod.Context.GetSegData(2))
rc = AddContextAttrib("", "", "C_BRAND", CopyMethod.Context.GetSegData(3))

## HasContextAttrib

You can use this function for Optiva Workflows and Copy Methods.

### Purpose
Checks for the presence of specific context attributes for an object and returns the number of attributes for the object.

### Syntax
```vbscript
Dim variable As Long = HasContextAttrib(Symbol, Object, AttribCode, Value1 [, Value2,...])
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
      <td>Type of object, such as formula or item. Use empty quotation marks to indicate the current symbol for the workflow or new symbol for the copy method.</td>
    </tr>
    <tr>
      <td>Object</td>
      <td>Object for which to check context attributes. Use empty quotation marks to indicate the current object for the workflow, or new object for the copy method.</td>
    </tr>
    <tr>
      <td>AttribCode</td>
      <td>The type of attribute to check for the context. Use these codes:
        <ul>
          <li>C_BRAND- brand</li>
          <li>C_PRODTYPE - product type</li>
          <li>SELLOC - selling location</li>
          <li>MFGLOC - manufacturing location</li>
          <li>C_ENDUSE - end use</li>
          <li>C_ENDUSER - end user</li>
        </ul>
      </td>
    </tr>
    <tr>
      <td>Value1</td>
      <td>The first attribute checks for the enumerated lists. These are the "c_" attributes. Use the entries from the **Value** column, not the **Custom Description**.</td>
    </tr>
    <tr>
      <td>Value2, ...</td>
      <td>Optional. Additional attributes to check.</td>
    </tr>
  </tbody>
</table>

### Description
This function checks for the assignment of a context attribute to an object. Returns the number of attributes to which the object is assigned.

### Examples
This example checks that the item SAUCE has selling location ENGLAND. If not, a message is displayed to the user and the workflow is cancelled.
```vbscript
Dim lSell A Long = HasContextAttrib("ITEM", "SAUCE", "SELLOC", "ENGLAND")
if (lSell = 0) then
  MessageList("This item can't be sold in Asia. Workflow cancelled.")
  Return 9111
else
return 111
end if
```

This example checks that the workflow object is assigned to the ACME brand. If not, the object is assigned to that brand.

```vbnet
Dim lBrand As Long = HasContextAttrib("", "", "C_BRAND", "ACME")
if (lBrand = 0) then
    Dim lAcme As Long = AddContextAttrib("", "", "C_BRAND", "ACME")
end if
```

This example checks the product type of the workflow object. If it is a moisturizer, it is added to the CREAMS set code.

```vbnet
Dim lType As Long = HasContextAttrib("", "", "C_PRODTYPE", "MOISTURIZER")
if (lType = 1) then
    Dim lSet As Long = AddSetCodes("", "", "CREAMS")
end if
```

## RemoveContextAttrib

You can use this function for Optiva Workflows and Copy Methods.

### Purpose
Removes an object from one, several or all context attributes.

### Syntax
```vbnet
Dim variable As Long = RemoveContextAttrib(Symbol, Object, AttribCode, Value1 [, Value2,...])
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
      <td>Type of object. Use empty quotation marks to indicate the symbol of the object for the workflow, or for the new object for a copy method.</td>
    </tr>
    <tr>
      <td>Object</td>
      <td>Object for which to remove the context attributes. Use empty quotation marks to indicate the object for the workflow, or the new object for a copy method.</td>
    </tr>
    <tr>
      <td>AttribCode</td>
      <td>The type of attribute whose context is removed. Use these codes:<ul><li>C_BRAND for brand</li><li>C_PRODTYPE for product type</li><li>SELLOC for selling location</li><li>MFGLOC for manufacturing location</li><li>C_ENDUSE for end use</li><li>C_ENDUSER for end user</li></ul></td>
    </tr>
    <tr>
      <td>Value1</td>
      <td>First attribute that is removed from the object. To remove all attributes, use empty quotation marks.</td>
    </tr>
    <tr>
      <td>Value2</td>
      <td>Optional. Additional attributes to which the object should be removed.</td>
    </tr>
  </tbody>
</table>

**Description**

Use this function to delete the context attribute values from an object.

**Examples**

This example removes ACME and GENERIC as brands for the PIZZASAUCE\003 formula.

```vbscript
Dim lRemovebrand As Long = RemoveContextAttrib("FORMULA", "PIZZASAUCE\003", "C_BRAND", "ACME", "GENERIC")
```

This example removes all selling location attributes from the workflow or new object.

```vbscript
Dim lRemovelocation As Long = RemoveContextAttrib("", "", "SELLOC", "")
```

This example checks if the workflow’s object belongs to the FROZEN set code. If so, the FRESH product type attribute is removed from the object.

```vbscript
Dim lSet As Long = HasSetCodes("", "", "FROZEN")
if (lSet = 1) then
  Dim lRemoveproduct As Long
  RemoveContextAttrib("", "", "C_PRODTYPE", "FRESH")
end if
```