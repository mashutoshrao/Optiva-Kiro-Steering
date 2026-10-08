---
inclusion: always
---
<!------------------------------------------------------------------------------------
   Add rules to this file or a short description and have Kiro refine them for you.
   
   Learn about inclusion modes: https://kiro.dev/docs/steering/#inclusion-modes
-------------------------------------------------------------------------------------> 
## ObjProperty

You can use this function for Optiva Workflows, Copy Methods, and Equations.

### Purpose
Retrieves a field from the database and places it in a local variable. You can also use this function to retrieve an array of data.

Syntax for ObjProperty is more complex than other functions. To clarify the documentation, the syntax and examples are identical in Workflow, Copy Method and Equations. Therefore, only the Workflow instance is shown unless an example is specific to Copy Method or Equation.

### Syntax for header data
```vbnet
Dim variable As Object = ObjProperty(PropertyName, [Symbol,Object,RowKey, ColumnKey])
```

### Arguments
You can use the **Validation** form to find the arguments for this function.
* Identify the property by looking at the **VALIDATION_CODE**, **VALIDATION_SUBCODE**, and **FIELD_NAME** columns.
* Look at the **VALIDATION_SUBCODE** column for the selected property to determine the arguments of this function.

### Description
ObjProperty retrieves the value of an object and places it in a local variable. Use this function to gather values for other statements.

You can also use it in conjunction with IsBlank to check if a field value is null before retrieving it.

Workflow examples that follow also apply to copy methods and equations.

### Detail codes
This table lists the detail codes and their corresponding VALIDATION_CODE.

<table>
<thead>
<tr>
<th>Detail Code</th>
<th>Detail Type</th>
<th>VALIDATION_CODE</th>
</tr>
</thead>
<tbody>
<tr>
<td>ATTACH</td>
<td>Attached file or URL</td>
<td>G.ATTACH</td>
</tr>
<tr>
<td>BYPROD</td>
<td>ByProduct</td>
<td>G.FORMBYPROD</td>
</tr>
<tr>
<td>COMP</td>
<td>Composition</td>
<td>G.FORMCOMP</td>
</tr>
<tr>
<td>CONTEXT</td>
<td>Context</td>
<td>G.OBJECTCONTEXT</td>
</tr>
<tr>
<td>DOC</td>
<td>Attached text</td>
<td>G.DOCUMENT</td>
</tr>
<tr>
<td>DTL</td>
<td>File Path</td>
<td>G.FILELOCATIONDTL</td>
</tr>
    <tr>
      <td>INGR</td>
      <td>Ingredient</td>
      <td>G.FORMINGRED</td>
    </tr>
    <tr>
      <td>LRSECT</td>
      <td>Ingredient Statement Rule</td>
      <td>G.LRULESECTION</td>
    </tr>
    <tr>
      <td>LRSUBSECT</td>
      <td></td>
      <td>G.LRULESUBSECTION</td>
    </tr>
    <tr>
      <td>PER</td>
      <td>Security / Permissions</td>
      <td>G.OBJSECURITY</td>
    </tr>
    <tr>
      <td>ST</td>
      <td>Set</td>
      <td>G.SETFORMULA<br>G.SETITEM</td>
    </tr>
    <tr>
      <td>STATUS</td>
      <td>Status</td>
      <td>G.STATUSOBJ</td>
    </tr>
    <tr>
      <td>TP</td>
      <td>Parameter for specifications</td>
      <td>G.SPECPARAM</td>
    </tr>
    <tr>
      <td>TPALL</td>
      <td>Parameter for formulas and items</td>
      <td>G.TECHPVAL</td>
    </tr>
    <tr>
      <td>TP0</td>
      <td>Rollup Parameter</td>
      <td>G.TECHPVAL</td>
    </tr>
    <tr>
      <td>TP1</td>
      <td>Information Parameter</td>
      <td>G.TECHPVAL</td>
    </tr>
    <tr>
      <td>TP2</td>
      <td>Equation Total Parameter</td>
      <td>G.TECHPVAL</td>
    </tr>
    <tr>
      <td>TP3</td>
      <td>Cost Parameter</td>
      <td>G.TECHPVAL</td>
    </tr>
  </tbody>
</table>

You must use these details for PropertyName if the VALIDATION_SUBCODE has a value.

<table>
  <thead>
    <tr>
      <th>Part</th>
      <th>Description</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <td>PropertyName<br>(VALIDATION_SUBCODE is blank)</td>
      <td>FIELD_NAME<br>When the VALIDATION_SUBCODE is blank, use the entry for FIELD_NAME as the PropertyName.</td>
    </tr>
    <tr>
      <td>PropertyName<br>(VALIDATION_SUBCODE has a value)</td>
      <td>FIELD_NAME. DetailCode.<br>If VALIDATION_SUBCODE has a value, then you must use the RowKey, ColumnKey parameters.</td>
    </tr>
    <tr>
      <td>Symbol</td>
      <td>Type of object. Leave out or use empty quotation marks for the current symbol for the workflow or equation, or new object for a copy method.</td>
    </tr>
    <tr>
      <td>Object</td>
      <td>Object for which to obtain a property. Leave out or use empty quotation marks for the current object for the workflow or equation, or new object for the copy method.</td>
    </tr>
    <tr>
      <td>RowKey</td>
      <td>VALIDATION_SUBCODE. Key for the row where the value exists.<br>Include only if VALIDATION_SUBCODE has a value.</td>
    </tr>
    <tr>
      <td>ColumnKey</td>
      <td>In some sections of rows, the FIELD_NAMES are not unique and the VALIDATION_SUBCODE has a value.<br><br>In this case, look for the row where:<ul><li>FIELD_NAME = KEYCODE for formulas</li><li>FIELD_NAME = PARAMCODE for specifications.</li></ul>In that row, take the FIELD_NO value for KEYCODE or PARAMCODE and divide it by 100.<br><br>Use this number for the ColumnKey argument. You can also use the FIELD_NAME KEYCODE or PARAMCODE as the argument.<br><br>Include only if VALIDATION_SUBCODE has a value. For example:<br><br>objProperty("VALUE.TPALL", "", "", "CALCIUM", 2)</td>
    </tr>
  </tbody>
</table>

### Single and array returns

ObjProperty almost always returns a single value. Occasionally, it returns an array.

Some detail data is a single value:

* Status code
* A single parameter
* Approval code
* Security

### Array data for single objects

Some details can have multiple values such as: All parameters or All ingredients. Use parentheses in either of these ways to specify an array return.

```vba
Dim variable1 As Object()
Dim variable() As Object
```

If you enter a value in the RowKey, then a single value is returned. For example, return the value of the PROTEIN parameter.

```vba
dim Protein1 as object = ObjProperty("VALUE.TPALL", "", "", "PROTEIN", 2)
```

The wildcard (*) that is used in the RowKey argument always returns an array.

```vbscript
dim Protein1 as object = ObjProperty("VALUE.TPALL", "", "", "*", 2)
```

Use the LENGTH argument to determine if the array is single or multiple values. Use the array in the For...Next loop to move through the values.

```vbscript
Dim k As Integer
Dim variable1() As Object = ObjProperty(PropertyName, [Symbol,Object,RowKey,ColumnKey])
For k = 0 to variable1.length -1
    'more processing logic here...
Next k
```

### Single and array data for references and contexts

Returns for references and contexts are unique because they can be:

*   A single value for a single object type or context
*   Multiple values for a single object type or context
*   Multiple object types or contexts with values

Use parentheses in either of these ways to specify an array return.

```vbscript
Dim variable1 As Object()
Dim variable() As Object
```

Usage considerations:

*   With references and contexts, you can enter a wildcard (*) for RowKey. In this case, you can get multiple object types or contexts, not only multiple instances of a single object type or context. For example, you can return formulas, specifications, and vendor references.

```vbscript
Dim oRef As Object = ObjProperty("OBJECTCODE.REF", "", "", "*", 2)
```

*   With references or contexts, you can enter a value in the RowKey. In this case, a single object type is returned. There can be more than one instance. There is no way to predict the return value; it can be a single object or an array.
    This example can return one specification or an array of specification references.

```vbscript
Dim oRef As Object = ObjProperty("OBJECTCODE.REF", "", "", "SPECIFICATION", 2)
```

*   You can use the field name instead of the column key.

```vbscript
Dim oRef As Object = ObjProperty("OBJECTCODE.REF", "", "", "SPECIFICATION", "OBJECTTYPE")
```

- Unknown returns can create a problem if your script is configured specifically for the return of a single object or an array. Set up your script to account for the return of either.
- It is best to use the form of ObjProperty that uses the Lookup Code for the reference instead of the Object Type. The FSLOOKUP table contains the correct **Lookup Code** for the desired reference.

Using the Lookup Code always returns an array of values even if there is only one element in the array. In this case, object codes are returned. If there are no references of the designated type, then “Nothing” is returned.

References have a unique solution for filtering object types. You can use a lookup query to specify an object type. This example uses `lookup_code V\REFITEM3` to filter the specification references from the items. The RowKey of “*” ensures an array.

```vbscript
Dim oSpec As Object = ObjProperty("OBJECTCODE.REF.V\REFITEM3", "", "", "*")
```

For references, use the * on RowKey. This enables you to always return an array and to filter for object type. To filter by object type, use a lookup on the PropertyName.

```vbscript
variable = ObjProperty("OBJECTCODE.REF.<LOOKUP_CODE>", Symbol, Object, "RowKey for RefSymbol", ColumnKey)
```

<table>
  <thead>
    <tr>
      <th>Object</th>
      <th>Description</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <td>PropertyName</td>
      <td>For references, use OBJECTCODE.REF and add &lt;LOOKUP_CODE&gt; for the reference type.<br><br>Use one of the LOOKUP_CODE values with the reference property to filter for a specific object type.<br><br>Use * for the RowKey to always return an array of object types. Filter the object type by a lookup on PropertyName.</td>
    </tr>
    <tr>
      <td>Symbol</td>
      <td>Symbol of the object that contains the reference. Use empty quotation marks to indicate the symbol of the object for the workflow.</td>
    </tr>
    <tr>
      <td>ObjectCode</td>
      <td>Object code of the object that contains the reference. Use empty quotation marks to indicate the object for the workflow.</td>
    </tr>
    <tr>
      <td>RowKey for RefSymbol</td>
      <td>Row that contain the reference symbols. Use * to always return an array of all reference symbols (formulas, specifications, vendors, etc.).<br><br>Use it in conjunction with a lookup code for the PropertyName argument to filter the array to a specific object type.</td>
    </tr>
    <tr>
      <td>ColumnKey</td>
      <td>Field<br>100.</td>
      <td>Number of the referenced object, divided by</td>
    </tr>
  </tbody>
</table>

### Array data for views

Retrieving values for columns in views is unique because the columns in the SQL statement are mapped to the generic validation columns (e.g. VIEWCOL1, VIEWCOL2, etc). These validation columns are the arguments for ObjProperty.

In this example:

*   FSLOOKUP contains **LOOKUP_CODE** for the name of the view and **QUERY_CODE** for the name of the query.
    &lt;img&gt;Data in Table 'FSLOOKUP' in 'Optiva52Demo7' on 'US-SBO-DB02\QA'&lt;/img&gt;
*   FSQUERY contains **QUERY_CODE** with the contents of the SQL in **TEXT_QUERY** for the contents of the SQL.
    &lt;img&gt;SQL Server Enterprise Manager - [Data in Table 'FSQUERY' in 'Optiva52Demo7' on 'US-SBO-DB02\QA']&lt;/img&gt;
*   FSVALIDATION contains **FIELD_NAME** for the generic columns mapped to **TEXT_QUERY**.
    &lt;img&gt;Data in Table 'FSVALIDATIONFIELD' in 'Optiva52Demo7' on 'US-SI&lt;/img&gt;

### Array data for the multi-column extension tables

Retrieving values for extension tables is a unique scenario.

*   There can be more than one extension table for an object.  
*   There are multiple columns of data in the table.

### Data type returns with Strict On and Strict Off

The return from the `objProperty` function is an Object because it can return many different types of data. This data includes strings, numbers, dates and arrays of these types.

When you use `CType`, ensure that you are converting to the correct data type. For example, if the `objProperty` function returns an array of integers, you must convert it as such. An error occurs when you attempt to convert to a different data type.

To determine what exactly is returned, you can use the `GetType` method on the returned object.

This syntax prints a message list that indicates the data type that was returned by the system.

```vbscript
Dim oItems As Object = ObjProperty("ITEMCODE.INGR.A", "", "", "*", "")
MessageList("oItems variable is of the type: ", oItems.GetType().Name)
```

The name of the data type is the .Net name for the data type. 

To test for the data type, use the `Typeof` keyword.

```vbscript
Dim oItems As Object = ObjProperty("ITEMCODE.INGR.A", "", "", "*", "")
If Typeof oItems Is Object() Then
    Dim oItemCds() As Object = CType(oItems, Object())
    ' processing logic...
End If
```

### Header data example

*   You can write a script to retrieve the formula class. Because the **VALIDATION_SUBCODE** is blank, only the `PropertyName` argument is required.

<table>
  <thead>
    <tr>
      <th colspan="5">FSVALIDATIONFIELD : Table</th>
    </tr>
    <tr>
      <th></th>
      <th>VALIDATION_CODE</th>
      <th>VALI</th>
      <th>VALIDATION_SUBCODE</th>
      <th>FIELD_NAME</th>
      <th>FIELD_NO</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <td></td>
      <td>FORMULA</td>
      <td>A</td>
      <td></td>
      <td>CLASS</td>
      <td>3300</td>
    </tr>
  </tbody>
</table>

This function statement retrieves the formula class:

```vbscript
Dim oClass As Object = ObjProperty("CLASS")
```

*   You can write a script to retrieve a version number. In this example, the value of `KEYCODE2` is the version number of an object. This value is placed in the local variable `oVersion`.

```vbscript
Dim oVersion As Object = ObjProperty("KEYCODE2", "", "")
```

### Returning single values

These examples are for a single value being returned.

### Retrieving the protein parameter

In this example, the **VALIDATION_SUBCODE** contains a value, PROTEIN. Consequently, all of the arguments are needed.

To get the value for the **COLUMN KEY**:

1. Find the property you want (for example, PROTEIN).
2. Look for **KEYCODE** with a **VALIDATION_CODE** that matches the one for the property.
3. Divide by 100.

<table>
<thead>
<tr>
<th colspan="5">FSVALIDATIONFIELD : Table</th>
</tr>
<tr>
<th></th>
<th>VALIDATION_CODE</th>
<th>VALI</th>
<th>VALIDATION_SUBCODE</th>
<th>FIELD_NAME</th>
<th>FIELD_NO</th>
</tr>
</thead>
<tbody>
<tr>
<td></td>
<td>G.TECHPVAL</td>
<td>A</td>
<td></td>
<td>KEYCODE</td>
<td>200</td>
</tr>
<tr>
<td></td>
<td>G.TECHPVAL</td>
<td>A</td>
<td>PROTEIN</td>
<td>VALUE</td>
<td>300</td>
</tr>
</tbody>
</table>

The function statement to retrieve the value of the Protein parameter is:

```vbscript
Dim oProtein1As Object = ObjProperty("VALUE.TPALL", "", "", "PROTEIN", 2)
Dim dProtein2 As Double = cdbl(oProtein1)
```

### Retrieving a status value

This example places the status (**STATUSIND.STATUS**) of the formula **PIZZA_SAUCE\003** into the local variable **oStatus**.

```vbscript
Dim oStatus As Object = objProperty("STATUSIND.STATUS", "FORMULA", "PIZZA_SAUCE\003")
```

When the status is too low, the user is notified in the alert box.

```vbscript
if (oStatus < 400) then
    MessageList("This formula not approved. Status: ", oStatus)
end if
```

Retrieving the owner and the group

The owner and the group codes are retrieved for the current item (OWNERCODE.PER, GROUPCODE.PER). Those codes are placed into the local variable oOwner, oGroup.

```vbnet
Dim oOwner As Object = ObjProperty("OWNERCODE.PER", "", "", _objectkey, "ITEM_CODE")
Dim oGroup As Object = ObjProperty("GROUPCODE.PER", "", "", _objectkey, "ITEM_CODE")
Return 111
```

Testing for a blank status value

You can do a comparison on a retrieved value. The retrieved value can be blank if a user removed the value or did not enter the value. In this case, add IsBlank to your script. IsBlank specifically tests for blank values.

In this example, the value of the VENDORSTATUS extension field is placed into the local variable oVStat.

```vbnet
Dim oVStat As Object
Dim sVStatStr As String
Dim lMessage As Long
```

Because this is an extension field, the detail syntax is not needed.

```vbnet
Dim oVStat As Object = ObjProperty("VENDORSTATUS","FORMULA", "PIZZA\003")
```

Check if the field is blank or not approved.

```vbnet
if (IsBlank(oVStat ) = 1) Or (oVStat <> "APPROVED" ) then
    MessageList("Vendor Status is not entered.")
Return 9111
end if
```

Retrieving an item status to specify the formula status

This example retrieves the lowest status value of all items in the formula and compares it to the status of the formula. If the status of the formula is greater than the lowest item status, then the formula status is set to the lowest item status value.

**Note:** When the FORMULASTATUS profile attribute value is INGRED CALC, you cannot use an equation to set the formula status to a number ending in 00. See the *Infor PLM for Process Application Configuration Guide*.

```vbnet
Dim oIngrStatus As Object = ObjMethod("", "", "INGRLOWESTSTATUS")
Dim oFStatus As Object = ObjProperty("STATUSIND.STATUS","","")
if (oIngrStatus < oFStatus then
    MessageList("Formula status is lower than item's status.")
    MessageList("Formula status set to lowest ingred status.")
    ObjPropertySet(oingrStatus,0,"STATUSIND.STATUS")
end if
```
If Strict is On, replace the “if” statement that is shown above with this code:

```vbnet
if (CDbl(oIngrStatus) < CDbl(oFStatus)) then
```

### Multiple details returned

ObjProperty returns an array of values if the object has multiple rows that meet the criteria. Or, ObjProperty returns a single string value if the object has only one matching row.

These examples show multiple possible values returned. The wildcard (*) that is used in the RowKey argument always returns an array.

### Retrieving information for formula ingredients

```vbnet
Dim oItems, oTypes, oFormulas, oQtys As Object
Dim oItems As Object = ObjProperty("ITEMCODE.INGR.A", "", "", "*", "")
Dim oTypes As Object = ObjProperty("COMPONENTIND.INGR.A", "", "", "*", "")
Dim oFormulas As Object = ObjProperty("FORMULACODE.INGR.A", "", "", "*", "")
Dim oQtys As Object = ObjProperty("QUANTITY.INGR", "", "", "*")
For k As Integer = 0 to oItems.length-1
    MessageList("Item ", oItems(k), " has quantity ",
        oQtys(k), " and value of formulacode is >", oFormulas(k), "< .")
Next k
```

The script returns this information:

```
01.Item FS-0011 has quantity 54.0541 and the value of formulacode is >FS-0011\0001< .
02.Item SHAPE DOUGH has quantity 1 and the value of formulacode is >< .
03.Item FS-0013 has quantity 8.1081 and the value of formulacode is >FS-0013\0001< .
04.Item 01028 has quantity 26.2688901990516 and the value of formulacode is >< .
05.Item 07057 has quantity 5.4054 and the value of formulacode is >< .
06.Item 11260 has quantity 7.7334331926661 and the value of formulacode is >< .
07.Item 01032 has quantity 8.90249719128252 and the value of formulacode is >< .
08.Item BAKE has quantity 425 and the value of formulacode is >< .
```

### Retrieving document attachments

Here is an example of a single document attachment. An absence of attachments are arrays.

```vbnet
'Retrieve the doc code input for the workflow for this object
Dim funccode As String = WipParamGet("DOCUMENT")
'Retrieve all attachments for this Doc code
Dim ret As Object = objProperty("LINK.ATTACH", "", "", funccode, "DOCCODE")
'Format document codes as an array.
Dim attStr() As String
if typeof ret is String() Then
    attStr = ret
Else if typeof ret is String Then
    attStr = New String() {ret}
Else
attStr = New String() {}
End If
```

### Retrieving specific context attributes

In this example, selling locations are retrieved.

```vbnet
Dim oSell As Object = ObjProperty("ATTRIBVAL.CONTEXT", "", "", "SELLOC", 1)
If TypeOf oSell is String then
    MessageList("Selling Location is ", oSell)
For z As Integer = 0 To oSell.length - 1
    MessageList("Selling Location is ", oSell(z))
Next z
```

The script returns this information:

01. Selling Location is CANADA
02. Selling Location is EU
03. Selling Location is THAILAND
04. Selling Location is UK
05. Selling Location is US

This example demonstrates how to retrieve the various types of context attributes. It also shows how to force the returned values to always be in an array.

```vbnet
Dim RawValue As Object = ObjProperty("ATTRIBVAL.CONTEXT", "", "", "C_BRAND", "ATTRIBCODE")
If IsArray(RawValue) Then
    ContextValues = RawValue
Else
    ReDim ContextValues(0)
    ContextValues(0) = RawValue
End If
```

Replace the fourth argument ("C_BRAND" in the example) with this information:

*   "SELLOC" -- selling location
*   "MFGLOC" -- manufacturing location
*   "C_BRAND" -- brand
*   "C_PRODTYPE" -- product type
*   "C_ENDUSE" -- end use
*   "C_ENDUSER" -- end user

### Retrieving all context values

This example retrieves all of the context values.

```vbnet
Dim oContext As Object= ObjProperty("ATTRIBVAL.CONTEXT", "", "", "*", 1)
Dim oType As Object = ObjProperty("ATTRIBCODE.CONTEXT", "", "", "*", 1)
For z As Integer = 0 to oContext.length-1
    MessageList("Context type ",oType(z), " is ", oContext(z))
Next z
```

The script returns this information:

01.Context type C_BRAND is PIZZA BARN
02.Context type C_ENDUSE is DINNER
03.Context type C_ENDUSER is ALL
04.Context type C_PRODTYPE is PIZZA
05.Context type MFGLOC is BERGEN OP ZOOM
06.Context type MFGLOC is MEXICO CITY
07.Context type MFGLOC is OAKVILLE
08.Context type MFGLOC is XINJIANG
09.Context type SELLOC is CANADA
10.Context type SELLOC is EU
11.Context type SELLOC is US

### Retrieving all references

This example retrieves all references and types from the current object.

```vbnet
Dim oRef As Object = ObjProperty("OBJECTCODE.REF", "", "", "*")
Dim oRefType As Object = ObjProperty("OBJECTTYPE.REF", "", "", "*")
```

Get the reference code for each reference in the formula. This goes through the array of references returned by the "*" argument.

```vbnet
For y As Integer = 0 to oRef.length-1
    MessageList("Referenced ", oReftype(y), " is ", oRef(y))
Next y
```

The script returns this information:

01.Referenced CUSTOMER is AHOLD
02.Referenced LABEL is FS-0008-US\0002.001
03.Referenced LABEL is FS-0008-EU\0002.001
05.Referenced LABEL is FS-0008-FR\0002.001
06.Referenced SPECIFICATION is FS-0008-FOR\0002.001
07.Referenced SPECIFICATION is FS-0008-PKG\0002.001
08.Referenced SPECIFICATION is FS-0008-PRC\0002.001

### Retrieving reference codes for vendors

This example finds all references to vendors in a formula.

```vbscript
Dim oVendors As Object ObjProperty("OBJECTCODE.REF.V\REFFORMULA3", "", "", "*", 2)
```

Get the reference code for each vendor reference in the formula. This goes through the array of vendors returned by the "VENDOR" argument.

```vbscript
For i As Integer = 0 to oVendors.length-1
    Dim oVendorRefCode As Object =
    ObjProperty("REFCODE.REF", "", "", "", oVendors(i), 2)
    oVendorRefCode = oVendorRefCode(0)
Next i
```

### Retrieving view data

In these examples, the name of the view for formulas that reference a project is V\PROJECT1. The first column, "VIEWCOL1", in the FsValidationField table, is the Formula Code. The version is fetched separately.

This syntax retrieves all of the formula codes.

```vbscript
fRaw = ObjProperty("VIEWCOL1.VIEWS.V\PROJECT1", "", "", "*", "")
```

This syntax only retrieves the formula codes for the formulas that have '0001' as the version. The last argument can be the name of the column in the database.

```vbscript
fRaw = ObjProperty("VIEWCOL1.VIEWS.V\PROJECT1", "", "", "0001", "VERSION")
```

Or, it can be the validation field name or number.

```vbscript
fRaw = ObjProperty("VIEWCOL1.VIEWS.V\PROJECT1", "", "", "0001", "VIEWCOL2")
fRaw = ObjProperty("VIEWCOL1.VIEWS.V\PROJECT1", "", "", "0001", "2")
```

### Retrieving extension table values

To retrieve extension table (matrix) values, use this syntax:

```vbscript
mvals = ObjProperty("FIELD<#>.MATRIX.V\DM<symbol#>", Symbol, ObjectKey, RowKey, ColumnKey)
```
| Part | Description |
|---|---|
| Field<#>. MATRI X.V\DM<Symbol#> | The first argument is in three parts: FIELD<#> - Data FIELD name. Data matrix columns are numbered internally. Use FIELD with a number suffix that indicates which column in the grid to retrieve, starting with 1. MATRIX - The name of the detail code. Symbol and number of the Data Matrix table. Data Matrix tables are numbered internally. Use “V\DM” followed by the symbol with a table number suffix, starting with 0. The symbol must match the second argument. Look at the FsDataMatrix table to find the table number. This shows that the Project symbol has two data matrix tables. Use the “VIEW_ID” column as the table number. |
| Symbol | Symbol that matches the symbol in the first argument. The Symbol argument is used with the third argument, the object key argument. Use these fields to define which system object you are accessing. Leave either field blank to indicate the current symbol or object key. |
| ObjectKey | Standard. Object key. |
| RowKey, ColumnK ey | Criteria for the row filter. These arguments enable you to filter the rows instead of returning all of the rows. Use an asterisk for the row key argument to return all of the rows. Or, specify the criteria for which rows to return. To filter rows, use both arguments to define this criteria: |
|  | * The column to filter by |
|  | * The value to filter by. For example, this syntax returns all of the values in the first column where the third column has the word “APPROVED”. ObjProperty("FIELD1.MATRIX.V\DMPROJECT2",""","",""APPROVED","FIELD3") |

This syntax retrieves all of the values in the first column of the third data matrix table.

```vbnet
Dim tbl As Object
ObjProperty("FIELD1.MATRIX.V\DMPROJECT2",""","",""*")
```
### Retrieving a file location

Returns the location of XML files for a report. The value, WEBREPORTS, is a set type in the **File Location** form.
The value, XMLWORK, is the name for the set type that specifies the path to the XML files.

```vbscript
Dim fl As Object = ObjProperty("FILEPATH.DTL", "FILELOCATION",
"WEBREPORTS", "XMLWORK", 1)
MessageList({"XMLWork = ( " +fl.ToString()))
```

### Specifying input variables

This example uses `ObjProperty` to specify the input values.

```
'define input variables
Dim Input1 As String, Input2 As String, Input2A As String, Input2B As String, Input3 As String, Input4 As String
Function wf_start() As Long
    Try
        Input1 = CStr(ObjProperty("..."))
        Input2 = CStr(ObjProperty("..."))
        Input2A = CStr(ObjProperty("..."))
        Input2B = CStr(ObjProperty("..."))
        Input3 = CStr(ObjProperty("..."))
        Input4 = CStr(ObjProperty("..."))
    Catch ex As Exception
        If Not ex.InnerException Is Nothing Then
            Messagelist(ex.InnerException.Message)
        End Try
'params with values are inputs #8 - 13 inclusive so 'empty' values
'must precede the scripted inputs
    LaunchWorkflow("FML_APPROVE", "FORMULA", "", "", "", "", "", "", "", "", "", "",
    "", "", Input1, Input2, Input2A, Input2B, Input3, Input4)
    Return 111
End Function
End Class
```

### Adding a data value as a code segment

This example works with the Script Segment of the Create Rule. The example shows you how to add the description of the `_TASKUSER` as a code segment.

You must use context.returnvalue for the variable of any VB .NET function in the Script segment of a create rule.

```vbscript
Dim context.returnValue As Object
context.returnValue = ObjProperty("DESCRIPTION", "USER", _TASKUSER)
```

Context variables that can be used to obtain values are: `_MODELOBJECTSYMBOL`, `_TASKUSER`, and `_TASKGROUP`

### Retrieving TP values for Label Content object

Retrieving parameter (TP) values for the Label Content object is different than for other objects. For objects other than Label Content, the field names are determined by examining the values in the FsValidationField table.

For the Label Content object, the field names change dynamically based upon the source object. In addition, the list of Analysis Row Tags can be customized by the user. Consequently, there are no rows in the FsValidationField table.

These examples describe the `ObjProperty()` calls that retrieve parameter information when the object that the script is running against is a Label Content object. The label content is associated with four other objects types: Formula, Item, Specification and Analysis.

The syntax for the `ObjProperty()` function varies depending on the type of object that is associated with the label content. The syntax is the same for Formula and Item objects that are associated with the label content.

These examples provide the syntax to use with the `ObjProperty()` function to retrieve the value for a specific technical parameter. In the case of Specification and Analysis, the function retrieves one of multiple parameter values.

*   **Item or Formula**

    Formula and Item objects can only have one value for a technical parameter. Suppose the script has to access the value for the parameter “CALCIUM” and the target object is an Item or Formula. Then the code to `ObjProperty` is shown here.

    ```
    Dim tpVal As Object = ObjProperty("PVALUE.TP", "", "", "CALCIUM", "PARAM_CODE")
    ```

*   **Specification**

    The Specification object supports three parameter values. These are the Min, Target, and Max value. The call to the `ObjProperty` to retrieve the target value for the CALCIUM technical parameter is:

    ```
    Dim tpTargetVal As Object = ObjProperty("PVALUE_TARGET.TP", "", "", "CALCIUM", "PARAM_CODE")
    ```

*   **Analysis**

    For the Analysis object, the call to `ObjProperty` to retrieve the Proposed value for the CALCIUM parameter is:

    ```
    Dim tpProposedVal as Object = ObjProperty("PVALUE_PROPOSED.TP", "", "", "CALCIUM", "PARAM_CODE")
    ```

Suppose you request a TP Value Type that is invalid for the given Label Content object. For example, you request the MIN value for a Label Content object whose source is an Item. Or, you request a RowTag that is not included in the **LABELCONTENT.ROWTAGS** profile attribute. Then, a validation error code of -5008 is returned.

### Examples of Label Content arrays

ObjProperty() can be used to retrieve a set (i.e., an array of values) from the label content’s target object. Retrieving all the parameter values is not useful if the script cannot determine the parameter code that is associated with the parameter value.

To retrieve the set of parameter codes and corresponding parameter values requires calling ObjProperty twice. First to get the set of parameter codes. Second to get the set of parameter values.

<table>
  <thead>
    <tr>
      <th>Object</th>
      <th>Array Example</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <td>Item or Formula</td>
      <td>
        Dim pCodes() As Object = .TP", "", "", "*"<br>
        Dim pVals() As Object = TP", "", "", "*"<br>
        ObjProperty("PARAMCODE<br>
        ObjProperty("PVALUE.
      </td>
    </tr>
    <tr>
      <td>Specification</td>
      <td>
        Dim pCodes() As Object = .TP", "", "", "*"<br>
        Dim pVals() As Object = TARGET.TP", "", "", "*"<br>
        ObjProperty("PARAMCODE<br>
        ObjProperty("PVALUE_
      </td>
    </tr>
    <tr>
      <td>Analysis</td>
      <td>
        Dim pCodes() As Object = .TP", "", "", "*"<br>
        Dim pVals() As Object = PROPOSED.TP", "", "", "*"<br>
        ObjProperty("PARAMCODE<br>
        ObjProperty("PVALUE_
      </td>
    </tr>
  </tbody>
</table>

* The array variable `pCodes()` contains the set of parameter codes. The first parameter that is passed to `ObjProperty` is “PARAMCODE.TP”. The last parameter is “*” indicating the request to return all parameter codes instead of a specific code.
* The array variable `pVals()` contains the set of parameter values. The “*” value indicates that the request is to return all of the parameter values.
* `pCodes()` and `pVals()` are arrays. They hold multiple values. The parameter code `pCodes(1)` value is stored in `pVals(1)` when 1 is used to indicate which value out of the array to access.

## ObjPropertyRemove

You can use this function for Optiva Workflows, Copy Methods, and Equations.

### Purpose

Clears the values of a property.

### Syntax

```
ObjPropertyRemove(SaveMode, PropertyName, Symbol, Object, [Code, ColumnKey])
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
      <td>SaveMode</td>
      <td>Specify `0` as a placeholder. SaveMode is reserved for future use and not currently supported. Changes are saved to the database at the end of the script.</td>
    </tr>
    <tr>
      <td>PropertyName</td>
      <td>See [ObjProperty]</td>
    </tr>
    <tr>
      <td>Symbol</td>
      <td>Type of object. Use empty quotation marks for the current symbol for the workflow.</td>
    </tr>
    <tr>
      <td>Object</td>
      <td>Object for which to obtain a property. Use empty quotation marks for the current object.</td>
    </tr>
    <tr>
      <td>Code</td>
      <td>Optional. Code for the property, for example parameter for `VALUE.TPALL` property.</td>
    </tr>
    <tr>
      <td>SET=parameter set code</td>
      <td>To remove parameter values, assign all parameters that are to be cleared to a set code. Then, specify "SET=set code name" here.</td>
    </tr>
    <tr>
      <td>ColumnKey</td>
      <td>In some sections of rows, the FIELD_NAMES are not unique and the VALIDATION_SUBCODE has a value.<br><br>In this case, look for the row where:<br><br>FIELD_NAME = KEYCODE for formulas<br>FIELD_NAME = PARAMCODE for specifications<br><br>In that row, take the FIELD_No value for KEYCODE or PARAMCODE and divide it by 100.<br>Use this number for the ColumnKey argument. You can also use the FIELD_NAME (KEYCODE or PARAMCODE) as the argument.<br><br>For example, if the ColumnKey is 200, then divide it by 100 and the value is 2.</td>
    </tr>
  </tbody>
</table>

Unlike ObjProperty, Symbol and Object are required arguments. Use empty quotation marks for the workflow’s object or the new object for the copy method.

### Description

ObjPropertyRemove clears the values of a property.

**Note:** You can override security and change a locked or read-only formula using workflow. Use discretion when overriding security.

### Examples

This example removes the Mfg. Item from the Formula > Main tab.

```
ObjPropertyRemove(0, "ITEMCODE", "FORMULA", "")
```

This example clears the value for the date parameter DATEPARAM1.

```
ObjPropertyRemove(0, "VALUE.TPALL", "", "", "DATEPARAM1", 2)
```
You can clear parameter values with this function by assigning the parameters to set codes; then clear all parameters belonging to the set. When a parameter is cleared, the **Level** is reset to **0**.

In this example, the technical parameters are removed from the ALLERGENS set.

```
ObjPropertyRemove(0, "VALUE.TPALL", "", "", "SET=ALLERGENS", 2)
```

## ObjPropertySet

You can use this function for Optiva Workflow, Copy Methods, and Equations.

### Purpose

Sets the value of a field in the database. The syntax is similar to ObjProperty. ObjPropertySet includes an argument for the new value and a save mode.

### Syntax

```
Dim variable As Long = ObjPropertySet(Value, SaveMode, PropertyName, Symbol, Object[, RowKey, ColumnKey])
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
<td>Value</td>
<td>New value for the field.</td>
</tr>
<tr>
<td>SaveMode</td>
<td>Specify **0** as a placeholder. SaveMode is reserved for future use and not currently supported. Changes are saved to the database at the end of the script.</td>
</tr>
<tr>
<td>PropertyName</td>
<td>See [ObjProperty]</td>
</tr>
<tr>
<td>Symbol</td>
<td>Type of object. Use empty quotation marks to indicate the symbol of the current object for the workflow, or the new object for the copy method.</td>
</tr>
<tr>
<td>Object</td>
<td>Object for which to obtain a property. Use empty quotation marks to indicate the current object for the workflow, or the new object for the copy method. For formulas and specifications, you must include a backslash (\\) followed by a version number if you include the object code.</td>
</tr>
<tr>
<td>RowKey</td>
<td>VALIDATION_SUBCODE. Key for the row where the value exists.<br/>Include only if VALIDATION_SUBCODE has a value.</td>
</tr>
<tr>
<td>ColumnKey</td>
<td>Value must be inside quotation marks ("").<br>In some sections of rows, the FIELD_NAMES are not unique and the VALIDATION_SUBCODE has a value. In this case, look for the row where:<br>FIELD_NAME = KEYCODE for formulas<br>FIELD_NAME = PARAMCODE for specifications<br>In that row, take the FIELD_NO value for KEYCODE or PARAMCODE (e.g., 200 or 300) and divide it by 100. Use this number for the ColumnKey argument (e.g., 2 or 1). You can also use the FIELD_NAME (KEYCODE or PARAMCODE) as the argument.</td>
</tr>
</tbody>
</table>

**Description**

Use the ObjPropertySet function to set values of fields in the database. The arguments of this function are the same as for the ObjProperty function. The SaveMode argument is reserved for future use.

When you include Auto Code segments that prompt users to specify a value for the Copy Method, use GetSegData(n). The n represents the Auto Code segment number to use.

**Note:** You can use a workflow to override security and change a locked or read-only formula. Use discretion when overriding security.

**Workflow examples**

In this example, PRIMARYFORMULAIND is one (1) for the default Symbol and ObjectKey, making the object a master formula.

```
ObjPropertySet(1,0,"PRIMARYFORMULAIND","","")
```

For this example, the status of PIZZASAUCE\003 is 300 (engineering). Use STATUSIND, HOLDCODE, APPROVALCODE like other details, not as header data.

```
ObjPropertySet(300,0,"STATUSIND.STATUS","FORMULA", "PIZZASAUCE\003")
```

If this statement were for the workflow’s object, then use empty quotation marks for the symbol and object key.

```
ObjPropertySet(300,0,"STATUSIND.STATUS","","")
```

In the next example, the TOTFAT equation total parameter in TOMATOES is 0.

```
ObjPropertySet(0,0,"VALUE.TP2","ITEM","TOMATOES", "TOTFAT", "2")
```
The relevant rows of the database table are shown here.

<table>
  <thead>
    <tr>
      <th>VALIDATION_CODE</th>
      <th>VALIDATION_RECFMT</th>
      <th>VALIDATION_SUBCODE</th>
      <th>FIELD_NAME</th>
      <th>FIELD_NO</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <td>G.TECHPVAL</td>
      <td>A</td>
      <td></td>
      <td>KEYCODE</td>
      <td>200</td>
    </tr>
    <tr>
      <td>G.TECHPVAL</td>
      <td>A</td>
      <td>TOTFAT</td>
      <td>VALUE</td>
      <td>300</td>
    </tr>
  </tbody>
</table>

In this example, QUALITY is the hold status for PIZZASAUCE\003.

```
ObjPropertySet(QUALITY, 0, "HOLDCODE.STATUS", "FORMULA", "PIZZASAUCE\003")
```

In this example, date is removed by setting an empty string to the date parameter, DATEPARAM1.

```
ObjPropertySet("", 0, "VALUE.TPALL", "", "", "DATEPARAM1", 2)
```

**Marking formulas for rollup**

To mark the object for rollup, the property ROLLUPID is negative one (-1).

```
ObjPropertySet(-1, 0, "ROLLUPID")
```

**Specifying values of ingredients in a formula**

Use `ObjPropertySet` to specify a single value of an ingredient in a formula. In this example, `objPropertySet` sets the amount of an ingredient, ITEM01, in the formula.

Run the workflow on the formula. The property name is of the format:

```
FIELD_NAME.INGR.VALIDATION_RECFMT
```

The FIELD_NAME and VALIDATION_RECFMT come from the FSValidationField table. Look for the VALIDATION_CODE entries of G.FORMINGRED. These are the properties of ingredients in a formula that you can specify using `ObjPropertySet`.

The RowKey is the item code for the ingredient. The ColumnKey is always ITEMCODE.

```
ObjPropertySet(2.43, 0, "QUANTITY.INGR.A", "", "", "ITEM01", "ITEMCODE")
```

**Specifying the Test Order field on the Test form**

The TESTORDERCODE field in the Test form is set to the name of a test order for an existing Test object.

```
ObjPropertySet(TestOrderInstanceName, 0, "TESTORDERCODE", "TESTRESEARCH", "TestDataEntryInstanceName")
```

### Specifying Ref Status on the References tab

In this example, the Approval event is for the action that is associated with the label object. The event updates the **Ref Status** field for a label reference on the **Formula** form.

```vbscript
'Get Keycode
Dim oKeycode As String = Objproperty("KEYCODE", "", "")

'Get Label Version
Dim oVersion As String = Objproperty("KEYCODE2", "", "")
'Join Label Code and Label Version
Dim sLbl As String = oKeycode & "\ " & oVersion

'Get Formula code linked to the label object
Dim oFormula As Object = Objproperty("FORMULACODE.LFORM", "", "", "*")
'Update refstatus field in on the Formula References tab.
ObjPropertySet(40, 0, "REFSTATUS.REF.V\REFFORMULA7", "FORMULA", oFormula, sLbl, "OBJECTCODE")
```

### Specifying the group code for a formula

This example reads a value from the input form and replaces the existing **GROUP_CODE** for a formula.

```vbscript
Dim sGroup As String = WipParamGet("JCS_GROUP")
ObjPropertySet(sGroup, 0, "GROUPCODE.PER", "", "")
Return 111
```

### Copy Method examples

The value of the class field is **CLASS D**.

```vbscript
ObjPropertySet("CLASS D", 0, "CLASS", "", "")
```

You can create an **Auto Code segment** to enable the user to select a class. The class selection is the third **Auto Code segment**:

```vbscript
ObjPropertySet(CopyMethod.Context.GetSegData(3), 0, "CLASS", "", "")
```

### Equation examples

Here are some examples of equations using `objPropertySet`.

This script sets the value of the LABCOST parameter to 50.

```vbscript
Dim lSet As Long
ObjPropertySet(50,0,"VALUE.TP3","","","LABCOST",2)
```

This script retrieves the lowest status of all the ingredients in a formula and the status of the formula. It compares the lowest ingredient status to the formula status. If the ingredient status is lower than the formula status, then the formula is reset to the lowest ingredient status.

```vbscript
Dim oIngrStatus As Object
ObjMethod("", "", "INGRLOWESTSTATUS")
Dim oFStatus As Object = ObjProperty("STATUSIND.STATUS","","")
if (oIngrStatus < oFStatus) then
    MessageList("Ingredient status lower than specified formula status.")
    MessageList("Formula status set to lowest ingredient status.")
    ObjPropertySet(oIngrStatus,0,"STATUSIND.STATUS","","")
end if
```


## ObjMethod

You can use this function for Optiva Workflows, Copy Methods, and Equations.

**Purpose**

Applies a method during a workflow, object copy, or calculation.

**Syntax**

```vbnet
Dim variable As Object = ObjMethod(Symbol, Object, Method, [Parameters])
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
<td>Type of object. Use empty quotation marks for the current symbol for the workflow or equation, or the new object for the copy method.</td>
</tr>
<tr>
<td>Object</td>
<td>Object on which to apply the method. Use empty quotation marks for the current object for the workflow or equation, or the new object for the copy method.</td>
</tr>
<tr>
<td>Method</td>
<td>System method that is applied.</td>
</tr>
<tr>
<td>Parameters</td>
<td>Parameters that are required for the chosen method.</td>
</tr>
</tbody>
</table>

**Description**

ObjMethod applies a system method to the workflow, copy method, or equation. Contact your implementation consultant for information on the methods currently available.

<table>
<thead>
<tr>
<th>Method</th>
<th>Description</th>
</tr>
</thead>
<tbody>
<tr>
<td>CALC</td>
<td>Calculate a formula.</td>
</tr>
    <tr>
      <td>CONSTITUENTROLLUP</td>
      <td>Roll up parameter values from a constituent formula to its raw material item.</td>
    </tr>
    <tr>
      <td>FILTERCHANGE</td>
      <td>Specify a filter for retrieving the parameters in an alternate unit of measure.</td>
    </tr>
    <tr>
      <td>INGRLOWESTSTATUS</td>
      <td>Find the lowest status for a formula.</td>
    </tr>
    <tr>
      <td>SCALE</td>
      <td>Rescale the formula ingredients.</td>
    </tr>
    <tr>
      <td>UOMSTRCONV</td>
      <td>Create a conversion factor. This is used for formula workflows only.</td>
    </tr>
    <tr>
      <td>CALC</td>
      <td>Execute the calculation for an Analysis object. The Symbol is ANALYSIS.RESULT.</td>
    </tr>
    <tr>
      <td>CLEARRESULTS</td>
      <td>Clear results from an analysis.</td>
    </tr>
    <tr>
      <td>GENERATE</td>
      <td>Use for ingredient statements. Apply the Ingredient Statement Rule to generate the Preview text.</td>
    </tr>
    <tr>
      <td>PREVIEW</td>
      <td>Update the Preview text with edits from the **Ingredient Statement** > **Ingredients** tab.</td>
    </tr>
    <tr>
      <td>SWITCHLANGUAGE</td>
      <td>Switch the language of an ingredient statement without altering the current order of the ingredients.</td>
    </tr>
    <tr>
      <td>LOADSOURCETP</td>
      <td>Use for **Label Content** objects. Execute the REFRESH on the **Parameters** tab.</td>
    </tr>
    <tr>
      <td>UPDATELABELTEXT</td>
      <td>Update the Ingredient Statement in the **Label Content** form.</td>
    </tr>
    <tr>
      <td>SETPASSWORD</td>
      <td>Assign a password to a user.</td>
    </tr>
    <tr>
      <td>EXECUTE</td>
      <td>Use for Web Reports. Execute a quick search and add the results detail object to the dataset for a report.</td>
    </tr>
    <tr>
      <td>RECALL</td>
      <td>Cancel a workflow after launching it.</td>
    </tr>
  </tbody>
</table>

### Secured scripting using ObjMethod

With secured scripting, you can specify only these methods as the third parameter in an ObjMethod call.

---

<table>
<thead>
<tr>
<th>Method Name</th>
<th>Example</th>
</tr>
</thead>
<tbody>
<tr>
<td>CurrentExternalCalculateDates</td>
<td>ObjMethod("", "", "CurrentExternalCalculateDates", arglist)</td>
</tr>
<tr>
<td>ObjectMethod</td>
<td>ObjMethod("", "", "ObjectMethod", "CLEARRESULTS", Nothing)</td>
</tr>
<tr>
<td>UomStrConv</td>
<td>Dim oFactor As Object = ObjMethod("", "", "UomStrConv", oUom, "KG")</td>
</tr>
<tr>
<td>IngrLowestStatus</td>
<td>Dim oLowStatus As Object = ObjMethod("", "", "IngrLowestStatus")</td>
</tr>
<tr>
<td>AddOrderDetails</td>
<td>ObjMethod("TESTORDER", stoCode, "AddOrderDetails", argsSC, argsTC)</td>
</tr>
<tr>
<td>ParentCode</td>
<td>Dim oParent() As String = ObjMethod("", "", "ParentCode")</td>
</tr>
<tr>
<td>SwitchLanguage</td>
<td>ObjMethod("", "", "SwitchLanguage", "FR-FR", "BEV")</td>
</tr>
<tr>
<td>Calc</td>
<td>ObjMethod("FORMULA", _OBJECTKEY, "CALC", 0, 2)</td>
</tr>
<tr>
<td>get_Sid</td>
<td>Dim Mspecid As Long = ObjMethod("SPECIFICATION", "", "get_Sid")</td>
</tr>
</tbody>
</table>

**Controlling the start time of an Optiva workflow**

The CurrentExternalCalculateDates method is used for Optiva workflows. This method is only used for the WIP ID that is currently running.

This method only requires the date that you are using for the start time. The DataSet is not required.

In this example, the start date is set to the **Effective Start Date** of the project; then the WIP ID is reassigned to FRADMIN.

```vbnet
Dim sStartDate As Object = ObjProperty("EFF_START_DATE", "", "")
If IsBlank(sStartDate) = 1 Then
    Messagelist("The project start date has not been set. Therefore the workflow cannot be started.")
    Return 9111
End If

Dim arglist As datetime = Ctype(sStartDate, DateTime)
ObjMethod("", "", "CurrentExternalCalculateDates", arglist)

Reassign("FRADMIN", "", "")
```

### Calling ObjectMethod from ObjMethod

ObjectMethod is one of the methods that can be called from objMethod.

ObjectMethod is specified as the Method argument. This is a generic method that can perform any of several specific methods depending on the object on which ObjectMethod is applied.

When you use objMethod to call ObjectMethod, you must have these entries as the Parameter arguments:

*   The name of the method to run, such as: Calc or ClearResults
*   Arguments for that particular method.

### Examples

```vbscript
Dim oResults As DataSet = ObjMethod("", "", "ObjectMethod", "CLEARRESULTS", Nothing)
Dim oScale As Object = ObjMethod("", "", "ObjectMethod", "Scale",
"newyield=50;newuom=GM;rsmode=1")
```

### Using conversion factors for Formulation

The UOMSTRCONV method is used for formula workflows only.

In this example, the UOMSTRCONV method creates a conversion factor. The parameters that follow the method are the unit of measure.

*   (oUom) is the unit of measure that is converted from.
*   ("KG") is the unit of measure that the number is converted to.

```vbscript
Dim oUom As Object = ObjProperty("UOMCODE")
Dim oFactor As Object = ObjMethod("", "", "UOMSTRCONV", oUom, "KG")
Dim oYield As Object = ObjProperty("YIELD") * oFactor
Dim oMethod As Object = ObjMethod("XGMITEM", oItemkey, "setdefaults",
oFormula, oVersion, oYield, oProcessYield)
```

### Finding the lowest status for a formula or formula ingredients

In this example, the status of each item in a formula is compared. The item with the lowest status is returned to the variable `oLowStatus`.

```vbscript
Dim oLowStatus As Object = ObjMethod("", "", "INGRLOWESTSTATUS")
```

### Scaling formula ingredients

This example scales a formula from 1000 KG to a 50 gram sample.

The argument `rsmode = 1` changes ingredient UOMs to “newUOM”.

The argument `rsmode=0` does not change ingredient UOMs.

```vbscript
Dim oScale As Object = ObjMethod("", "", "ObjectMethod", "Scale",
"newyield=50;newuom=GM;rsmode=1")
```

### Calculating a formula

In this example, the `Calc` method, is applied to a formula.

```vbscript
Dim oCalc as Object = ObjMethod("FORMULA", "", "ObjectMethod", "CALC",
"mode=0;calcflags=0")
```

<table>
<thead>
<tr>
<th>Parameter</th>
<th>Flag</th>
<th>Description</th>
</tr>
</thead>
<tbody>
<tr>
<td>Mode</td>
<td><strong>0</strong></td>
<td>Not currently used by the system. The value is <strong>0</strong>.</td>
</tr>
<tr>
<td>CalcFlags</td>
<td><strong>0=As Necessary</strong></td>
<td>First-level calculation. Conditional. A formula recalculation is determined by the system. This occurs when there has been a change in ingredients or parameter values. The calc is performed only if there is a change.</td>
</tr>
<tr>
<td>CalcFlags</td>
<td><strong>1=CostOnly</strong></td>
<td>First-level calculation.<br>Roll up costs only</td>
</tr>
<tr>
<td>CalcFlags</td>
<td><strong>2=ForceAll</strong></td>
<td>First-level calculation. Unconditional. Calculation occurs regardless of whether there are changes to the object.</td>
</tr>
<tr>
<td>CalcFlags</td>
<td><strong>4=CalcAll</strong></td>
<td>All-level calculation. Calculate all of the sub-formulas before calculating the selected formula.</td>
</tr>
</tbody>
</table>

### Obtaining the parent code

This example, used in a workflow script, retrieves the parent code of an object such as a formula or project.

```vbscript
Dim oPCode As Object = ObjMethod("", "", "ParentCode")
```

It returns a string array containing the components of the parent's key. The result is either "Code" or "Code\Version".

```vbscript
Dim sPCode As String = String.Join("\", oPCode)
```

If no parent exists for the current object, then the ParentCode method returns Nothing.

### Changing the filter to retrieve parameters

The data set only supports a single PVALUE node. You get the node of whatever filter is currently active.

By default, the PVALUE uses the UOM of the parameter. If you use the FILTERCHANGE object method, you can switch to the alternate UOM for that filter.

The FILTERCHANGE object method requires the parameter detail code. Therefore, there must be an initial call to ObjectXML. If your workflow is already retrieving parameter values, then you may not require that step.

```vbscript
ObjectXML("", "", "HEADER;TPALL")
ObjMethod("", "", "ObjectMethod", "FILTERCHANGE", "TPALL;MY_FILTER")
Dim xml As String = ObjectXML("", "", "HEADER;TPALL")
MessageList(xml)
Return 111
```

### Rolling up constituent values

Rolls up the parameter values of a constituent formula to the same parameters in the raw material item. In this example, the ConstituentRollup method is applied to a post-save script on the **Formula Symbol > Script** tab.

The script runs when a constituent formula is saved after being created or updated. The Save event occurs after the formula is calculated and its parameter values are updated by the standard rollup process.

This example adds “ObjectMethod” as the third argument.

```vbscript
Option Strict Off
Imports System
Imports System.Data
Imports System.Diagnostics

Class HookScript
    Inherits FcProcFuncSetEventHook
Function postsave() As Long
    Dim oRollup as Object = ObjMethod("", "", "ObjectMethod", "ConstituentRollup", "")
Return 1
End Function
End Class
```

This object method does not work properly in an equation because all of the item parameters are copied. Equations are part of the calc cycle. Therefore, the copy of the parameters cannot be performed until the calc cycle is over.

This object method can also be invoked from a workflow. If a workflow is running on the constituent formula, then the syntax is:

```vbscript
Dim rc As Object = ObjMethod("", "", "ObjectMethod", "CONSTITUENTROLLUP",
"")
```

If the workflow is running on another object, then the symbol is the first argument. The key is the second.
The rest of the arguments are always the same.

```vbnet
Dim rc As Object = ObjMethod("FORMULA","F1\0001","ObjectMethod",
"CONSTITUENTROLLUP", "")
```

The **Level** column on the **Item > Parameters tab** is affected according to these rules.

<table>
<thead>
<tr>
<th>Level before Constituent Rollup</th>
<th>Update Level?</th>
<th>Description</th>
</tr>
</thead>
<tbody>
<tr>
<td><strong>0 = Pending</strong></td>
<td>Yes</td>
<td>Overwrite the item value with the value of the Constituent Formula.</td>
</tr>
<tr>
<td><strong>1 = Manual Override</strong></td>
<td>No</td>
<td>Do not change the parameter value.</td>
</tr>
<tr>
<td><strong>2 = Calculated</strong></td>
<td>No</td>
<td>Do not change the parameter value. Subsequent Save or Calc on the Item updates this value; the update is based on the parameter definitions.</td>
</tr>
<tr>
<td><strong>3 = Derived from Constituents</strong></td>
<td>Yes</td>
<td>Update the item value with the Constituent Formula</td>
</tr>
</tbody>
</table>

### Calculating Analysis results

CALC can be used to calculate and populate the results of the **Analysis** object.

```vbnet
ObjMethod("ANALYSIS.RESULT", _ObjectKey, "ObjectMethod", "CALC", Nothing)
```

### Clearing results from an Analysis

The **CLEARRESULTS** method of the **ObjMethod** function can be used to clear all values from the results section of an analysis.

In this example, the **CLEARRESULTS** method is applied to an analysis to clear all results values.

```vbnet
Dim oResults As DataSet = ObjMethod("", "", "ObjectMethod", "CLEARRESULTS", Nothing)
```

To invoke this from the **Analysis** form directly, this profile can be used for a custom button on the **Analysis** form:

```plaintext
Clear Results;AE;CLRANALYRES;0^{%OBJECTSYMBOL};{%SKEY}
```

Then an action called **CLRANALYRES** can be created. Add the previous workflow example to the start event of the action.

This command can also be invoked from the calcprecalc VB script hook.

```vbnet
Public Function calcprecalc()
Dim oResults As DataSet = ObjMethod("", "", "ObjectMethod", "CLEARRESULTS", Nothing)
End Function
```

Executing calculations

This example executes a formula calculation.

```vbnet
Dim oCalc as Object = ObjMethod("FORMULA","","ObjectMethod", "CALC", "mode=0;calcflags=4")
```

Calling Global Calc from ObjMethod

You can use EXECFORMULACALC to run a workflow on a Global Calc object.

```vbnet
ObjMethod("", "", "Method", "EXECFORMULACALC", "", Nothing)
```

You can also use EXECFORMULACALC to run a Global Calc job.

```vbnet
ObjMethod("GLOBALCALC", "JOB1", "Method", "EXECFORMULACALC", "", Nothing)
```

Applying an ingredient statement rule

You can apply an ingredient statement rule to an ingredient statement and generate the preview text.

```vbnet
ObjMethod("LABEL", _Objectkey, "ObjectMethod", "GENERATE", "")
```

Updating the preview in an ingredient statement

This example applies edits from the **Ingredient Statement > Ingredients** tab to the **Preview** tab.

```vbnet
ObjMethod("LABEL", _OBJECTKEY, "ObjectMethod", "PREVIEW")
```

Using SWITCHLANGUAGE for ingredient statements

You can use the SWITCHLANGUAGE ObjMethod to change the language of an ingredient statement without altering the current order of the ingredients.

This ObjMethod performs these steps:

1. Takes an argument list with two parameters, language and category (e.g., language=FR-CA and category=BEV).
2. Changes the language and category of the current object.
3. Moves through the list of rows in the Ingredients detail code.
4. Retrieves the appropriate text in the new language. This is the same logic that was used during the original Ingredient Statement explosion.
5. Replaces the old language on the **Ingredients** grid.

There are two ways to call this object method through scripting. Both examples change the language to French (FR-FR) and the category to BEV.

```vba
ObjMethod("", "", "ObjectMethod", "SWITCHLANGUAGE", "language=FR-FR; category=BEV")
ObjMethod("", "", "SwitchLanguage", "FR-FR", "BEV")
```

### Updating the label content

You can use ObjectMethod to update the Label Ingredient Statements in the **Label Content** form.

The syntax is:

```vba
Dim UpdateText as Object = ObjMethod("LABELCONTENT", _OBJECTKEY, "ObjectMethod", "UPDATELABELTEXT", Nothing)
```

This example demonstrates how to execute the **Update** and **Refresh** buttons in the **Label Content** form.

```vba
Function wf_start() As Long
    Dim UpdateText as Object = ObjMethod("LABELCONTENT", _OBJECTKEY, "ObjectMethod", "UPDATELABELTEXT", Nothing)
    Messagelist("Update Complete")
    Return 111
End Function
```

This example executes the REFRESH on the **Parameters** tab of the **Label Content** object.

```vba
Dim IcmTP As Object = ObjMethod("LABELCONTENT", _OBJECTKEY, "ObjectMethod", "LOADSOURCETP", Nothing)
Messagelist("Refresh Complete")
Return 111
End Function
```

### Refreshing parameter data

```vba
ObjMethod(_objectsymbol, _objectkey, "ObjectMethod", "LOADSOURCETP", Nothing)
```

### Including Quick Search results in Web Reports

You can run WebReports from a quick search. The results from quick searches are not added to the dataset or into the XSD until the search is executed from the workflow. You execute the search from the workflow using the EXECUTE object method.

The syntax is:

```vbscript
Dim rc as DataSet = ObjMethod(symbol, _OBJECTKEY, "ObjectMethod", "EXECUTE", "")
```

This workflow example demonstrates how to add search results data to a web report.

```vbscript
' This executes the search and adds Result detail object to the dataset.
Dim dsQS As DataSet = ObjMethod("FORMULASEARCH.HEADER;CRITERIA",
    _ObjectKey "ObjectMethod", "EXECUTE", "")
' To export the XML for the search object, add this line after the
ObjMethod call.
Dim XMLStr As String = ObjectXML("", "", "HEADER;SEARCHRESULT")
```

### Recalling (Cancel) a workflow

You can recall (cancel) a workflow after it has been launched. Use the RECALL object method on an action set object.

The syntax is:

```vbscript
Dim rc as DataSet = ObjMethod(symbol, code, "ObjectMethod", "RECALL")
```

This workflow cancels the selected workflow from the **Views** tab of a formula:

```vbscript
Function wf_start() As Long
    ObjMethod("", "", "ObjectMethod", "RECALL", "")
    MessageList("Leaving RECALL-02. _ObjectKey = ", _ObjectKey)
End Function
```

The RECALL performs the same steps that the **Workflow Maintenance** form performs when the workflow status is 3-Canceled. Each WIP step is given a status of 2-Closed.

You can send email to notify the user that the formula approval has been recalled.

Set up a View for workflows to see the workflows that have been launched against an object. Then, you can launch any workflow for a selected workflow-in-progress. For example, use this query for a formula view.

```sql
SELECT distinct W.ActionWIP_ID, A.ActionSet_Code, A.Description, S.Due_date
FROM FsActionWIP W, FsActionWIPSteps S, FsActionSet A, fsFormula F
WHERE S.ActionWIP_ID = W.ActionWIP_ID
AND W.ActionSet_code = A.ActionSet_code
AND F.Formula_id = [%1]
AND S.TARGET_OBJECT_KEY = F.Formula_code + '\ ' + F.Version
```

From the **Reports** tab, you can launch a Recall workflow for a workflow-in-progress instead of for an object. Select a workflow and click **Launch Workflow**. The **Launch Workflow** form opens normally, as if it were launched for a business object.

Select a Recall Workflow action set. After the Recall workflow has run on WIP ID 5330, the status of the workflow is **Cancelled**.

&lt;img&gt;Workflows in Progress 5330 : Formula Approval Workflow - Main Status: Cancelled (Not Started, Started, Suspended, Cancelled, Completed)&lt;/img&gt;

The Queue Status of the step is **Closed**.

&lt;img&gt;Workflows in Progress 5330 : Formula Approval Workflow - Detail Action Code: CHECK FORMULA Object Key: EX-003\0001 Object Symbol: FORMULA Queue Status: Closed REGULATORY REVIEW EX-003\0001 FORMULA Closed Not Started In Process Queue Closed&lt;/img&gt;

You can edit the workflow-in-progress. In the Object page, select **Reports tab > All Workflows**. Then, click the **Workflow ID** link to open the **Workflows in Progress** page.


