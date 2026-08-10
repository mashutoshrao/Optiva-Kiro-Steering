---
inclusion: always
---
<!------------------------------------------------------------------------------------
   Add rules to this file or a short description and have Kiro refine them for you.
   
   Learn about inclusion modes: https://kiro.dev/docs/steering/#inclusion-modes
-------------------------------------------------------------------------------------> 

# DataSet objects

You can access the **DataSet** object using the **ObjMethod** function. A more efficient and robust method is to use these commands:

<table>
<thead>
<tr>
<th>Command</th>
<th>Description</th>
</tr>
</thead>
<tbody>
<tr>
<td>ObjectDataSet</td>
<td>Retrieves the **ADO.Net** object.</td>
<td>**DataSet** object for the requested</td>
</tr>
<tr>
<td>DataSetTableName</td>
<td>Retrieves the logical table name of the requested detail code. This facilitates the retrieval of the data from the **ADO.Net** **DataSet** object.</td>
<td></td>
</tr>
<tr>
<td>GetNewRow</td>
<td>Returns a blank row for the requested detail code. If the detail code has an auto-sequenced line ID column, the next row number value, for the new row, is pre-filled.<br><br>Some detail codes, such as CUSTOM, do not support the addition of new rows. Calling this method on those detail codes throws an exception.</td>
<td></td>
</tr>
<tr>
<td>CommitNewRow</td>
<td>Adds the new row that is returned from the **GetNewRow** command to the data table. Rows can be committed to the table only if each column that requires a value has a value.<br><br>When a row with same key value already exists in the table as the new row, an exception is thrown.</td>
<td></td>
</tr>
<tr>
<td>RowUpdate</td>
<td>Enables the script to call the same **RowUpdate** logic that the user interface would use for a given detail row. This command is optional. It is useful, especially for the Formula Ingredient detail, because the new row’s **Quantity %** column is calculated.</td>
<td></td>
</tr>
</tbody>
</table>

These commands offer more direct access to the underlying data objects that Optiva utilizes internally. They can be used to complete these tasks:

*   Add rows to any Optiva object.
*   Access the **DataSet** object in a workflow script.
*   Import ION information to Optiva.

## CommitNewRow

Adds the new row that is returned from the GetNewRow command to the data table. Rows can be committed to the table only if each column that requires a value has a value. If a row with same key value already exists in the table as the new row, then an exception is thrown.

### Syntax

```csharp
Public Sub CommitNewRow(ByVal objSymbol As String, ByVal objKey As String, ByVal DetailCode As String, ByRef RowToAdd As DataRow)
```

### Arguments

<table>
  <thead>
    <tr>
      <th>Argument</th>
      <th>Description</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <td>objSymbol</td>
      <td>The data object type or the current object type if blank.</td>
    </tr>
    <tr>
      <td>objKey</td>
      <td>The data object key or the current object key if blank.</td>
    </tr>
    <tr>
      <td>DetailCode</td>
      <td>The system detail code that is used to add the new data row. Examples: “INGR” for Formula Ingredients; “HEADER”<br>If the detail code does not exist, or does not support rows being added to it (e.g., the VIEW detail), then an exception is thrown.</td>
    </tr>
  </tbody>
</table>

## DataSetTableName

Retrieves the logical table name of the requested detail code. This facilitates the retrieval of the data from the ADO.Net DataSet Object.

### Syntax

```csharp
Public Function DataSetTableName(ByVal objSymbol As String, ByVal objKey As String, ByVal dtlCode As String) As String
```

### Arguments

<table>
  <thead>
    <tr>
      <th>Argument</th>
      <th>Description</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <td>objSymbol</td>
      <td>The data object type or the current object type if blank.</td>
    </tr>
    <tr>
      <td>objKey</td>
      <td>The data object key or the current object key if blank.</td>
    </tr>
    <tr>
      <td>dtlCode</td>
      <td>The system detail code that is used to return the table name information.<br>Examples: “INGR” for Formula Ingredients; “HEADER”; “VIEW.V\VIEWINFO1”</td>
    </tr>
  </tbody>
</table>

Return The string name of the table in the DataSet. Normally, the name of the table in the DataSet is the same as the table name in the database.

## GetNewRow
Returns a blank row for the requested detail code. Suppose the detail code has an auto-sequence line ID column; then the value of the next row number is assigned to the new row automatically.

### Syntax
```vbnet
Public Function GetNewRow(ByVal objSymbol As String, ByVal objKey As String, ByVal DetailCode As String) As DataRow
```

### Arguments
<table>
  <thead>
    <tr>
      <th>Argument</th>
      <th>Description</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <td>objSymbol</td>
      <td>The data object type or the current object type if blank.</td>
    </tr>
    <tr>
      <td>objKey</td>
      <td>The data object key or the current object key if blank.</td>
    </tr>
    <tr>
      <td>DetailCode</td>
      <td>The system detail code that is used to return the new row.<br>Examples: “INGR” for Formula Ingredients; “HEADER”; “ST”<br>When the detail code does not exist, an exception is thrown.<br>Some detail codes (such as CUSTOM or VIEW) do not support the addition of new rows. Calling this method on those detail codes throws an exception.</td>
    </tr>
  </tbody>
</table>

Return ADO.Net DataRow object. See http://msdn.microsoft.com

## ObjectDataSet
Retrieves the ADO.Net DataSet object for the requested object.

### Syntax
```vbnet
Public Function ObjectDataSet(ByVal objSymbol As String, ByVal objKey As String, Optional ByVal
```
dtls As String = "HEADER") As DataSet

### Arguments

<table>
<thead>
<tr>
<th>Argument</th>
<th>Description</th>
</tr>
</thead>
<tbody>
<tr>
<td>objSymbol</td>
<td>The data object type or the current object type if blank.</td>
</tr>
<tr>
<td>objKey</td>
<td>The data object key or the current object key if blank.</td>
</tr>
<tr>
<td>Dtls</td>
<td>Semi-colon delimited list of detail codes to ensure that they are loaded in the returned **DataSet** object.</td>
</tr>
</tbody>
</table>

Return ADO.Net **DataSet** object. See [http://msdn.microsoft.com](http://msdn.microsoft.com).

## RowUpdate

Enables the script to call the same **RowUpdate** logic that the system user interface utilizes for a given detail row. This command is optional. It is useful, especially for the Formula Ingredient detail because the **Quantity %** column is calculated for the new row.

Only use RowUpdate for symbols or details that support RowUpdate events in hook scripts. For example, ingrpre(and post)rowupdate on formula Item lines, tppre(and post)rowupdate on Spec Parameters.

### Syntax

```vbnet
Public Function RowUpdate(ByVal objSymbol As String, ByVal objKey As String, ByVal DetailCode As String, ByVal updateRow As DataRow, Optional ByVal originalRow As DataRow = Nothing, Optional ByVal flags As Integer = 0) As DataSet
```

### Arguments

<table>
<thead>
<tr>
<th>Argument</th>
<th>Description</th>
</tr>
</thead>
<tbody>
<tr>
<td>objSymbol</td>
<td>The data object type or the current object type if blank.</td>
</tr>
<tr>
<td>objKey</td>
<td>The data object key or the current object key if blank.</td>
</tr>
<tr>
<td>DetailCode</td>
<td>The system detail code that has the **RowUpdate** logic to execute.</td>
</tr>
<tr>
<td>updateRow</td>
<td>The ADO.Net **DataRow** object that has the current row values.</td>
</tr>
<tr>
<td>originalRow</td>
<td>Optional. The ADO.Net **DataRow** object that has the old row values.</td>
</tr>
<tr>
<td>Flags</td>
<td>Reserved for future use.</td>
</tr>
</tbody>
</table>

Return ADO.Net DataSet object. See http://msdn.microsoft.com.

### Example and workaround for updating formula ingredients and passing the original row

**Note:** An error message may be encountered when a workflow uses RowUpdate on formula ingredients, while the Formula symbol pre/post Ingredient Row Update event includes Context.GetCurrentRowValue.

When a user is working on the Item Lines of a formula, if a pre/post row update **Symbol** script needs the original row values (Context.GetCurrentRowValue), the application has passed it to the script and there is no error. However, if a modification comes from a workflow action, the original row value is not normally passed. Then, when the RowUpdate() Workflow function causes the pre/post row update to run, the original row value is missing. You may encounter an error such as “Cannot perform '=' operation on System.Double and System.String". This condition can be avoided with a simple modification:

### Example Symbol script excerpt:

```vbnet
Dim currVals As Object = Context.GetCurrentRowValue("UOM_CODE")
```

### Example Action script excerpt:

```vbnet
Dim IngrTableName As String = DataSetTableName("FORMULA", _OBJECTKEY, "INGR")
Dim rows as DataRow
Dim ds As DataSet = ObjectDataSet("FORMULA", _OBJECTKEY)
Dim IngrTable As DataTable = ds.Tables(IngrTableName)
Dim rowcnt As Integer=IngrTable.Rows.Count-1
while k <= rowcnt
    rows=IngrTable.Rows(k)
    rows("UOM_CODE")="KG"
    rows("ATTRIBUTE3")="A"
    RowUpdate("FORMULA", _OBJECTKEY, "INGR", rows)
```

In the **Action** script, Use the workaround below to in place of the last line of code above:

```vbnet
Dim origRow as DataRow = rows.Table.NewRow
origRow.ItemArray = rows.ItemArray
RowUpdate("FORMULA", _OBJECTKEY, "INGR", rows, origRow)
```

### Formula ingredients example

This example demonstrates how to add a formula ingredient.

```vbnet
Option Strict Off
Imports System
Imports System.Data
Imports System.Diagnostics

Class ActionScript
    Inherits FcProcFuncSetEventWF

    Function wf_start() As Long
        Dim ds As DataSet = ObjectDataSet("", "")

Dim IngrTableName As String = DataSetTableName("", "", "INGR")
Dim IngrTable As DataTable = ds.Tables(IngrTableName)

Dim newIngrRow As DataRow = GetNewRow("", "", "INGR")
newIngrRow("ITEM_CODE") = "01001"
newIngrRow("QUANTITY") = 15.5
newIngrRow("UOM_CODE") = "LB"
RowUpdate("", "", "INGR", newIngrRow)

CommitNewRow("", "", "INGR", newIngrRow)

Return 111
End Function
End Class
```

### Extension table example

This example demonstrates how to add rows to an extension table.
```
Option Strict Off
Imports System
Imports System.Data
Imports System.Diagnostics

Class ActionScript
    Inherits FcProcFuncSetEventWF

Function wf_start() As Long

    Dim newDM0Row As DataRow = GetNewRow("", "", "MATRIX.V\DMITEM0")
    newDM0Row("FIELD2") = "new row added by WF "
    newDM0Row("FIELD3") = DateTime.Now
    newDM0Row("FIELD4") = 879.654
    CommitNewRow("", "", "MATRIX.V\DMITEM0", newDM0Row)

    Dim newDM1Row As DataRow = GetNewRow("", "", "MATRIX.V\DMITEM1")
    newDM1Row("FIELD2") = 456.123
    newDM1Row("FIELD3") = DateTime.Now
    CommitNewRow("", "", "MATRIX.V\DMITEM1", newDM1Row)

    Return 111
End Function
End Class
```

You can create an action set and an action script that specifies the extension tables to display in a wizard.
Your action script can include a create rule that copies the data in the wizard extension table to a new object.
