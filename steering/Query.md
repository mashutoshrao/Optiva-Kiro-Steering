---
inclusion: auto
name: Security
description: Use this file if Optiva needs to call a query for a task, equation, script library
---
<!------------------------------------------------------------------------------------
   Add rules to this file or a short description and have Kiro refine them for you.
   
   Learn about inclusion modes: https://kiro.dev/docs/steering/#inclusion-modes
-------------------------------------------------------------------------------------> 

# Common arguments used while writing a query

The common arguments for the ResultTableRead and ResultTableUpdate are listed here.

*   QueryCode = The entry in FsQuery that has the SQL statement to execute.
*   QueryArgs = A list of replacements values to be used when executing the SQL. This list replaces the numbered tokens in the SQL, i.e., [%1], [%2], etc.
*   QueryNamedArgs = A list of named replacement values to be used when executing the SQL. This list replaces the named tokens in the SQL, i.e., [%%NAMED_TOKEN1], [%%NAMED_TOKEN2]
*   The query is scanned by the system and the list of named tokens is replaced.
    *   %%SEARCHTABLE – the name of the current Search Result table
    *   %%KEYFIELD – the name of the primary key field for the current search object
    *   %USER – the code of the current user
    *   %LAB – the code of the current lab
    *   %GROUP – the code of the current group
    *   %%CURRENT_DB_NAME – the code of the current database connection
    *   %%CURRENT_SESSION_ID – the current session ID value
*   Public ReadOnly Property SearchedKeyField As String
    *   Returns the name of the primary key field for the current search object. (i.e., "ITEM_CODE" when searching for Items, "FORMULA_ID" when searching for Formulas)

*   Public Property SearchSQL() As String
    Enables the script to retrieve and update the SQL that is used to execute the search.
    *   Returns Nothing during the PreSearch script hook.
    *   Updating the SQL during the PreSearch script hook is ignored.
    *   Updating the SQL during the PreSearchExecute script hook is permitted.
    *   Updating the SQL during the PostSearch script hook is ignored.


## TableLookup

You can use this function for Optiva Workflows, Copy Methods, and Equations.

### Purpose

Execute a SQL Query with the supplied parameters. Returns the value from the first row and column of an array from the query. The object is a valid data type for the function.

The TableLookup function uses the same standard name query arguments as described for TableLookupEx and TableReader.

This function is best for a quick single value answer.

### Syntax

```vbnet
Dim variable As Object = TableLookup(QueryName[, paramN])
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
      <td>QueryName</td>
      <td>Name of the query in FSQUERY to be executed. The tokens are replaced in the SQL by the system for these names:
        <ul>
          <li>[%USER]</li>
          <li>[%LAB]</li>
          <li>[%GROUP]</li>
          <li>[%%CURRENT_DB_NAME]</li>
          <li>[%%SEARCHTABLE] - This token is only available during the Search's PostSearch event.</li>
        </ul>
      </td>
    </tr>
    <tr>
      <td>paramN</td>
      <td>Optional. The parameters can be in the form of an array or a comma-delimited string of the individual arguments of the query.</td>
    </tr>
  </tbody>
</table>

### Examples

```vbnet
Dim rv As Object
TableLookup("MyCustomQueryCode", arg1, arg2, arg3...)
```
or

```vbnet
Dim argArray As Object = array(arg1, arg2, arg3)
Dim rv As Object = TableLookup("MyCustomQueryCode", argArray)
```

## TableLookupEx

You can use this function for Optiva Workflows, Copy Methods, and Equations.

### Purpose

Executes a SQL Query with the supplied parameters. Returns an ADO.NET data table object that is filled with columns and rows from the query. The return is a complete data table, and should be defined as DataTable type.

The "Imports System.Data" declaration can also be added at the beginning of the script. This avoids adding a prefix to each line in your script, where needed, with “System.Data.DataTable”. From there, use DataTable functions to count or specify table positions. These are standard VB.net concepts you can find in MSDN.

You can use this function in most scripting scenarios. This function allows for flexible post-query filtering, sorting and manipulations of the DataTable. You can also access all of the data returned from the database. This means you can access any data row at any time.

This function is heavy on memory usage. The DataTable object is a complete in-memory copy of all the rows and columns that the SQL returned. You could consume all available memory.

### Syntax

```vbnet
Dim variable As Object = TableLookupEx(QueryName, DataTableName[, paramN])
```

As an alternative, you can use this syntax. This syntax supports named query arguments.

```vbnet
Dim MyDataTable As DataTable = TableLookupEx(QueryName, DataTableName, QueryArgsNamed[, paramN])
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
      <td>QueryName</td>
      <td>Name of the query in FSQUERY to be executed. The tokens are replaced by the system in the SQL for these names:
        <ul>
          <li>[%USER]</li>
          <li>[%LAB]</li>
          <li>[%GROUP]</li>
          <li>[%%CURRENT_DB_NAME]</li>
          <li>[%%SEARCHTABLE] - This token is only available during the Search's PostSearch event.</li>
        </ul>
      </td>
    </tr>
    <tr>
      <td>DataTableName</td>
      <td>Optional. The name of the returned ADO.NET data table. The default is “TableLookup”. To draw an analogy, a DataTable is a sheet in an Excel workbook. It has the columns and rows from the SELECT statement you used to fill that table. The object is in-memory and not directly attached to the database. You can update values on this datatable object and it does not update the database. You can filter or sort the data in this table. The DataTableName argument that is in the TableLookupEx Workflow function is optional.<br>The default value is “TableLookup”. Whatever you have for that value becomes the name of the DataTable object.<br><br>It is not likely that you are going to change this value. Most of the time you are not working with the name of the DataTable. You are working with the DataTable object itself.<br><br>You can do this command multiple times in your script, each time returning a DataTable object. To merge all of these tables into a single DataSet object, you may want each DataTable to have a unique table name. Or, each table can be merged into a single DataTable object.<br><br>See more information about this at the Microsoft help link: http://msdn.microsoft.com/en-us/library/y4b211hz(v=VS.90).aspx</td>
    </tr>
    <tr>
      <td>paramN</td>
      <td>Optional. The parameters can be in the form of an array or a comma-delimited string of the individual arguments of the query.</td>
    </tr>
    <tr>
      <td>QueryArgsNamed</td>
      <td>Enables you to use custom named query arguments in your SQL.<br><br>Numbered query arguments, which are used by the QueryArgs argument. The arguments are wrapped with single quotes in the SQL.<br><br>Named query arguments are not wrapped with quotes automatically. This can be useful when you pass a numeric value to your SQL.<br><br>When you use the QueryArgsNamed argument, add the namespace, Imports System.Collections.Generic, to the top of your script .</td>
    </tr>
  </tbody>
</table>

### Example 1

```vbnet
Dim rv As DataTable = TableLookupEx("MyCustomQueryCode", "TABLELOOKUP",
arg2, arg3...)
```

or

```vbnet
Dim argArray as Object = array(arg1, arg2, arg3)
Dim rv As DataTable = TableLookupEx("MyCustomQueryCode", "TABLELOOKUP",
argArray)
```

### Example 2
This example uses named query arguments.

```vbnet
Dim NamedArgs As New Dictionary(Of String, String) From {{"%TOKEN1", "VALUE1"}, {"%TOKEN2", "VALUE2"}}
NamedArgs.Add("%TOKEN3", "VALUE3")
Dim MyDataTable As DataTable = TableLookupEx("MYQUERYCODE", NamedArgs, arg1, arg2, arg3)
```

## TableReader
You can use this function for Optiva Workflows.

### Purpose
The TableReader function exposes an ADO.Net DbDataReader object.

A DbDataReader is a fast, light weight, and read only mechanism. It is useful when reading large amounts of data, which may otherwise result in an out-of-memory exception. The trade-off is that the DbDataReader is more complex to work with and has to be closed manually when done reading the data.

You should only use this function when the amount of data being queried is large. Only a single row of the returned data from the database is in memory at a time, so the memory usage is minimal. When you are using this function, complicated coding is required for reading values and advancing the record pointer. This function is forward-only. For example, once you have read row 5 and moved the pointer to row 6, you cannot get back to row 5 without re-running the query.

### Syntax

```vbnet
Dim MyReader As DbDataReader = TableLookupEx(QueryCode, QueryArgsNumbered, QueryArgsNamed[, DataReaderTimeOut])
```

The arguments for TableReader are similar to the TableLookupEx in purpose and usage.

For TableReader, the numbered arguments are passed in a List object instead of individually listed. Optionally, you can specify a time-out value, in seconds, for queries that run a long time.

### Recommendations
* Use the Visual Basic Using statement. This allows .Net to automatically close the reader and dispose of it when the statement is finished. See Example 1.
* Add the Imports System.Data.Common namespace to the top of your script when using the TableReader method.
* Add the Imports System.Collections.Generic namespace to the top of your script when using the QueryArgsNamed argument or the QueryArgsNumbered argument. See Example2.

### Example 1

```vbnet
Using TPReader As DbDataReader = co.TableReader("MYQUERYCODE", Nothing, Nothing)
    If TPReader Is Nothing Then
        MessageList("Unable to open DbDataReader for QueryCode = 'MYQUERYCODE'. Missing FsQuery entry?")
    Else
        While TPReader.Read()
            Dim Field1Value As String = TPReader(1).ToString()
        End While
    End If
End Using
```

### Example 2

```vbnet
Imports System.Collections.Generic
Imports System.Data
Imports System.Data.Common
Dim NumberedArgs As New List(Of String) From {"A", "B", "C"}
Dim NamedArgs As New Dictionary(Of String, String) From {{"%TOKEN1", "VALUE1"}, {"%TOKEN2", "VALUE2"}}
NamedArgs.Add("%TOKEN3", "VALUE3")
Using TPReader As DbDataReader = co.TableReader("MYQUERYCODE", NumberedArgs, NamedArgs, 120)
    If TPReader Is Nothing Then
        MessageList("Unable to open DbDataReader for QueryCode = 'MYQUERYCODE'. Missing FsQuery entry?")
    Else
        While TPReader.Read()
            Dim Field1Value As String = TPReader(1).ToString()
        End While
    End If
End Using
```

The system automatically replaces the tokens in the SQL for these names:

*   [%USER]
*   [%LAB]
*   [%GROUP]
*   [%%CURRENT_DB_NAME]
*   [%%SEARCHTABLE] - This token is only available during the Search's PostSearch event.

