---
inclusion: auto
name: Multi langugage
description: Use this file if asked to apply anything related to multi language
---
<!------------------------------------------------------------------------------------
   Add rules to this file or a short description and have Kiro refine them for you.
   
   Learn about inclusion modes: https://kiro.dev/docs/steering/#inclusion-modes
-------------------------------------------------------------------------------------> 
## SetMultiLangDescription
You can use this function for Optiva Workflows, and Copy Methods.

### Purpose
The purpose is to have a scripting function which can allow workflow users to update multi lang description(s) for header and details fields.

### Syntax
```vbnet
SetMultiLangDescription(sourceSymbol, sourceSymbolCode, languageCode, sourceField, sourceFieldValue)
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
<td>sourceSymbol</td>
<td>The object type of symbol.<br/>Use empty quotation marks to indicate the current symbol for the workflow<br/>or the new object for the copy method.</td>
</tr>
    <tr>
      <td>sourceSymbolCode</td>
      <td>The object code for which context attributes can be added.<br/>Use empty quotation marks to indicate the current object for the workflow<br/>or the new object for the copy method.</td>
    </tr>
    <tr>
      <td>languageCode</td>
      <td>The language code for which the multi-language value of the field is updated.</td>
    </tr>
    <tr>
      <td>sourceField</td>
      <td>The field for which the multi-language value is updated.</td>
    </tr>
    <tr>
      <td>sourceFieldValue</td>
      <td>The Multi-Language value.</td>
    </tr>
  </tbody>
</table>

### Description

SetMultiLangDescription is used to update multi language descriptions of a field in header.

This is an example to set the description for the current symbol and current object for a language of FR-FR for **Description** with FR-FR description updated using SetMultiLangDescription.

```javascript
SetMultiLangDescription("", "", "FR-FR", "DESCRIPTION", "FR-FR description updated using SetMultiLangDescription").
```

### Syntax

```javascript
SetMultiLangDescription(sourceSymbol, sourceSymbolCode, sourceFieldsandValues).
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
      <td>sourceSymbol</td>
      <td>The object type of symbol.<br/>Use empty quotation marks to indicate the current symbol for the workflow<br/>or the new object for the copy method.</td>
    </tr>
    <tr>
      <td>sourceSymbolCode</td>
      <td>The object code for which context attributes can be added.<br/>Use empty quotation marks to indicate the current object for the workflow or<br/>the new object for the copy method.</td>
    </tr>
    <tr>
      <td>sourceFieldsandValues</td>
      <td>The source fields and values are an array of type Optiva.ScriptProxy.MultiLangData.<br/>MultiLangData contains languageCode, multiLangField, description.</td>
    </tr>
  </tbody>
</table>

### Description

SetMultiLangDescription is used to update multi language descriptions for multiple fields and languages in header.

This is an example to set the description for the current symbol and current object for a language of EN-US and FR-FR for Description with respective descriptions.

```vbnet
dim sourceFieldsandValues(1) As Optiva.ScriptProxy.MultiLangData
    sourceFieldsandValues(0) = new Optiva.ScriptProxy.MultiLangData()
    sourceFieldsandValues(0).languageCode = "EN-US"
    sourceFieldsandValues(0).multiLangField = "DESCRIPTION"
    sourceFieldsandValues(0).description = "EN-US description updated using SetMultiLangDescription"
    sourceFieldsandValues(1) = new Optiva.ScriptProxy.MultiLangData()
    sourceFieldsandValues(1).languageCode = "FR-FR"
    sourceFieldsandValues(1).multiLangField = "DESCRIPTION"
    sourceFieldsandValues(1).description = "FR-FR description updated using SetMultiLangDescription"
SetMultiLangDescription("", "", sourceFieldsandValues)
```

### Syntax

```vbnet
SetMultiLangDescription(sourceSymbol,sourceSymbolCode, dtlCode,languageCode,sourceLineId, sourceField, sourceFieldValue).
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
<td>sourceSymbol</td>
<td>The object type of symbol.<br>Use empty quotation marks to indicate the current symbol for the workflow or the new object for the copy method.</td>
</tr>
<tr>
<td>sourceSymbolCode</td>
<td>The object code for which context attributes can be added.<br>Use empty quotation marks to indicate the current object for the workflow or the new object for the copy method.</td>
</tr>
<tr>
<td>dtlCode</td>
<td>The Detail code of the tab in which multi language column is displayed.</td>
</tr>
<tr>
<td>languageCode</td>
<td>The language code for which the multi-language value of the given field is updated.</td>
</tr>
<tr>
<td>sourceLineId</td>
<td>The LineId of the row.</td>
</tr>
<tr>
<td>sourceField</td>
<td>The field for which the multi-language value is updated.</td>
</tr>
<tr>
<td>sourceFieldValue</td>
<td>The Multi-Language value.</td>
</tr>
</tbody>
</table>

### Description

SetMultiLangDescription is used to update multi language descriptions for a field on the **Details** tab.

This is an example to set the description for ACTIONSET with set code SFTEST1 in **Steps** tab.

```vbscript
dim sourceSymbol as string = "ACTIONSET"
dim sourceSymbolCode as string = "SFTEST1"
dim languageCode as string = "EN-US"
dim sourceField as string = "DESCRIPTION"
dim sourceFieldValue as string = "step 1 description"
dim dtlcode as string = "LINE"
dim sourceLineId as int32 = 1

SetMultiLangDescription(sourceSymbol, sourceSymbolCode, dtlCode, languageCode, sourceLineId, sourceField, sourceFieldValue).
```

### Syntax

```vbscript
SetMultiLangDescription(sourceSymbol, sourceSymbolCode, dtlCode, dtllangData).
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
<td>sourceSymbol</td>
<td>The object type of symbol.<br><br>Use empty quotation marks to indicate the current symbol for the workflow or the new object for the copy method.</td>
</tr>
<tr>
<td>sourceSymbolCode</td>
<td>The object code for which context attributes can be added.<br><br>Use empty quotation marks to indicate the current object for the workflow or the new object for the copy method.</td>
</tr>
<tr>
<td>dtlCode</td>
<td>The Detail code of the tab in which multi language column is displayed.</td>
</tr>
<tr>
<td>dtllangData</td>
<td>dtllangData is a type of dictionary.<br><br>The Dictionary(Of int32, Optiva.ScriptProxy.MultiLangData()),<br>MultiLangData contains languageCode, multiLangField , description .</td>
</tr>
</tbody>
</table>

### Description

SetMultiLangDescription is used to update multi language descriptions for multiple fields in **Details** tab.

### Examples

This is an example to set the description for **ACTIONSET** with set code **SFTEST1** in **Steps** tab.

```vbscript
dim sourceSymbol as string = "ACTIONSET"
dim sourceSymbolCode as string = "SFTEST1"

dim dtlLangData As Dictionary(Of int32, Optiva.ScriptProxy.MultiLangData()) = New Dictionary(Of int32, Optiva.ScriptProxy.MultiLangData())
dim sourceField as string = "DESCRIPTION"

dim sourceFieldsandValues(1) As Optiva.ScriptProxy.MultiLangData.

sourceFieldsandValues(0) = new Optiva.ScriptProxy.MultiLangData()
sourceFieldsandValues(0).languageCode = "EN-US"
sourceFieldsandValues(0).multiLangField = "DESCRIPTION"
sourceFieldsandValues(0).description = "description in EN-US for row 2"

sourceFieldsandValues(1) = new Optiva.ScriptProxy.MultiLangData()
sourceFieldsandValues(1).languageCode = "FR-FR"
sourceFieldsandValues(1).multiLangField = "DESCRIPTION"
sourceFieldsandValues(1).description = "description in FR-FR for row 2"

dtlLangData.Add(2,sourceFieldsandValues)

dim sourceFieldsandValuesRow3(1) As Optiva.ScriptProxy.MultiLangData

sourceFieldsandValuesRow3(0) = new Optiva.ScriptProxy.MultiLangData()
sourceFieldsandValuesRow3(0).languageCode = "EN-US"
sourceFieldsandValuesRow3(0).multiLangField = "DESCRIPTION"
sourceFieldsandValuesRow3(0).description = "description in EN-US for row 3"

sourceFieldsandValuesRow3(1) = new Optiva.ScriptProxy.MultiLangData()
sourceFieldsandValuesRow3(1).languageCode = "FR-FR"
sourceFieldsandValuesRow3(1).multiLangField = "DESCRIPTION"
sourceFieldsandValuesRow3(1).description = "description in FR-FR for row 3"

dtlLangData.Add(3,sourceFieldsandValuesRow3)

SetMultiLangDescription(sourceSymbol, sourceSymbolCode, dtlCode, dtlLangData)
```


## RemoveMultiLanguageValues

You can use this function for Optiva Workflows, and Copy Methods.

### Syntax

```vbscript
RemoveMultiLanguageValues(objectSymbol As String, objectKey As String) As Integer
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
      <td>objectSymbol</td>
      <td>The object type of symbol.<br>Use empty quotation marks to indicate the current symbol for the workflow or the new object for the copy method.</td>
    </tr>
    <tr>
      <td>objectkey</td>
      <td>The object key for which context attributes can be added.<br>Use empty quotation marks to indicate the current object for the workflow or the new object for the copy method.</td>
    </tr>
  </tbody>
</table>

### Description

RemoveMultiLanguageValues is used to remove language values of a field.

This is an example to remove the multi-language values for all supported fields.

```vbscript
Dim rc As Long = RemoveMultiLanguageValues("", "")
```

### Syntax

```vbscript
RemoveMultiLanguageValues(objectSymbol As String, objectKey As String, fieldNames As List(Of String)) As Integer.
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
<td>objectSymbol</td>
<td>The object type of symbol.<br>Use empty quotation marks to indicate the current symbol for the workflow or the new object for the copy method.</td>
</tr>
<tr>
<td>objectkey</td>
<td>The object key for which context attributes can be added.<br>Use empty quotation marks to indicate the current object for the workflow or the new object for the copy method.</td>
</tr>
<tr>
<td>fieldnames</td>
<td>The name of the field for which the multi-language value of the field is updated.</td>
</tr>
</tbody>
</table>

### Description

RemoveMultiLanguageValues is used to remove the **Description** and **Comment** values in all supported languages.

This is an example to remove the **Description** and **Comment** values for all supported fields.

```csharp
RemoveMultiLanguageValues("", "", New List(Of String) From {"DESCRIPTION", "COMMENT"})
```

### Syntax

```csharp
RemoveMultiLanguageValues(objectSymbol As String, objectKey As String, fieldNames As List(Of String)) As Integer.
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
<td>objectSymbol</td>
<td>The object type of symbol.<br>Use empty quotation marks to indicate the current symbol for the workflow or the new object for the copy method.</td>
</tr>
<tr>
<td>objectkey</td>
<td>The object key for which context attributes can be added.<br>Use empty quotation marks to indicate the current object for the workflow or the new object for the copy method.</td>
</tr>
<tr>
<td>fieldnames</td>
<td>The name of the field in which the multi-language value of the field is updated.</td>
</tr>
<tr>
<td>languageCode</td>
<td>The language code in which the multi-language value of the field is updated.</td>
</tr>
</tbody>
</table>

### Description

RemoveMultiLanguageValues is used to remove the **Description** specified in multi-language.

This is an example to remove the **Description** specified in multi-language.

```vbnet
RemoveMultiLanguageValues("", "", New List(Of String) From {"DESCRIPTION"}, New List(Of String) From {"EL-GR"})
```

```vbnet
RemoveMultiLanguageValues("", "", New List(Of String), New List(Of String) From {"FR-FR"})
```
