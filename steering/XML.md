---
inclusion: auto
name: Email
description: Use this file if asked to get xml of an object
---
<!------------------------------------------------------------------------------------
   Add rules to this file or a short description and have Kiro refine them for you.
   
   Learn about inclusion modes: https://kiro.dev/docs/steering/#inclusion-modes
-------------------------------------------------------------------------------------> 
## ObjectXml

You can use this function for Optiva Workflows.

### Purpose

Returns XML data for an Optiva object.

### Syntax

```
Dim variable As String =ObjectXml(Symbol, KeyStr, Details, languageCodes)
```

### Arguments

<table>
  <thead>
    <tr>
      <th>Part</th>
      <th>Type</th>
      <th>Description</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <td>Symbol</td>
      <td>String</td>
      <td>Optiva symbol</td>
    </tr>
    <tr>
      <td>KeyStr</td>
      <td>String</td>
      <td>Object key code</td>
    </tr>
    <tr>
      <td>Details</td>
      <td>String</td>
      <td>Object details, such as INGR (ingredients) or PARAM (parameters).</td>
    </tr>
    <tr>
      <td>languageCodes</td>
      <td>String</td>
      <td>Optional. Use this parameter to export Enumerated Value Lists in more than one language. Use a semi-colon delimited list to separate each language code.<br><br>When you specify “EN-US; FR-FR”, the EnumVal List Labels for English and French are added as nodes to the XML. The node names include the language code in the suffix (i.e., _LABEL_EN-US and _LABEL_FR-FR respectively).<br><br>An invalid language code, or a code that does not exist in the database, is added as a node. In this scenario, the English label is retrieved by the system.<br><br>If you omit this parameter, leave it blank, or specify Nothing. Then the default suffix of “_LABEL” is used.<br><br>Inside each XML file, the Enum Val Lists are displayed in the appropriate language. Enumerated Queries only show the one Description that is retrieved in the Query. Typically, it is the default language of the database.</td>
    </tr>
  </tbody>
</table>

### Description

Sets the object key. It can be different than the key that was used previously. The key or the symbol can be derived or based on data other than that which was in context when the workflow was launched. A workflow can be designed to populate the tag parameters with the master formula for the item and the formula symbol.

One or more description fields are added to the enumerated list values, depending on the value given to the languageCodes parameter.

```xml
<statusind>200</statusind>
<statusind_label>Approved</statusind_label>
```

### Examples

This action gets values for the current symbol and object. You must start the script with imports System.xml if XML documents are part of it.

```vbnet
Dim sXn2 As String = ObjectXml(objectSymbol, objectKey, "HEADER;PARAM")
```

This action retrieves the XML for the items in the formula. The XML is retrieved to the constituent level.

```vbnet
Dim sXn2 As String = ObjectXml("FORMULAADJUST", "", "HEADER;INGRMULTIRM")
```

When you specify “EN-US;FR-FR”, the **Enum Val List** Labels for English and French are added as nodes to the XML. The node names include the language code in the suffix (i.e., _LABEL_EN-US and _LABEL_FR-FR respectively).

```vbnet
Dim xml As String = ObjectXml("", "", "HEADER;CUSTOM;TPALL", "EN-US;FR-FR")
MessageList(xml)
```

This example retrieves the current browser language for reporting.

```vbnet
Dim sLanguage as String = Context.SessionInfo.Language
Dim xml As String = ObjectXml("", "", "HEADER;CUSTOM;TPALL", sLanguage)
MessageList(xml)
```

Attached documentation for multiple objects for reports

By default, the OptivaXML.xml file in FsSvcCore contains a section for renaming the destinations of document attachments. This destination is used when you are exporting multiple objects for reports.

For example, if you are exporting attachment data for a specification for a report, then rename the DOC table to FSSPECIFICATIONDOC. Rename the tables to match the objects for the report.

- `<rename>`
  `<table name="FSDOC">FS{SYMBOL}DOC</table>`
  `<table name="FSEMBEDDEDOBJECTS">FS{SYMBOL}EMBEDDEDOBJECTS</table>`
`</rename>`

### Retrieving materials

Use these formats to retrieve materials.

*   INGRMULTI
    Retrieves unexploded materials
*   INGRMULTIRM
    Retrieves materials and exploded raw materials. This is the same as selecting **Explode to Constituents** on the **Item Contributions** form.
*   INGRMULTINM
    Retrieves materials and non-materials. This is the same as selecting **Display Non Materials** on the **Item Contributions** form.
*   INGRMULTINMRM
    Retrieves raw materials and non-materials. This is the same as selecting **Explode to Constituents** and **Display Non Materials** on the **Item Contributions** form.

This script retrieves the materials in a formula BOM. It shows the same information as the **Item Contribution** form.

```vbscript
Dim BOMXML As String = ObjectXML("FORMULAADJUST", _ObjectKey,
"HEADER;INGRMULTI")
```

This script retrieves the exploded materials in a formula BOM. It shows the same information as when you specify **Explode to Constituents** on the **Item Contribution** form.

```vbscript
Dim BOMXML As String = ObjectXML("FORMULAADJUST", _ObjectKey,
"HEADER;INGRMULTIRM")
```

This script retrieves the material and non-materials items in a formula BOM. It shows the same information as when you specify **Display Non-Materials** on the **Item Contribution** form.

```vbscript
Dim BOMXML As String = ObjectXML("FORMULAADJUST", _ObjectKey,
"HEADER;INGRMULTINM")
```

This script retrieves the exploded materials and non-materials items in a formula BOM. It shows the same information as when you specify **Explode to Constituents** and **Display Non-Materials** on the **Item Contribution** form.

```vbscript
Dim BOMXML As String = ObjectXML("FORMULAADJUST", _ObjectKey,
"HEADER;INGRMULTINMRM")
```

You must include a format code and a parameter type to retrieve parameters. For example, `INGRMULTI` retrieves the material ingredients, but you must request `"INGRMULTI.A\PARAM ROLLUP"` to see the rollup parameters.

```vbscript
Dim BOMXML As String = ObjectXml("FORMULAADJUST", _objectKey,
"INGRMULTI.A\PARAM ROLLUP")
```

You can specify `"INGRMULTI.A\PARAM COST"` to retrieve the cost parameters.

```vbscript
Dim BOMXML As String = ObjectXml("FORMULAADJUST", _objectKey,
"INGRMULTI.A\PARAM COST")
```
