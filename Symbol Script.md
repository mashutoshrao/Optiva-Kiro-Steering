---
inclusion: auto
name: Symbol Script/ Event Script
description: Use this file if asked to trigger a set of action during an event like save of object, Update row of item lines or parameter, Calculation of formula.
---
<!------------------------------------------------------------------------------------
   Add rules to this file or a short description and have Kiro refine them for you.
   
   Learn about inclusion modes: https://kiro.dev/docs/steering/#inclusion-modes
-------------------------------------------------------------------------------------> 

# Scripts for Optiva object events

Visual Basic.NET Framework scripts enable you to apply Optiva scripting to events for objects such as:

*   Save
*   Update row
*   Calculate formula

At the trigger event, an Optiva method looks for a script function and the context object. After the script is executed, the method returns any messages or results.

Most of the Optiva objects support scripts. The **Script** tab is enabled for their symbol.

Interactions are between the Optiva server and .NET Framework scripts.

## Scripts are used with events for Objects

<mermaid>
graph LR
    subgraph Object
        A[Event Trigger]
        B[Save]
    end

    subgraph Script in the database for the Object
        C[Function]
        D[Pre-save()]
    end

    A -- "1 Locate a script for the object." --> C
    C -- "2 Locate a function for the event." --> D
    D -- "4 Execute a function." --> A
    A -- "4 Return control to Optiva." --> D
</mermaid>

## Summary of functions

**Note:** Functions are case sensitive. Use lowercase only.

<table>
<thead>
<tr>
<th>Script Function</th>
<th>Object/Notes</th>
</tr>
</thead>
<tbody>
<tr>
<td><code>presave()</code>: Before the standard save, after security check has been passed. The result determines whether the save continues or aborts.</td>
<td>Any objects that can be edited in the <strong>Symbol</strong> form.</td>
</tr>
<tr>
<td><code>postsave()</code>: Immediately after any other post-save processing and does not alter the outcome of the event. For example, can write information about an object to a text file.</td>
<td></td>
</tr>
<tr>
<td><code>postload()</code>: Immediately after loading the header of an object.</td>
<td></td>
</tr>
<tr>
<td><code>&lt;detail&gt;postload()</code>: Immediately after loading the detail of an object. For example, <code>ingrpostload()</code> can be used with formula ingredients.</td>
<td></td>
</tr>
<tr>
<td><code>prerowupdate()</code>: Before ingredient or parameter row update.</td>
<td>Formula</td>
</tr>
<tr>
<td><code>postrowupdate()</code>: After successful update of ingredient or parameter row.</td>
<td>Item</td>
</tr>
<tr>
<td>The detail code must be before the function. For example <code>ingrpre rowupdate</code> and <code>ingrpostrowupdate</code>.</td>
<td>Project</td>
</tr>
<tr>
<td></td>
<td>Specification</td>
</tr>
<tr>
<td>TPPreRowUpdate, TPPostRowUpdate</td>
<td>Specification</td>
</tr>
<tr>
<td>TPAllPreRowUpdate, TPAllPostRowUpdate</td>
<td>Formula, Item, Project</td>
</tr>
<tr>
<td>IngrPreRowUpdate, IngrPostRowUpdate</td>
<td>Formula</td>
</tr>
<tr>
<td>PreSearch, PreSearchExecute, PostSearch</td>
<td>See <a>Search &lt;Object&gt; forms in the Web Client</a> on page 245.</td>
</tr>
<tr>
<td>ObjProperty and ObjPropertySet</td>
<td>Single value</td>
</tr>
<tr>
<td>GetCurrentRowValue</td>
<td>Gives you access to some fields that ObjProperty and ObjPropertySet cannot.</td>
</tr>
<tr>
<td><code>prerescale()</code>: Before the formula ingredients, byproducts or alternate ingredients are re-scaled.</td>
<td>Formula</td>
</tr>
<tr>
<td><code>postrescale()</code>: After the formula ingredients, byproducts or alternate ingredients are re-scaled</td>
<td></td>
</tr>
<tr>
<td><code>calcprecalc()</code>: Before calculations. Can abort the calculation.</td>
<td>Formula</td>
</tr>
<tr>
<td><code>calcpostcalc()</code>: After calculations. Can generate an error message and abort a save operation although the calculation has completed.</td>
<td>Project</td>
</tr>
<tr>
<td></td>
<td>Item</td>
</tr>
</tbody>
</table>

## Script Function

importXmlPreSave(): When you add this function to the script of a symbol, it will be called when importing an object before the object is saved.

This is an example of updating a newly imported formula with a status of 90 with the importXMLPreSave script hook:

```vbnet
Class HookScript
    Inherits FcProcFuncSetEventHook

    Function importXmlPreSave () As Long
        'set status to 90-informational for new formula imported
        ObjPropertySet(90, 1, "STATUSIND.STATUS", "", "")
        return 0
    End Function
End Class
```

## Object/Notes

Any objects that can be edited in the **Symbol** form.

## PreGenerate()

PreGenerate(): Before the Generate event starts. The result determines whether the event continues or aborts.

## PostGenerate()

PostGenerate(): After the label text has been generated.

## Configuring Scripts for Optiva objects

The Optiva scripts are defined in the **Symbol** form for each supported object. The script is empty by default.

If you are not going to implement a script for a symbol, do not include an outline of a script. This helps with system performance.

&lt;img&gt;Screenshot showing the Symbol FORMULA form with sections labeled Validate, Script, and a large area for script input.&lt;/img&gt;

You specify the script in the **Symbol** form.

*   For the ClassHookScript, you must specify an uppercase “S” for Script.
*   For the functions, such as presave and postsave, use lowercase only.

To verify the syntax of the script, click **Validate**.

See the *Infor CloudSuite PLM for Process Application Configuration Guide* for more information about symbols.

The person responsible for writing the script should be an experienced .NET Framework scripting programmer. Work with an implementation consultant for guidance in preparing to write and integrate scripts with Optiva.

## Templates

Include a template in the **Symbol** form only if you are using a script for the symbol.

### Generic template NEW

```vbnet
Option Strict Off
imports System
imports System.Data
imports System.Diagnostics

Class HookScript
    Inherits FcProcFuncSetEventHook
    Function presave() As Long
        End Function
    Function postsave() As Long
        End Function
End Class
```

## Examples of presave and prerowupdate

Some examples of functions are provided in this section.

### presave

This example uses `presave` to check to see if a manufactured item exists for an approved formula; if one does not exist, the save operation is canceled.

```vbnet
Option Strict Off
imports System
imports System.Data
imports System.Diagnostics

Class HookScript
    Inherits FcProcFuncSetEventHook
    Function presave() As Long
        Dim status As Long
        Dim item As String
        status = CLng(ObjProperty("STATUSIND"))
        item = CStr(ObjProperty("ITEMCODE"))
        if status > 100 and item = "" Then
            Dim lmsg as Long = MessageList("A Mfg Item is required. The save

---


## Page 73

# Scripts for Optiva object events

failed.")
Return -1
End If
Return 1
End Function
End Class


# prerowupdate

Use prerowupdate and postrowupdate with objProperty and objPropertySet. You can make changes as a result of a row being changed.

This prerowupdate script tests the value of the **QUANTITY** field. Then, it uses objPropertySet to update the **SECTION** column for the row that was changed. Both changes to the row are saved when the object is saved. The ingr detail code precedes the prerowupdate function.

```vb
Option Strict Off
Imports System
Imports System.Data
Imports System.Diagnostics
Imports Microsoft.VisualBasic

Class HookScript
    Inherits FcProcFuncSetEventHook
    Function ingrprerowupdate() As Long
        Dim qty As Double = FcType.FixDouble(objProperty("Quantity.Ingr"))
        If qty > 100 Then
            objPropertySet("BIG", 0, "SECTIONNAME.INGR")
        Else
            objPropertySet("Small", 0, "SECTIONNAME.INGR")
        End If
        Return 1
    End Function
end Class
```

In this example, the quantity in row 6 changed. Consequently, a change is made to the **Section** column before the object is saved.

## Item Lines

<table>
<thead>
<tr>
<th>L#</th>
<th>Item Code</th>
<th>Qty</th>
<th>UM</th>
<th>Descripion</th>
<th>Qty%</th>
<th>Section</th>
</tr>
</thead>
<tbody>
<tr>
<td>1</td>
<td>11250</td>
<td>250.0</td>
<td>KG</td>
<td>Lettuce butterhead</td>
<td>47.4293</td>
<td></td>
</tr>
<tr>
<td>2</td>
<td>01032</td>
<td>50.0</td>
<td>KG</td>
<td>Cheese, parmesan</td>
<td>9.4859</td>
<td></td>
</tr>
<tr>
<td>3</td>
<td>05064</td>
<td>150.0</td>
<td>KG</td>
<td>Chicken, broilers</td>
<td>28.4576</td>
<td></td>
</tr>
<tr>
<td>4</td>
<td>18243</td>
<td>75.0</td>
<td>KG</td>
<td>Croutons</td>
<td>14.2288</td>
<td></td>
</tr>
<tr>
<td>5</td>
<td>04701</td>
<td>1.0</td>
<td>KG</td>
<td>Hydrocan</td>
<td>0.1897</td>
<td></td>
</tr>
<tr>
<td>6</td>
<td>11587</td>
<td>1.1</td>
<td>KG</td>
<td>Vinespinach</td>
<td>0.2087</td>
<td>Small</td>
</tr>
</tbody>
</table>

The postrowupdate function updates related information outside the row that depends on the row update being successful. You cannot change the value in a row using postrowupdate.

DataRowCurrent and DataRowProposed are not available in the secured scripting environment. They have been replaced by GetCurrentRowValue, GetProposedRowValue, and SetProposedRowValue.

You can use GetCurrentRowValue, GetProposedRowValue, and SetProposedRowValue to access a single value on the row, not the entire row. This is quicker than using ObjProperty and ObjPropertySet for this type of work.

```vb
Option Strict Off
Imports System
Imports System.Data
Imports System.Diagnostics
Imports Microsoft.VisualBasic

Class HookScript
    Inherits FcProcFuncSetEventHook

    Function ingrprerowupdate() As Long
        Dim currVal as Object = Context.GetCurrentRowValue("UOM_CODE")
        MessageList("currVal = " & currVal.ToString())
        Dim propVal as Object = Context.GetProposedRowValue("UOM_CODE")
        MessageList("propVal = " & propVal.ToString())
        Context.SetProposedRowValue("UOM_CODE", "LB") ' This change will be honored
    End Function
End Class
```

Additionally, you can return a negative value from your PreRowUpdate script hook. The RowUpdate process is halted by the system.

```vb
Function ingrprerowupdate() As Long
...
MessageList("Error processing Ingredient Row Update script.")
Return -1
End Function
```
