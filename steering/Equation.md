---
inclusion: auto
name: Equation
description: Use this file if asked for anything related to equations and doing CRUD operations in parameters.
---
<!------------------------------------------------------------------------------------
   Add rules to this file or a short description and have Kiro refine them for you.
   
   Learn about inclusion modes: https://kiro.dev/docs/steering/#inclusion-modes
-------------------------------------------------------------------------------------> 
## Scripting in equations
Equations are defined in the Equation form.
Equations can be associated with a parameter. When users click Calc in a formula, the Equation Totalparameters are calculated according to the equation.

```
Function evaluate() As Long
End Function
```

Class used for equation script = EquationScript

## Param

You can use this function for Optiva Equations.

#### Purpose

Returns the value of an object’s (formula, item) parameter, using the parameter’s unit of measure. For example, if the value is 16.666 and the unit of measure is MG/100GM, **16.666** is returned.

The `param` functions return a string unless you convert them to a numeric data type before performing numeric operations, such as addition. Otherwise, you get a concatenated string when you add several values.

Equations for Alt UOM conversions should use `ObjProperty` to obtain the parameter values. Do not use `Param` or `ParamItem`.

#### Syntax

```vbscript
Dim variable As String = param(parameter name)
```

### Description

If the parameter is a concentration or percentage, then it uses the concentration or percentage. It does not use the actual amount (i.e., mass) of the parameter in the formula.

Suppose the parameter value is 50 and the unit of measure is %. Then, **50** is returned using this function. If the value is 16.666 and the unit of measure is MG/100GM, **16.666** is returned.

Calculation of this value:
10mg/100gm for the item x 33.333% of the formula = 3.33mg per 100 gm of the formula

<table>
<thead>
<tr>
<th>Items</th>
<th colspan="3">Parameters<br>SODIUM</th>
</tr>
<tr>
<th>Item Code</th>
<th>Qty</th>
<th>Qty % (of the formula)</th>
<th>UNIT (in the item)</th>
<th>EXTENDED (in the formula)</th>
</tr>
</thead>
<tbody>
<tr>
<td>Item 1</td>
<td>1 kg</td>
<td>33.3333</td>
<td>10mg/100gm</td>
<td>3.333mg/100gm</td>
</tr>
<tr>
<td>Item 2</td>
<td>2 kg</td>
<td>66.6667</td>
<td>20mg/100gm</td>
<td>13.333mg/100gm</td>
</tr>
<tr>
<td colspan="4"></td>
<td><strong>16.666mg/100gm</strong></td>
</tr>
</tbody>
</table>

`param("sodium")=16.666`

Use `ParamItem` to return the concentration or percentage amount for an item’s parameter in a formula. This value is in the parameter’s unit of measure. What if the concentration of the parameter for one item is 10MG/100GM and the other item is 20MG/100GM? `ParamItem` returns **10** and **20** respectively.

Use `Tparam` to return the net contribution of a formula’s parameter (i.e., the actual value). This value is based on the percentage or concentration of the parameter and the formula’s mass. If the total concentration of the parameter for all items is 16.666MG/100GM and the total quantity for the items is 3KG (3000GM), `Tparam` returns **499.98**. The calculation is 16.666MG/100GM X 3000GM = 499.98.

Use `TparamItem` to return the net contribution of an item’s parameter (i.e., the actual value). This value is based on the percentage or concentration of the parameter and the formula’s mass. If the parameter value for one item is 3.33MG/100GM of the formula and there is 3KG (3000GM) of the item, the `tparamitem` value is **99.9**. The calculation is 3.33MG/100GM X 3000GM\100GM = 99.9.

---

### Example
The equation returns the value of the parameter **CALCIUM**; the value is divided by two. The parameter is defined initially by an equation at the item level. Then, the parameter is rolled up. Assign the parameter to the lab as Rollup. If the **CALCIUM** parameter is overridden at the formula level, then the `param` function uses the overridden value.

```vbscript
Dim sHalf As String = param("CALCIUM") / 2
Dim iValue As Integer = SetParam("CALCIUMHALF", sHalf)
```

## ParamItem

You can use this function for Optiva Equations.

### Purpose

Returns the concentration or percentage value of an item’s parameter for a formula, using the parameter’s unit of measure. The value is not a function of how the item is used in the formula; but it is an attribute of the item.

Suppose the concentration of the parameter for one item is 10MG/100GM and the other item is 20MG/100GM. ParamItem returns 10 and 20 respectively.

The param functions return a string unless you convert them to a numeric data type before performing numeric operations, such as addition. Otherwise, you get a concatenated string when you add several values.

Equations for Alt UOM conversions should use ObjProperty to obtain the parameter values. Do not use Param or ParamItem.

### Syntax

```vba
Dim variable As String = paramitem(parameter name, <item code> OR <line ID>)
```

### Description

Alternatively, replace the line number with an item code to retrieve the value of the parameter for that item. Use for formula ingredient parameters.

When the parameter is a concentration or percentage, it uses the concentration or percentage. It does not use the actual amount of the parameter. Suppose the concentration of the parameter for one item is 10MG/100GM and the other item is 20MG/100GM. Then ParamItem returns **10** and **20** respectively.

A flowchart titled "paramitem" illustrating how to calculate the contribution of a parameter (Sodium) across multiple items.  
• The flow starts at the top with "paramitem".  
• It branches into two main sections:  
  – **Item 1:**  
    – Quantity: 1 kg  
    – Percentage of formula: 33.3333%  
    – Unit (in the item): 10mg/100gm  
    – Extended (in the formula): 3.333mg/100gm  
  – **Item 2:**  
    – Quantity: 2 kg  
    – Percentage of formula: 66.6667%  
    – Unit (in the item): 20mg/100gm  
    – Extended (in the formula): 13.333mg/100gm  
• Both Item 1 and Item 2 contribute to a total extended value of 16.666mg/100gm.  
• At the bottom, two lines of code are shown:  
  – paramitem("sodium", item 1)=10  
  – paramitem("sodium", item 2)=20&lt;/img&gt;

Use Param to return the current (i.e., original or overwritten) concentration or percentage value for the formula’s parameter. The value is in the parameter’s unit of measure. If the total (extended) contribution of the parameter for all items is 15MG/100GM, then Param returns 15.

Use Tparam to return the net contribution of a formula’s parameter (i.e., the actual value). This value is based on the percentage or concentration of the parameter and the formula’s mass. What if the total concentration of the parameter for all items is 16.666MG/100GM and the total quantity for the items is 3KG (3000GM)? Then Tparam returns 499.98. The calculation is 16.666MG/100GM X 3000GM = 499.98.

Use TparamItem to return the net contribution of an item’s parameter (i.e., the actual value). This value is based on the percentage or concentration of the parameter and the formula’s mass. Suppose the parameter value for one item is 3.33MG/100GM of the formula and there is 3KG (3000GM) of the item. Then, the tparamitem value is 99.9. The calculation is 3.33MG/100GM X 3000GM\100GM = 99.9.

### Examples

The sValue variable assigns the amount of calcium from an item to the parameter that is assigned this equation. The item is on line two of the **Item Lines** tab.

```vbscript
Dim sValue As String = paramitem("CALCIUM", "#02")
```

You can create a parameter that is calculated at the item level and then rolled up. For this scenario, use equation indicator 4=Calculated Parameter. You can use the param function to retrieve the item parameter values in your equation.

```vbscript
Dim sValue as String = param("CALCIUM")
```

## SetParam

You can use this function for Optiva Equations.

#### Purpose

Sets the parameter that is assigned to the equation to a specific value.

### Syntax

```plaintext
Dim variable As Integer = SetParam(name,new value, rollupFlag)
```

### Arguments

<table>
  <tr>
    <th>Part</th>
    <th>Description</th>
  </tr>
  <tr>
    <td>rollupFlag</td>
    <td>Applies to rollup parameters only. It is ignored by non-rollup parameters.<br>
    *   **0** - Keep the current rollup value.<br>
    *   **1** - Override the current rollup value.</td>
  </tr>
</table>

### Description

If the returned parameter is Numeric Info, then the number that is used in the equation is returned. If the parameter type is Unit Activity, Weight % or Volume %, then the number is converted to the appropriate units or percentage.

### Example

The equation returns the value of the parameter CALCIUM. This value is divided by two. The parameter is defined initially by an equation at the item level. Then, the parameter is rolled up. Assign the parameter to the lab as Rollup. If the CALCIUM parameter is overridden at the formula level, then the param function uses the overridden value.

```vbnet
Dim sHalf As String = param("CALCIUM")/2
Dim iReturncode As Integer = SetParam("CALCIUMHALF", sHalf)
```

## Tparam

You can use this function for Optiva Equations.

### Purpose
Returns the net contribution of a formula’s parameter. This value is based on the percentage or concentration of the total parameter and the formula’s mass.

The `param` functions return a string unless you convert them to a numeric data type before performing numeric operations, such as addition. Otherwise, you get a concatenated string when you add several values.

### Syntax
```vbscript
Dim variable As String = tparam(<parameter name> or <quantity type>)
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
      <td>parameter name</td>
      <td>The name of the formula parameter.</td>
    </tr>
    <tr>
      <td>quantity type</td>
      <td>Mass or Volume quantity. Double. The values are expressed without units of measure.<br><ul><li>TotalMass</li><li>Total mass in a formula. This accounts for byproducts and compositions.<br>To find the unit of measure, you must retrieve it using ObjProperty on the UOMCODE for the formula. If the variables are required in another unit of measure, then you must enter a conversion factor into the equation.</li><li>TotalMassNoEq</li><li>Total mass in a formula, not including an ingredient set to scale indicator 3: Equation Adjust to Total. An equation based calculation does account for compositions; it does not account for byproducts.</li><li>TotalMassIngred</li><li>Total mass of the ingredients in a formula. It does not account for byproducts or compositions.</li><li>TotalVol</li><li>Total volume in a formula. It accounts for byproducts and compositions.</li><li>TotalVolNoEq</li><li>Total volume in a formula, not including an ingredient set to scale indicator 3: Equation Adjust to Total. An equation based calculation does account for compositions; it does not account for byproducts.</li><li>TotalVolIngred</li><li>Total volume of the ingredients in a formula in liters. It does not account for byproducts or compositions.</li></ul></td>
    </tr>
  </tbody>
</table>

### Description

This function takes the concentration (MG/100GM) or percentage (WT%) of parameters in the formula. Then, it returns the actual total mass of the parameter in the formula.

<table>
  <thead>
    <tr>
      <th colspan="3">Items</th>
      <th colspan="2">Parameters<br>SODIUM</th>
    </tr>
    <tr>
      <th>Item Code</th>
      <th>Qty</th>
      <th>Qty % (of the formula)</th>
      <th>UNIT (in the item)</th>
      <th>EXTENDED (in the formula)</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <td>Item 1</td>
      <td>1 kg</td>
      <td>33.3333</td>
      <td>10mg/100gm</td>
      <td>3.333mg/100gm</td>
    </tr>
    <tr>
      <td>Item 2</td>
      <td>2 kg</td>
      <td>66.6667</td>
      <td>20mg/100gm</td>
      <td>13.333mg/100gm</td>
    </tr>
    <tr>
      <td colspan="4"></td>
      <td><b>16.666mg/100gm</b></td>
    </tr>
  </tbody>
</table>

```plaintext
tparam("SODIUM") = 16.666mg/100gm x 3 kg x (1000gm/1 kg)= 499.98
```

```plaintext
- tparam("Totalmass")= 3
- tparam("TotalVol")= 3
```

The `tparam` function returns a different value. This value is based upon the line type of the technical parameter in the lab for the tp whose value is being requested. When `tparam` is used to retrieve the value of a rollup parameter, the function returns the total value for the rollup parameter. The value is calculated from the values of the same parameter for the items in the formula.

This is regardless of the actual value that appears for the requested technical parameter on the formula itself. That value may have been manually overridden, or otherwise assigned through another equation.

Use the `Param` function to return the current (i.e. original or overwritten) concentration or percentage value for the formula’s parameter. This value is in the parameter’s unit of measure. Suppose the total (extended) contribution of the parameter for all items is 16.666MG/100GM. Then `Param` returns **16.666**.

Use the `ParamItem` function to return the concentration or percentage amount for an item’s parameter in a formula. This value is in the parameter’s unit of measure. Suppose the concentration of the parameter for one item is 10MG/100GM and the other item is 20MG/100GM. Then `ParamItem` returns **10** and **20** respectively.

Use the `TparamItem` function to return the net contribution of an item’s parameter (i.e., the actual value). This value is based on the percentage or concentration of the parameter and the formula’s mass. What if the parameter value for one item is 3.33MG/100GM of the formula and there is 3KG (3000GM) of the item? Then the `TparamItem` value is **99.9**. The calculation is 3.33MG/100GM x 3000GM\100GM = **99.9**.

---

### Examples

This example returns the total amount of sodium per serving. The SODIUM parameter has a unit of measure of MG/100GM. The `Tparam` for SODIUM has a unit of measure of MG and the SERVING_COUNT has a unit of measure of servings.

```vbnet
Dim count As String = param("SERVING_COUNT")
Dim sValue As String = tparam("SODIUM") / count
MessageList("Sodium per serving = ",sValue," mg/serving")
```

In the next example, the local variable `sBp` is assigned the value of the parameter that is assigned to this equation. The parameter is type numeric min where the value represents the lowest value of all items (i.e., parameter type = Numeric Min).

The If/Then/Else statement instructs the system to show alert messages if the boiling point is too low or if it is acceptable. You can also include return commands inside the conditional statements to alter the value of the parameter.

```vbscript
Dim sBp As String = tparam("")
if (sBp < 50) then
    MessageList("Ingredient boiling point too low.")
else
    MessageList("Boiling point okay.")
end if
```

In this example, the local variable `sLoc` is assigned the value of the parameter SELLINGLOC. This is an enumerated list parameter. If `sLoc` is EUROPE, then the parameter that is assigned to this equation is assigned the number 4. Otherwise, the parameter is assigned the number 10. This can be used to create a constant that varies depending upon another parameter.

```vbscript
Dim sLoc As String = tparam("SELLINGLOC")
if (sLoc = "EUROPE") then
    Dim sValue As String = 4
else
    sValue = 10
end if
```

## TparamItem

You can use this function for Optiva Equations.

### Purpose

Returns the net contribution of an item parameter (i.e., the actual value). The value is based on the percentage or concentration of the parameter and the formula mass. Suppose the parameter value for one item is 3.33MG/100GM of the formula and there is 3KG (3000GM) of the item. Then the `tparamitem` value is **99.9**. The calculation is 3.333MG/100GM X 3000MG\100GM = 99.9.

The `param` functions return a string unless you convert them to a numeric data type before performing numeric operations, such as addition. Otherwise, you get a concatenated string when you add several values.

### Syntax

```plaintext
Dim variable As String = tparamitem(parameter name,<item code> or <line ID>)
```

### Description

Returns the value of a parameter for each item in the formula. The value is based on the percentage or concentration that is converted for the formula’s mass unit of measure.

<table>
  <thead>
    <tr>
      <th>Items</th>
      <th colspan="5">Parameters<br>SODIUM</th>
    </tr>
    <tr>
      <th>Item Code</th>
      <th>Qty</th>
      <th>Qty % (of the formula)</th>
      <th>UNIT (in the item)</th>
      <th>EXTENDED (in the formula)</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <td>Item 1</td>
      <td>&lt;img&gt;1 kg&lt;/img&gt;</td>
      <td>33.3333</td>
      <td>10mg/100gm</td>
      <td>&lt;img&gt;3.333mg/100gm&lt;/img&gt;</td>
    </tr>
    <tr>
      <td>Item 2</td>
      <td>&lt;img&gt;2 kg&lt;/img&gt;</td>
      <td>66.6667</td>
      <td>20mg/100gm</td>
      <td>&lt;img&gt;13.333mg/100gm&lt;/img&gt;</td>
    </tr>
    <tr>
      <td colspan="4"></td>
      <td>16.666mg/100gm</td>
    </tr>
  </tbody>
</table>

`tparamitem("SODIUM", item 1) = 13.333mg/100gm x 3000mg\100 gm = 99.99`
`tparamitem("SODIUM", item 2) = 3.333mg/100gm x 3000mg\100 gm = 99.99`

Use the Param function to return the current (i.e., original or overwritten) concentration or percentage value for the formula’s parameter. This value is in the parameter’s unit of measure. Suppose the total (extended) contribution of the parameter for all items is 16.666MG/100GM. Then Param returns **16.666**.

Use the ParamItem function to return the concentration or percentage amount for an item’s parameter in a formula. This amount is in the parameter’s unit of measure. What if the concentration of the parameter for one item is 10MG/100GM and the other item is 20MG/100GM? Then ParamItem returns **10** and **20** respectively.

Use the Tparam function to return the net contribution of a formula parameter (i.e., the actual value). This value is based on the percentage or concentration of the parameter and the formula’s mass. Suppose the total concentration of the parameter for all items is 16.666MG/100GM and the total quantity for the items is 3KG (3000GM). Then Tparam returns **499.98**. The calculation is 16.666MG/100GM X 3000GM = 499.98.

### Examples

The equation returns the total value and concentration of sodium for item 1.

```plaintext
Dim sValue As String = paramitem("SODIUM", "Item1")
Dim sValue1 As String = tparam("SODIUM")
MessageList("Concentration of sodium for Item1 =", sValue, "MG/100GM")
MessageList("Total value of sodium in the formula from Item1 =", sValue1, "MG")
```

## Valparam

You can use this function for Optiva Equations.

### Purpose

Like Param, ValParam returns the value of an object’s (formula, item) parameter, using the parameter’s unit of measure. In some cases, the modified state of the parameter returned by this function may lead to an incorrect value if called using the Param function on a parameter assigned to the lab as Equation Total. Therefore, ValParam is recommended for use with Equation Total Parameters.

### Syntax

```vbscript
Dim variable As String = valparam(parameter name)
```

## Equation tips for script library
- FsProcFuncSetEvent is used for Copy Method scripts by default.
- If an equation-specific function does not pass Validation, and you are sure it is spelled correctly, thenchange the method to FcProcFuncSetEq. See the example.
- When you alter the Method name, add a notation to the script of what you changed and why. Make itclear that this library can be used only for Equations, or whatever you changed it to.

E.g. The DailyValueRule method can be called from the script library to calculate equations. This nutritionalexample calculates the serving size.
```
Option Strict Off
Imports System
Imports System.Diagnostics

Public Class NUTRITIONCALC
Private co As FcProcFuncSetEq

	Public Sub New(ByRef context As FcProcFuncSetEq)
	co = context
	End Sub

	Function dv_US(byref nutrient as string) As Double 
		dimstpval as string = co.Param(nutrient & "_SERV") 
		dim calcresult as double = -999 
		if co.isblank(stpval) = 0 then 
			dim dvrule as double = co.DailyValueRule(nutrient, "DV_US") 
			calcresult = cdbl(stpval) * cdbl(ssize) * 0.01 
			calcresult = math.round(calcresult,3) 
		end if 
		Return calcresult
	End Function
End Class
```
The DVRule function can call a rule that does not exist or a parameter that is not on the rule. In this case, thefunction returns a 0. It does not return a null, a blank, or an error.
Daily values use the returned number as the divisor. You should verify that the number is greater than 0 beforecontinuing with the math.

## Examples of Equations

This chapter provides scripting examples for equations, copy methods and workflows.

### Equations for SCC and UPC codes

You can use existing equation functions to construct scc, upc and other types of codes. These codes are strings of numbers that typically include a check digit, based on the other digits in the code. The check digit is based on an algorithm that you construct.

Two examples, EQ_SCC and EQ_UPC, are in the seed database. These equations define an informational parameter.

#### SCC

```vb
imports System
imports System.Diagnostics
Imports Microsoft.VisualBasic

Class EquationScript
    Inherits FcProcFuncSetEQ

    Function evaluate() As Long
        Dim pCode as String
        If Context._OBJECTSYMBOL = "ITEM" Then
            pCode = ObjProperty("KeyCode")
        Else
            If Context._OBJECTSYMBOL = "FORMULA" Then
                pCode = ObjProperty("ItemCode")
            End If
        End If
        Dim pLen as Integer = pCode.Length
        If pLen = 5 Then
            Dim coCode as String = tparam("Company Code")
            Dim cLen as Integer = coCode.Length
            If cLen = 6 Then
                Dim sccCode as String = "10" & coCode & pCode
                Dim i, oddSum, evenSum as Integer
                For i = 0 to sccCode.Length - 1
                    If i Mod 2 = 0 Then
                        evenSum = evenSum + CInt(sccCode.Substring(i, 1))
                    Else
oddSum = oddSum + CInt(sccCode.Substring(i, 1))
End If
Next i
oddSum = oddSum * 3
Dim codeSum as Integer = oddSum + evenSum
Dim checkDigit as Integer = Math.Ceiling(codeSum / 10) * 10
checkDigit = checkDigit - codeSum
context.ReturnValue = sccCode & checkDigit
return 1
End If
End If
End Function
End Class
```

#### UPC

```vb
imports System
imports System.Diagnostics
Imports Microsoft.VisualBasic

Class EquationScript
    inherits FcProcFuncSetEQ

    Function evaluate() As Long
        Dim pCode as String
        If Context._OBJECTSYMBOL = "ITEM" Then
            pCode = ObjProperty("KeyCode")
        Else
            If Context._OBJECTSYMBOL = "FORMULA" Then
                pCode = ObjProperty("ItemCode")
            End If
        End If
        Dim pLen as Integer = pCode.Length
        If pLen = 5 Then
            Dim coCode as String = tparam("Company Code")
            Dim cLen as Integer = coCode.Length
            If cLen = 6 Then
                Dim upcCode as String = "10" & coCode & pCode
                Dim i, oddSum, evenSum as Integer
                For i = upcCode.Length - 1 to 0 Step -1
                    If i Mod 2 = 0 Then
                        evenSum = evenSum + CInt(upcCode.Substring(i, 1))
                    Else
                        oddSum = oddSum + CInt(upcCode.Substring(i, 1))
                    End If
                Next i
                oddSum = oddSum * 3
                Dim codeSum as Integer = oddSum + evenSum
                Dim checkDigit as Integer = Math.Ceiling(codeSum/10) * 10
                checkDigit = checkDigit - codeSum
                Dim contextReturnValue As String = upcCode & checkDigit
                return upcCode & checkDigit
            End If
        End If
    End Function
End Class
```
