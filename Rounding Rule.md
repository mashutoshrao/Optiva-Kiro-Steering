---
inclusion: auto
name: Rounding Rule
description: Use this file if asked to apply any rounding rules present in Optiva.
---
<!------------------------------------------------------------------------------------
   Add rules to this file or a short description and have Kiro refine them for you.
   
   Learn about inclusion modes: https://kiro.dev/docs/steering/#inclusion-modes
-------------------------------------------------------------------------------------> 
## RoundingRule

You can use this function for Optiva Workflows and Equations.

### Purpose

Round a parameter value to a rule that is defined in the **Rounding Rule** form.

See the *Infor PLM for Process Analysis Administration Guide*.

### Syntax

```vbscript
Dim variable As Long = RoundingRule(Value, Parameter, Param Value Rounding Code)
```

### Example

This example shows rounding the value of the CALCIUM parameter using the parameter value rounding code of DCL1.

```vbscript
Dim CALCIUMVAL as Object = ObjProperty("VALUE.TPALL", "", "", "CA", 2)

MessageList("Calcium starting value = ", CALCIUMVAL)

Dim CALCIUMROUNDED as Long = RoundingRule(CALCIUMVAL, "CA", "DCL1")

MessageList("Calcium is rounded to = ", CALCIUMROUNDED)
```

## RoundingRuleEx

You can use this function for Optiva Workflows and Equations.

### Purpose

Round a parameter value to a rule that is defined in the **Rounding Rule** form.

See the *Infor PLM for Process Analysis Administration Guide*.

### Example

This example shows rounding the value of the TEST parameter using the parameter value rounding code of PVTEST.

```vbscript
dim dProt as double = 26.5
dim dRoundProt as RoundingRuleResult = RoundingRuleEx(dProt, "PROTEIN", "PVTEST")
messagelist("Passing test:")
if dRoundProt.RoundingOperationResult > 0 then
    messagelist("Result is ", dRoundProt.RoundedValue)
else
    messagelist("Error rounding, leave value as ", dRoundProt.OriginalValue)
end if
messagelist("Bad Parameter test")
dRoundProt = RoundingRuleEx(dProt, "PROTEINX", "PVTEST")
if dRoundProt.RoundingOperationResult > 0 then
    messagelist("Result is ", dRoundProt.RoundedValue)
else
    messagelist("Error rounding, leave value as ", dRoundProt.OriginalValue)
end if
messagelist("Bad Rule test")
dim dRoundProtBad2 as RoundingRuleResult = RoundingRuleEx(dProt, "PROTEIN", "PWTEST")
if dRoundProt.RoundingOperationResult > 0 then
    messagelist("Result is ", dRoundProt.RoundedValue)
else
    messagelist("Error rounding, leave value as ", dRoundProt.OriginalValue)
end if
Return 111
```