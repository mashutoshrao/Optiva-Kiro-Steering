---
inclusion: manual
---
<!------------------------------------------------------------------------------------
   Add rules to this file or a short description and have Kiro refine them for you.
   
   Learn about inclusion modes: https://kiro.dev/docs/steering/#inclusion-modes
-------------------------------------------------------------------------------------> 
## DailyValueRule(Param, DailyValueCode)

Gets the daily value for a given daily value rule and parameter.


### Arguments

<table>
  <thead>
    <tr>
      <th>Parameter</th>
      <th>Description</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <td>Param</td>
      <td>Optional. String containing the name of the parameter whose daily value is to be returned. If this argument is missing, the function returns the daily value for the current parameter.</td>
    </tr>
    <tr>
      <td>DailyValueCode</td>
      <td>Optional. String containing the code of a Daily Value rule whose value is to be returned. If this argument is missing, the function uses the daily value rule of the current analysis object.</td>
    </tr>
    <tr>
      <td>Return Value</td>
      <td>Returns the daily value of the specified parameter for the specified daily value rule.</td>
    </tr>
  </tbody>
</table>

### DailyValueRule example

This example returns the daily value of the current parameter for **DailyValueRule**.

```vba
Dim DV As Double = DailyValueRule( )
```

This example returns the daily value of the current parameter for a daily value rule whose code is **CANADA**.

```vba
Dim DV As Double = DailyValueRule("", "CANADA")
```