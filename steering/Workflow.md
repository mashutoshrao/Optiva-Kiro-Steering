---
inclusion: auto
name: Security
description: Use this file if asked to start a action-set/ workflow from another task, Skip Task, Reassingment of task
---
<!------------------------------------------------------------------------------------
   Add rules to this file or a short description and have Kiro refine them for you.
   
   Learn about inclusion modes: https://kiro.dev/docs/steering/#inclusion-modes
-------------------------------------------------------------------------------------> 

## StartWorkflow

You can use this function for Optiva Workflows.

### Purpose

Launch a workflow from within another workflow. The StartWorkflow function saves all data from the first action set before launching a new action set.

StartWorkflow does not allow input parameters to be changed by the user. This function creates a pending task for a wizard, but does not launch the wizard for the user to see.

To launch a wizard workflow from another script, use LaunchWorkflow. See [LaunchWorkflow](#) on page 101.

### Syntax

```plaintext
StartWorkflow(ActionSetCode [,Symbol, Object] [,StartDate] [,InputParameter1, InputParameter2,...])
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
<td>ActionSetCode</td>
<td>Code for the workflow action set.</td>
</tr>
<tr>
<td>Symbol</td>
<td>Optional. The workflow runs automatically on the current symbol.<br>You can also use empty quotation marks to indicate the current symbol or enter the symbol name.</td>
</tr>
<tr>
<td>Object</td>
<td>Optional. The workflow runs automatically on the current object.<br>You can also use empty quotation marks to indicate the current object or enter the object code.</td>
</tr>
<tr>
<td>StartDate</td>
<td>Optional. The start date for the workflow. Must be a Date/Time type. If you pass in a start date, then the workflow starts; the durations for the action set step are used to calculate the due dates of the steps in the Pending Tasks form. The due dates include durations of both the current step and any associated prior steps.</td>
</tr>
<tr>
<td>Input Parameters</td>
<td>Optional. List the input parameter arguments in the same sequence as they are in the **Input** tab of the **Action Set** form. Input Array variables can be used.<br>Use the WipParamGet function in the Workflow Action Event of the StartWorkflow to retrieve the values.<br>StartWorkflow is a silent launch.<br>See [LaunchWorkflow](#) on page 101 for prompted inputs.</td>
</tr>
</tbody>
</table>

### Examples

This example launches the ITEM_APPROVAL workflow from within the FORMULA_APPROVAL workflow. The input parameter is the user who started the workflow.

```vbscript
'Iterate on items.
'Check status of each item. If status values is below a specified value,
'start workflow to approve the item.
startWorkflow("ITEM_APPROVAL", "ITEM", "00031", Context._STARTUSER)
```

The next example retrieves the start date of the current project and turns it into a Date/Time type. The due date is calculated for each step of the SHIP_APPROVAL in the **Pending Tasks** form. To do this, it adds the step duration values from the **Action Set** form, plus the start date value from this workflow.

```vbscript
Dim startDate as DateTime
startDate = CType (ObjProperty("EFF_START_DATE"), DateTime)
startWorkflow("SHIP_APPROVAL", "PROJECT", "00031", startDate)
```

This example launches an Action Set.

```vbscript
Dim Startdate As Date = Now()
startWorkflow("INPUT PARAM SEQUENCE", "FORMULA", _OBJECTKEY, StartDate, "abc1", "abc2", "abc3")
```

The name of the workflow is INPUT PARAM SEQUENCE. The action set is comprised of three input codes: PARAM1, PARAM2, PARAM3. The parameter values are abc1, abc2, and abc3. These values are displayed in the Input tab of the **Workflows in Progress** form. The parameter values can also be defined as variables and inserted in place of abc1, abc2, and abc3.)

Use WipParamGet to retrieve the values for use in the Input Param Sequence workflow.


## Wskip

You can use this function for Optiva Workflows.

### Purpose

Skips a step in a workflow.

### Syntax

```vbscript
Dim variable As Integer = Wskip(Action)
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
<td>Action</td>
<td>The name of the action that is skipped.</td>
</tr>
</tbody>
</table>

### Description

Wskip skips an action in a workflow. Use it in a conditional statement. For example, a packaging formula in a workflow for approval requires the user to pick a new package. The contents are too dense for the customary packaging. The workflow is written so that if the contents are below a specified density, the action is skipped.

If an action contains Wskip, the **Next Line** field on the **Action Set** form must list more than one next line. For example, action 2 contains Wskip for action 3. Action 4 depends on the completion of action 3, but if action 3 is skipped, action 4 still must occur. Both action 3 and 4 must be listed in Next Line for action 2.

### Examples

In this example, the if statement checks the status of items for a formula. If all of the items are not on the blacklist, the TECH_DIR action is skipped. If any of the items are on the blacklist, giving the blacklist parameter a positive (1) value, then the TECH_DIR action is not skipped.

```vbnet
Dim oBlklist As Object = ObjProperty("VALUE.TP0","" , "" ,"BLACKLIST", 2)
if (oBlklist = 0) then
    Dim iSkip As Integer = Wskip("TECH_DIR")
end if
```

### Skipping multiple actions

If the status of the workflow object is **400**, then the next two actions in the workflow are skipped. In this example, NEW_STATUS and APPROVE_STATUS, are skipped.

```vbnet
Dim oStatus As Object = ObjProperty("STATUSIND.STATUS")
if (oStatus = 400) then
    Dim iSkip As Integer = WSkip("NEW_STATUS")
    iSkip = WSkip("APPROVE_STATUS")
end if
Return 111
```

In the **Action Set** form, the **Next Line** column for this action includes the action lines for NEW_STATUS, APPROVE_STATUS and the next action. For example, if the above script is in the second action and NEW_STATUS is action three and APPROVE_STATUS is action four, then **Next Line** is **3;4;5**.


## LaunchWorkflow

You can use this function for Optiva Workflows.

### Purpose

Launch a workflow from within another workflow and include input parameters. The `LaunchWorkflow` function saves all data from the first action set before launching a new action set. This is related to `StartWorkflow`.

`LaunchWorkflow` does not include a start date. It provides default input parameters that can be changed by the user when the workflow is run.

### Syntax

```
LaunchWorkflow(ActionSetCode [,Symbol, Object][,Input Parameter1, Input Parameter2...])
```

<table>
  <thead>
    <tr>
      <th>Part</th>
      <th>Description</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <td>ActionSetCode</td>
      <td>Code for the workflow action set.</td>
    </tr>
    <tr>
      <td>Symbol</td>
      <td>Optional. The workflow automatically runs on the current symbol.<br>You can also use empty quotation marks to indicate the current symbol or enter the symbol name.</td>
    </tr>
    <tr>
      <td>Object</td>
      <td>Optional. The workflow automatically runs on the current object.<br>You can also use empty quotation marks to indicate the current object or enter the object code.</td>
    </tr>
    <tr>
      <td>Input Parameters</td>
      <td>Optional. Input parameters from the action set are used as defaults. The **Launch Workflow** dialog opens where the default parameters can be changed as necessary by the user before starting the workflow. Input parameters can also be read-only.<br>Along with information such as the system file name and other attributes, the URL must not exceed the limit of 2000 characters.</td>
    </tr>
  </tbody>
</table>

### Example

This example launches the `ITEM_APPROVAL` workflow from within the `FORMULA_APPROVAL` workflow. The input parameters are values for the `ITEM_APPROVAL` workflow. These parameters are VENDOR and DENSITY.

```
'Check status of an item. If status value is below a specified value,
'start workflow to approve the item.
launchWorkflow("ITEM_APPROVAL", "ITEM", "00031", "ACME", 1.75)
```

After the workflow starts, a dialog displays the values of the input parameters. For this example, the values are for VENDOR and DENSITY. Users can specify the values manually.


## Reassign

You can use this function for Optiva Workflows.

### Purpose

Sends the workflow task to an alternate user.

### Syntax

```vbscript
Dim variable As Integer = Reassign(User, Group, Role)
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
      <td>User</td>
      <td>New user to be assigned the current task.</td>
    </tr>
    <tr>
      <td>Group</td>
      <td>New group to be assigned the current task.</td>
    </tr>
    <tr>
      <td>Role</td>
      <td>New role to be assigned the current task.</td>
    </tr>
  </tbody>
</table>

### Description

Use this function to change task assignments from one user, group or role to another user, group or role. Assign this function to the REASSIGN event and enable users to select the new assignment. Or assign this function to a different event and hard-code the reassignment.

### Examples

In this example, the script is included in the REASSIGN event. When users click **Reassign**, they are prompted to select a new user, group or role to assign the task to.

```vbscript
MessageList("Assign this task to a user, group, or role.")
Dim iReassign As Integer = Reassign("", "", "")
Return 1
```

In the next example, the script is included in the RETURN event. When a user clicks **Return**, the task is reassigned to the FORMULATION group.

```vbscript
MessageList("This task is reassigned to the formulas group.")
Dim iReassign As Integer = Reassign("", "FORMULATION", "")
Return 111
```



