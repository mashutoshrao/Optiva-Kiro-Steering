---
inclusion: auto
name: Security
description: Use this file if asked related to assign, Remove User, Group, Role securtiy to an object, Assign ACL to an object, Add/ Remove group from a user.
---
<!------------------------------------------------------------------------------------
   Add rules to this file or a short description and have Kiro refine them for you.
   
   Learn about inclusion modes: https://kiro.dev/docs/steering/#inclusion-modes
-------------------------------------------------------------------------------------> 
## SetSecurity

You can use this function for Optiva Workflows and Copy Methods.

### Purpose

Assigns new security values to an object.

### Syntax

```vbnet
Dim variable As Long  = SetSecurity(Symbol,Object,OwnerSecurity,GroupSecurity, RoleSecurity)
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
<td>Symbol</td>
<td>Type of object. Use empty quotation marks to indicate the symbol of the object for the workflow, or the new object for a copy method.</td>
</tr>
<tr>
<td>Object</td>
<td>Object for which to change security. Use empty quotation marks to indicate the object for the workflow, or the new object for a copy method.</td>
</tr>
<tr>
<td>OwnerSecurity</td>
<td>Owner security for the new object. Use empty quotation marks to leave owner security at its current value. Use these security values:</td>
</tr>
<tr>
<td></td>
<td><ul><li>0 - No access</li><li>3 - Read/Copy</li><li>7 - Read/Copy/Write</li><li>15 - Read/Copy/Write/Delete</li></ul></td>
</tr>
<tr>
<td>GroupSecurity</td>
<td>Group security for the new object. Use empty quotation marks to leave group security at its current value. Use the same security values as OwnerSecurity.</td>
</tr>
<tr>
<td>RoleSecurity</td>
<td>Role security for the new object. Use empty quotation marks to leave role se-curity at its current value. Use the same security values as OwnerSecurity.</td>
</tr>
</tbody>
</table>

### Description

SetSecurity assigns new security values to an object. This function provides more or less access to an object. For example, if a formula is in a workflow for approval, set tighter security so that users cannot change the formula.

For a raw material item with a constituent formula, specify the security at the beginning of the script. This allows for updates for the items in the constituent formula.

### Examples

For formulas and specifications, you must include a backslash (\) followed by a version number.

In this example, the security for ITEMCODE is: owner privileges of 7 (read/copy/write), group privileges of 7 (read/copy/write), and role privileges of 1 (read).

```vbscript
Dim oItemkey As Object = ObjProperty("ITEMCODE")
Dim lSecurity As Long = SetSecurity("ITEM",oItemkey,7,7,1)
```

In the next example, only the security for group is changed. Use empty quotation marks ("") to keep the security that is currently defined:

```vbscript
Dim lSecurity As Long = SetSecurity("ITEM",oItemkey, "",7,"")
```

## SetSecurityACL

You can use this function for Optiva Workflows, Copy Methods, and Equations.

### Purpose

Assigns security values to objects and users through an Access Control List.

### Syntax

```vbscript
Dim variable As Integer = SetSecurityAcl([Symbol], [Object] detailCode, typeCode, TypeKey, AccessSecurity)
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
      <td>Symbol</td>
      <td>Optional Type of object symbol supporting ACL: Formula, Item, Project.<br>Omit to indicate the symbol of the current object for a workflow, copy method, or equation.</td>
    </tr>
    <tr>
      <td>Object</td>
      <td>Optional. The code to the object. For formulas, use code\version. Omit to indicate the current object for the workflow, copy method or equation.</td>
    </tr>
    <tr>
      <td>DetailCode</td>
      <td>Object details, such as:<ul><li>INGR - ingredients</li><li>TP0 - rollup parameters</li><li>Doc - attached documents</li></ul>Use '@' as a wildcard to indicate the security settings that apply to any detail code that is not specifically defined. For a description of detail codes, see ObjProperty.</td>
    </tr>
    <tr>
      <td>TypeCode</td>
      <td>USER, GROUP, ROLE. This argument is compatible with refObj in previous versions of Optiva.</td>
    </tr>
    <tr>
      <td>TypeKey</td>
      <td>The code of the User, Group or Role, such as FSI for USER or RL_ADMIN for ROLE. This argument is compatible with refCode in previous versions of Optiva.</td>
    </tr>
    <tr>
      <td>Security</td>
      <td>Standard system security. Use these security values:<ul><li>1 - Delete an existing ACL entry</li><li>0 - No access</li><li>3 - Read/Copy</li><li>7 - Read/Copy/Write</li><li>15 - Read/Copy/Write/Delete</li></ul></td>
    </tr>
  </tbody>
</table>

### Description

This function returns 0 if successful, -1 if there is an error.

### Examples

This example creates an ACL for the FSI user for the current object. The Document detail can be viewed, but not updated or deleted.

```vba
Dim rc As Integer = SetSecurityACL("DOC", "USER", "FSI", 3)
```

This example removes any existing ACL from the current object in context for the RL_ADMIN role and the INGR detail code.

```vba
Dim rc As Integer = SetSecurityACL("INGR", "ROLE", "RL_ADMIN", -1)
```
This example creates an ACL that allows all users read access to the CUSTOM detail for the ITEM object with the key “01001”.

```vbnet
Dim rc As Integer = SetSecurityACL("ITEM", "01001", "CUSTOM", "USER", "@DFLT", 3)
```

## UserGroupAdd, UserGroupRemove

You can use these functions for Optiva Workflows.

### Purpose

Adds a user to an Optiva group or removes a user from an Optiva group.

### Syntax

```vbnet
Dim variable As Integer = UserGroupAdd(UserCode, GroupList)
Dim variable As Integer = UserGroupRemove(UserCode, GroupList)
```

### Argument

<table>
  <tr>
    <th>Part</th>
    <th>Description</th>
  </tr>
  <tr>
    <td>UserCode, LabCode</td>
    <td>User code and lab code. Default to the current values if not supplied.</td>
  </tr>
  <tr>
    <td>GroupList</td>
    <td>A semicolon-separated list of groups to be added or removed.</td>
  </tr>
</table>


## UserRoleAdd, UserRoleRemove

You can use these functions for Optiva Workflows.

### Purpose

Adds a user to an Optiva role or removes a user from an Optiva role.

### Syntax

```vbscript
Dim variable As Integer = UserRoleAdd(UserCode, LabCode, RoleList)
Dim variable As Integer = UserRoleRemove(UserCode, LabCode, RoleList)
```

### Argument

<table>
  <thead>
    <tr>
      <th>Part</th>
      <th>Description</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <td>UserCode, LabCode</td>
      <td>User code and lab code. Default to the current values if not supplied.</td>
    </tr>
    <tr>
      <td>RoleList</td>
      <td>A semicolon-separated list of roles that are to be added or removed.</td>
    </tr>
  </tbody>
</table>
