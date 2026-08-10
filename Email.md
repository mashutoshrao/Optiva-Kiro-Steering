---
inclusion: auto
name: Email or Notify
description: Use this file if it is asked for to email people
---
<!------------------------------------------------------------------------------------
   Add rules to this file or a short description and have Kiro refine them for you.
   
   Learn about inclusion modes: https://kiro.dev/docs/steering/#inclusion-modes
------------------------------------------------------------------------------------->
## Notify

You can use this function for Optiva Workflows and Copy Methods.

### Purpose

Sends an email message to Optiva users. To attach a file, use SendMail.

You can configure an email group in Outlook mail. See the *Infor PLM for Process Application Configuration Guide*.

### Syntax

```vba
Dim variable As Long = Notify(User/Group/Role, TemplateCode[, Mode, User/Group/RoleIndicator, CustomParameter1, CustomParameter2,...])
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
      <td>User/Group/Role</td>
      <td>An Optiva user, group or role code that is enclosed in quotation marks; or a global argument without quotation marks. 1 addressee per Notify call.</td>
    </tr>
    <tr>
      <td>Template Code</td>
      <td>Using the **Email Template** form, the system administrator can create the message in plain text format or HTML format.</td>
    </tr>
    <tr>
      <td>Mode</td>
      <td>[Optional] 1- External SMTP server (Microsoft, Lotus Notes, UNIX). External notification requires the mail server process on the Application server. This entry is 1 by default; any other values for the mode are ignored.</td>
    </tr>
    <tr>
      <td>User/Group/RoleIndicator</td>
      <td>Optional. Identify if the notify message is sent to:
        <ul>
          <li>0 - User (default, if not indicated)</li>
          <li>1 - Role</li>
          <li>2 - Group</li>
        </ul>
      </td>
    </tr>
    <tr>
      <td>CustomParameter1, CustomParameter2,...</td>
      <td>Custom parameters included in the message. Include local or global variables as parameters in the notify message. In the template code, refer to the variable as [%n], where n is the number of the custom parameter.</td>
    </tr>
  </tbody>
</table>

### Description

Notify sends an email message to a user, all users in a group or all users of a role. The message is sent through an external email package that is SMTP compliant. When you run an action set with the Notify function, emails are only sent to users that do not have **Deactivate Email** selected. You must set up the Optiva profiles to use an external email package in workflow notifications. See the *Infor CloudSuite PLM for Process Application Configuration Guide* for the appropriate values. External notification requires the SMTP process on the Application server.

Notify sends messages to specific users, groups, or roles by their user names, group name, or role code. Messages can be sent to global users too. The messages are sent according to the individual's position in the workflow. For example, a lab director can require notification each time a formula is approved and sent to the director’s user name or a group. Or a message can be sent to the person who began the workflow, using the global variable `_STARTUSER` ; or to the person who clicked the button on the action, using the global variable `_SOURCEUSER`.

### Working with notifications

Using the **Email Template** form, create the text for the email message.

To construct a general message that references Optiva objects specific to a workflow, precede the objects with % and enclose the object in brackets.

<table>
  <thead>
    <tr>
      <th>Example</th>
      <th>Description</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <td>[%1]</td>
      <td>Inserts the first custom parameter into the message; use [%2], [%3 ], and so on for subsequent custom parameters.</td>
    </tr>
    <tr>
      <td>[%ACTIONCODE]</td>
      <td>Inserts the current action code into the message.</td>
    </tr>
    <tr>
      <td>[%SOURCEUSER]</td>
      <td>Inserts the current user who performed the action.</td>
    </tr>
    <tr>
      <td>[%STARTUSER]</td>
      <td>Inserts the user who started the workflow.</td>
    </tr>
    <tr>
      <td>[%TASKUSER]</td>
      <td>Inserts the username of the taskuser who caused the message to be sent.</td>
    </tr>
    <tr>
      <td>[%WIPALLUSERS]</td>
      <td>Inserts all users, members of groups and members of roles that are connected to a workflow.</td>
    </tr>
    <tr>
      <td>[%WIPGROUPS]</td>
      <td>Inserts members of all groups that are connected to a workflow.</td>
    </tr>
    <tr>
      <td>[%WIPID]</td>
      <td>Inserts the workflow ID into the message.</td>
    </tr>
    <tr>
      <td>[%WIPLINEID]</td>
      <td>Inserts the line number for the workflow.</td>
    </tr>
    <tr>
      <td>[%WIPROLES]</td>
      <td>Inserts members of all roles that are connected to a workflow.</td>
    </tr>
    <tr>
      <td>[%WIPUSERS]</td>
      <td>Inserts all users that are connected to a workflow.</td>
    </tr>
  </tbody>
</table>

### Example

In the template code, the appropriate parameters replace the arguments enclosed in brackets when the message is sent.

The text in the template code reads:

Workflow [%WIPID] Action [%ACTIONCODE] has been rejected.
Reason: Formula [%1] belonging to class: [%2] has been rejected by [%TASKUSER].

The workflow script is:

```vbscript
Dim oClass As Object = ObjProperty("CLASS", "", "")
Dim lNotify As Long = Notify("RL_ADMIN", "REJECT", 1, 1, Context._OBJECTKEY, oClass)
```

The notification message reads:

```vbscript
From: <address>
Sent: <date>
To: <address>
Subject: Workflow 2464 Action LAB_MANAGER has been rejected.
Reason: Formula PIZZA_SAUCE\003 belonging to class: CLASS A = Finished Product has been rejected by JOE_SMITH.
``` 