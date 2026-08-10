---
inclusion: auto
name: Message
description: Use this file if asked to show message to a user on screen
---
<!------------------------------------------------------------------------------------
   Add rules to this file or a short description and have Kiro refine them for you.
   
   Learn about inclusion modes: https://kiro.dev/docs/steering/#inclusion-modes
-------------------------------------------------------------------------------------> 

## MessageList

You can use this function for Optiva Workflows and Equations.

### Purpose

Shows a message to the user in the **Alert Message** dialog.

### Syntax

The **MessageList** has two signatures:

```vbnet
Public Function MessageList(ByVal msg As String) As Long
Public Function MessageList(ByVal msg As String, ByVal ParamArray params() As String) As Long
```

The second method concatenates the text from the 'params' argument. The concatenation does not contain any spaces or separators. That text is appended to the 'msg' value. It then forwards the call to the first method.

### Description

**MessageList** shows a message to the current user. The message can be opened automatically by inserting the statement in a workflow event such as **START**. Or, it can be button-prompted by inserting the statement into a buttoned event. You can also use **MessageList** functions during configuration as a diagnostic tool.

### Example 1

```vbnet
MessageList("The Approve event completed successfully.")
```

### Example 2

This example outputs the list of arguments (i.e., Argument 1Argument 2Argument N) to the user. Note the lack of space between the argument values.

```vbnet
Dim MsgArgs() As String = New String() {"Argument 1", "Argument 2", "Argument N"}
MessageList("The list of arguments: ", MsgArgs)
```

### Workflow examples

In this example, the message is entered into the **REJECT** event. When the user clicks **Reject**, the workflow is cancelled and the user is notified.

```vbnet
MessageList("You have selected the Reject button. This workflow is cancelled.")
Return 9111

'Include information that is assigned to the local variables in the alert message. This example shows the **oStatus** of the object of the workflow in the alert message box.

Dim oStatus As Object = ObjProperty("STATUSIND.STATUS")
MessageList("The status is: ", oStatus)
```

## Applying font styles to messages

You can apply bold, italics, underline, strikethrough, and color to your message text. Use the system pseudo tags shown in the code example. These tags are depicted with curly brackets `{ }` and are translated into the appropriate HTML tags.

For example, if you want the message text to be underlined, you can insert your text in between these curly brackets `{u}{/u}`. The curly brackets are translated to the HTML tags `<u></u>`.

For a complete list of supported system pseudo tags, see the *Infor PLM for Process User Guide*.

```vbnet
Function wf_approve() As Long
    Dim cr As String = Environment.NewLine
    Dim message As String = "{b}Bold (b){/b}" & cr & _
                            "{i}Italics (i){/i}" & cr & _
                            "{u}Underline (u){/u}" & cr & _
                            "{s}Strikethrough (s){/s}" & cr & _
                            "{font color='red'}Font colors (font color='xxx'){/font}" & cr & _
                            "Tab characters and carriage returns are also respected"
    MessageList(message)
    Return 1
End Function
```

You can display message text in a table format. These HTML tags are also supported: `<table>`, `<tr>`, `<td>`, `<th>`, and `<style>`. In the actual message, use curly brackets instead of the angle brackets.

```vbnet
Function wf_approve() As Long
    Dim cr As String = Environment.NewLine
    Dim msg As String = "{style}table, th, td {border: 1px solid black;}{/style}"
    msg &= "{table}{th}{b}Column One{/b}{/th}{tr}{td}Val One{/td}{/tr}{tr}{td}Val" & cr & "two{/td}{/tr}{/table}"
    MessageList(msg)
    Return 1
End Function
```




