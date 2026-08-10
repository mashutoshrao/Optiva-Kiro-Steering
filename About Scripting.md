---
inclusion: always
---
<!------------------------------------------------------------------------------------
   Add rules to this file or a short description and have Kiro refine them for you.
   
   Learn about inclusion modes: https://kiro.dev/docs/steering/#inclusion-modes
-------------------------------------------------------------------------------------> 
# Scripting
Each company has a different set of business processes, or business rules, that dictate workflows and calculations. These rules are different from one organization to the next. It is not feasible to create every scenario in Optiva. Instead of imposing a set of predefined operations, each organization can implement its own rules. You can use the Optiva scripting language along with Visual Basic .NET.

###  Optiva Scripting for Workflows, Copy Methods, Equations
Optiva scripting is a programming language that is used in the Optiva software suite. These scripts enable
each company to customize how:

- Parameters and formulas are calculated
- New objects are created
- Attributes are applied
- Workflows are carried out

This example shows the scripting language that is used in Optiva Workflow Management. This workflow first
checks the status of a formula. If the status of the formula is **400** , it is ready for approval. Then, the formula
is approved and security is applied so that users cannot edit it.

```
'Get formula status
Dim oStatus As Object = ObjProperty("STATUSIND.STATUS")
'If status is ready for approval, lock down formula.
if (oStatus = 400) then 
MessageList("This formula has been approved.") 
SetSecurity("","",3,3,3) 
return 111
' If not, cancel workflow.
else 
MessageList("Formula is not ready for approval. Workflow is cancelled.") 
return 9111
end if
```
You can use scripts to decide what constitutes an approved formula. In this case, an approved formula hasa status value of 400. You decide what an approved formula requires; in this case, a formula cannot be edited.
Optiva Scripting integrates with Visual Basic .NET, which give you ability to extend the scripting functionsaccordingly to meet your business needs.
In this example, the Visual Basic .NET concatenation and ctype functionality are incorporated into the Optivascripting (ObjProperty). The script includes the sender's full name instead of a user code.
```
dim ofirst, olast as object
dim sname as string
ofirst = ObjProperty("FIRSTNAME","USER",Context._STARTUSER)
olast = ObjProperty("LASTNAME","USER",Context._STARTUSER)
sname = ctype(ofirst, string) & ctype (olast, string)
```
###  .NET Framework scripting for Optiva events
Some Optiva events are triggers for script functions.
- Save
- Update row
- Calculate formula

At the trigger event, an Optiva method looks for a .NET Framework function and the context object. After the script is executed, the method returns any existing messages from the script engine.

## Scripting in Workflows
Scripts are configured in the Action form. Action scripts are collected together into Action Sets. When userslaunch a workflow from an object form, they select an Action Set.
After a user initiates a workflow, the actions within the action set are displayed in the Pending Tasks grid.For each action, there are Event buttons. These buttons have labels that describe the function that the usercan perform for that action. Some examples are Edit, Reject, View, and Approve.
The system’s behavior after a user selects one of the buttons is defined by the script functions entered in theAction form. Using the Optiva scripting language and Visual Basic .NET provides many ways to control howyou do business.
Each action corresponds to one task. So, the Reject, View, and Approve buttons are scripted and enabledin a single action.
For example, when a user clicks Reject, the script for the Reject event is run.
The behavior for each button can vary for different types of workflows. What happens when a user clicksComplete in one workflow can be different than when another user clicks Complete in a different workflow.

## Scripting in Optiva object events
In addition to Optiva scripting for workflows, copy methods and equations, Optiva provides hooks at variousevents. These hooks enable Optiva scripts to run to alter or enhance these events.
- Save (i.e., pre and post)
- Update formula row (i.e., pre and post)
- Calculate formula (i.e., pre and post)

These scripts are defined in the Script tab of the Symbol form for each relevant object. You specify the scriptin .NET Framework scripting language. To verify the syntax of the script language, click Validate.

# Conventions for scripting

The first line of the script has the entry: Option Strict
1. When Option Strict OFF: Option Strict OFF reduces Visual Basic .NET validation by turning off extra error messaging as you createscripts. This enables you to take some shortcuts that reduce the development time. If Strict is OFF, then somesyntax errors can cause problems. Those problems are not caught through the validation. Specify Option Strict OFF as the first line to disable the extra validation. Otherwise, Option Strict ON isassumed.

2. When Option Strict ON: If Option Strict line is not in the script, the default value is ON; or you can add Option Strict ON to the firstline to turn on the Strict option. You must be specific about data types. You are notified when a conversion of data types can result in someloss of data; or when a specific operand cannot be used with a data type. With Option Strict ON, the configuration of the script must be more specific than if Option Strict is OFF. This is true, especially for functions like objProperty where the returns can vary. When the option is ON, implicit data type conversions are restricted to only widening conversions. This explicitlydisallows any data type conversions in which data loss can occur and any conversion between numeric typesand strings. Specifically, you can assign an integer value to a long variable because a long variable can hold more datathan an integer. You can assign a double value to a string because anything can be converted to a string. Youcannot assign a string to a double without explicitly converting it to a double (that is, Double d = CDbl(stringValue)).

## Templates for scripts
The forms for workflows, copy methods, and equations contain templates for scripts. You specify the scriptbetween the Function and End Function in each template.
```
Function wf_start() As Long
End Function
Function wf_return() As Long
End Function
Function wf_complete() As Long
End Function
Function wf_approve() As Long
End Function
```

In Workflow, you can delete events that do not have a script. You are not required to do so. The first action in a workflow must have some script in the wf_start event, even if it is only return 111.
## Class Script
Optiva VB Scripts (actions, copy methods, and event hooks) must have the class declared exactly as ClasstypeScript. Otherwise, an error occurs.
- Class can begin with uppercase C or lowercase c.
- Use uppercase S for Script. If you use class typescript, then an error occurs.
- The type declares the type of script:
> workflow script = ActionScript
event hook script = HookScript

## Workflow events and writing scripts in it
Each action consists of scripted events that determine what the user can do and how the user input is processedby the system. For example, if a user processing a formula approval workflow is asked to view the formula,then select Approve or Reject. This action contains the VIEW, APPROVE and REJECT events, even though only twoof the three events occur.
Selecting a button of the same name as the event in the user interface activates the scripted event; and itinstructs the system to open a Formula form, approve the formula or reject the formula. This collection ofevents constitutes one action.
They can be scripted as shown in the table.
|Event|Script|
|-|-|
|START|The first workflow event. If the workflow needs to pause then this function ends with return 1 else we can write return 111|
|VIEW|Opens the Formula form and instructs the system to advance to the next event. E.g. Dim rc As Long = StartForm("frmformula") Return 1|
|APPROVE| Shows this message to the user and instructs the system to advance to the next action. MessageList("You approved this formula.") Return 111|
|REJECT| Shows this message to the user and instructs the system to cancel the workflow. MessageList("This item rejected. Workflow cancelled.") Return 9111|
|COMPLETE| All events except APPROVE, with return code 111 or 811, also trigger the COMPLETE event. Return 111 instructs the workflow engine to finish the current event and then processthe COMPLETE event. If you do not want to process theCOMPLETE event, then use Return 1. Similarly, if youloop back with 811, this completes the current Action's COMPLETE event and then re-turns to the specified loop line. COMPLETE is a way to ensure a final bit of script is always run. If this is not appropriatefor your workflow, then use the APPROVE event. Do not configure COMPLETE; use Return111 with another event.|

Return codes conclude each event and instruct the system to move forward, pause or cancel the workflow.

# Variables
You can use local variables and global variables in your workflow scripts. You can use certain functions toretrieve parameter values, too. 
## Local Variables
Use arguments as local variables to hold values and use them in later statements. For example:
```
Dim oCls As Object = ObjProperty("CLASS")
```
The local variable oCls holds the value for the class of the object. Local variables are accessible only in theevent in which they were declared.
Variables, which are user defined, can be constructed of any letters. Infor encourages the use of meaningfulvariable names.
Adding a prefix that is the first letter of the data type ensures no conflict with reserved words. The prefix keepsthe declared data type with the variable throughout the workflow.
```
Dim oCls As Object = ObjProperty("CLASS")
```
## Parameter Values
Use WipParamGet to retrieve a parameter that was specified either by user input or by a previous action.
## Global Variables
Use global variables, such as _STARTUSER, as function arguments.

| **Workflow Global Variable** | **Data Type** | **Refers to** |
|------------------------------|--------------|---------------|
| _ACTIONCODE | String | The Action code |
| _ACTIONSETCODE | String | The Action Set code |
| _COMMENT | String | The comment that is entered by a user for an event that requires a reason code, a comment, or an electronic signature. |
| _OBJECTKEY | String | The object for the workflow |
| _OBJECTSYMBOL | String | The symbol of the object for the workflow |
| _BASEVALUE | String | Base value for an alternate unit of measure. |
| _REASONCODE | String | The reason code that is selected by a user for an event requiring a reason.|
| _REASONDESCRIP | String | This applies to an event that requires the user to specify a description for the reason code. |
| _SIGNEDBY | String | User who provided the electronic signature. |
| _SOURCEUSER | String | The current user, who performed the action. |
| _STARTUSER | String | The user who started the workflow. |
_TASKDB | String | The current name of the database name. This is the logical name of the database that is in the Optiva Configuration application. It is not the physical name of the database or its location. |
| _TASKGROUP | String | The group that is assigned to a task |
| _TASKLAB | String | The current lab code |
| _TASKROLE | String | The role that is assigned to a task |
| _TASKUSER | String | The user that is assigned to a task |
| _WIPALLUSERS | String | All users, groups, and roles that are involved in a workflow |
| _WIPGROUPS | String | All members of all groups that are involved in a workflow |
| _WIPID | Long | The Work ID |
| _WIPLINEID | Long | The line number for the workflow |
| _WIPROLES | String | All members of all roles that are involved in a workflow |
| _WIPUSERS | String | All users that are involved in a workflow |

# Variables in Copy methods
## Segments
For Create Rules, you define segments. You can add segments that are not part of the key. Then, these segmentscan collect information from the user. For example, if you are creating a project, you can add a segment sothat the user can tell you the description for the project.
To retrieve the value of a segment in a copy method, use the Context.GetSegData() method. You can thenuse ObjPropertySet to change the project description.

## Local Variables
A local variable can serve as an argument once a value has been assigned to it by a previous statement, usuallyinvolving another function.
In the example, the status of a formula is checked during a Copy Method: oStatus and lShow are localvariables, ObjProperty and MessageList are functions, and STATUSIND.STATUS is an argument that retrievesinformation from the the system database.
The first statement takes the status of the new formula and places it in the local variable, oStatus . Theif/then statement determines if the value of oStatus indicates that the formula is experimental or obsolete.If this is true, then the MessageList function displays a message to the user. The message is enclosed inquotation marks in an Alert Messages box.
```
Dim oStatus As Object = ObjProperty ("STATUSIND.STATUS", "","")
If (oStatus < 400 Or oStatus = 999) Then 
MessageList ("This formula is experimental or obsolete.")
End If
```
Brackets are used to specify optional arguments. You can omit an optional argument, if it is not needed toidentify the property. Required arguments must have a value or left as empty quotation marks ("").

## Global Variables
Function arguments can be global variables. For example, _TASKUSER represents the current user.

| **Copy Methods Global Variable** | **Data Type** | **Refers to** |
|---------------------------------|--------------|---------------|
| _MODELOBJECTKEY | String | The object code of the object being copied. This includes the version of the object. |
| _MODELOBJECTSYMBOL | String | The symbol, or type of object, of the object being copied. |
| _OBJECTKEY | String | The new object code, including the version. |
| _OLDCODE | String | The object code of the object being copied. If the object, such as a formula, supports versions, the version is not included. |
| _OLDVER | String | The version of the object being copied for objects that support version control, such as a formula. |
| _REFOBJECTKEY | String | The object code of the object on whose **References** tab an object is being created; used in copy methods for creating objects on the **References** tab of another object. |
| _REFOBJECTSYMBOL | String | The symbol of the object on whose **References** tab an object is being created; used in copy methods for creating objects on the **References** tab of another object. |
| _TASKUSER | String | The current user. |

## Variables in Equations
1. _basevalue (Double): The amount of the parameter in a formula. The unitof measure (UOM Extend/Per UOM) is displayed ac-cording to the parameter definition.
2. _CALCOBJECTKEY and _CALCOBJECTSYMBOL (String): The main usage of the _CALCOBJECTKEY and _CALCOBJECTSYMBOL variables is to perform object-specificalternate unit of measure equation conversions.
These two variables generally have the same valueas _ObjectKey and _ObjectSymbol, unless you run an alternate UOM conversionequation. In that scenario, these variables referencethe business object (Item, Formula, etc.).Both _ObjectKey and _ObjectSymbol reference the particular technical parameterthat is being converted.
3. context.ReturnValue (String): For item calculated equations, use the variable context.ReturnValue.
This variable defines what is specified for the param-eter. A return code of 1 completes the equation. Forexample, the equation 2+2 looks like this:
```
Dim A As Integer = 2+2
Context.ReturnValue = A
Return 1
```
4. _ObjectKey and _ObjectSymbol (String): The object key and the symbol of the object that isbeing calculated. Or, the parameter code of theparameter for which the alternate unit of measureis calculated.
For example, FS-0001\0001 is an object key for FORMULA.

# Return Codes
Return codes conclude workflow events, conditional statements, and equation calculations. Return is a VB.NETkeyword. The value that is returned determines the behavior of Optiva.
- A negative value results in an error. An error message can be displayed by the system, but not necessarily.If the calculation is part of a SAVE, the save is ignored, too.
- A positive value results in a non-failure state.
-  Any portion of the script that is located after the Return keyword is not processed. The control is passedback to the system immediately.

| **Return Code** | **Description** |
|-|-|
| Return negative_value | Indicates an error. Depending upon the script, an error dialog can be shown to the user. In this scenario, the text of the error is retrieved from the FsAppMessageHeader and FsAppMessageText tables. |
| Return 111 | Advance to the next action. Use in a Start event for a single object to close the **Start Workflow** form. |
| Return 811 | Loop to the specified line in the action set (i.e., to repeat tasks). |
| Return 9111 | Cancel the workflow. If this is from a Start event for a single object, the **Start Workflow** form does not close. All the parameters that are entered by the user are retained and corrections can be made. The user can start the workflow again and a new work ID is created.<br>Use 111 to close the **Start Workflow** form without enabling the user to try to start the workflow again.<br>Cancel the creation of a new object from a copy method. |
| Return 9112 | Suspends a WIP when an unexpected error happens. The status of the WIP is set to "Suspended". Steps with the Queue Status "In Process Queue" is set to the Queue Status "Closed". The Receiving Indicator column will be set to "Suspended" for these steps as well. Steps with the Queue Status "Not Started" remain in that status, and the Receiving Indicator is unchanged for these steps. This will allow an administrator to fix the error and then put the Workflow back into an active state. |
| Return 1 | Hold and wait for additional user input.<br>Continue with the formula calculation.<br>For item calculated equations, use the variable `context.ReturnValue` to define the parameter value. A return code of `1` completes the equation.<br>For example, the equation 2+2 looks like this:<br>```<br>Dim a As Integer = 2+2 Context.ReturnValue = a Return 1``` |

# Detail Codes
 Detail codes can be used by functions such as ObjectXSD, ObjectXML, and ObjProperty. Business objects suchas formulas, items, specifications, projects, and company, all use tags that are not special to other objects.
 | **Detail** | **Description** | **Special To** |
|-|-|-|
| ALTINGR | Alternate ingredient | Formula |
| ALTUM | Alternate unit of measure | Parameter |
| APPROVAL | Approval Code for symbol | Symbol |
| APPSETTINGS | Application settings for user | User |
| ATTACH | Attached documents, files and URLs | |
| BYPROD | ByProduct | Formula, Specification |
| CONTEXT | Context Attributes | |
| CUSTOM | Extension field for objects.<br>Single-value extension fields are included in the Header data. | |
| DEFAULTS | Standard work time | Shop Calendar |
LINELIST | Explodes but does not combine like-item codes from any level. Explodes down to the constituent item level. The intermediate item codes, including raw materials that have constituent formulas, are shown.<br>Cannot be used with `ObjProperty`, `ObjPropertySet` or `ObjectsXMLForeign`. | Formula |
| LKUP SEARCHES | Saved search for lookups | Symbol |
| MATRIX | Extension tables | |
| NOTIFICATION | Controls the loading and saving of notifications on fields.<br>You can also retrieve the current notifications by using `ObjProperty` or `ObjectDataSet`.<br>Formatting the detail code data table is left to the scripting author. | Analysis, Analysis Search, Calendar, Company, Company Search, Formula, Formula Search, Global Calc, Guidelines Search, Item, Item Search, Item Guidelines and Restrictions Search, Ingredient Replace, Ingredient Statement, Ingredient Statement Rule, Item, Label Claim, Label Claim Definition, Label Claim Rule, Label Content, Label Content Search, Label Search, Label Statement Group, Parameter Guidelines and Restrictions Search, Project, Project Search, Rule, Sample, Sample Search, Specification Search, Test Search, Test Order Search, Specification, Test, Test Order, Workflow in Progress |
| ORDERDTL2 | Tests | Test Order, Test Results |
| PER | Security | |
| PHRASETEXT | Text for phrase | Phrase |
| PROFILE | Profile | User |
| REF | From the **References** tab | |
| REFERENCE | Reference code | Symbol |
| REFSTATUS | Reference status | Symbol |
| ROLE | User Role | User |
| SETS | Set classification for symbol | Symbol |
| SET | Set membership | |
| STATUS | Status | |
| SYMDOC | Attached document or embedded object. | Symbol |
| TESTCOND | Test condition | Test Method |
| TPALL | Parameter | Formula, Item, Project |
| TP | Parameter | Specification |
| TP | Parameter Values for all Rowtags | Label Content |
| TPVAL | Parameter Values per Rowtag | Label Content |
| VIEWS | View | |

This information is added to the FsFormulaTechParam table:
- any parameters for this formula
- parameters for the lab that the formula belongs to, if those parameters are different than the formulas parameters

Parameter values for manufactured items are saved to the primary formula in the system. You can use theimport function to update parameters for formulas of manufactured items.

### Tips for Wizards
Many Optiva scripting functions enable you to pass in empty quotes for the Symbol and Object parameters.The empty quotes indicate that you are using the Symbol and Code of the object that the script is runningagainst.
Unlike most other scripts, the Wizard can be run without a target object. Consequently, there is no place forthese functions to retrieve a default Symbol or Object name.
If you are writing a script for a Wizard, you must provide a non-blank Symbol and an Object parameter forthose methods that require one. You cannot use empty quotes for the Symbol and Object.
The global variables _OBJECTKEY and _OBJECTSYMBOL are blank in this scenario. You should not write ascript that relies on these variables having some specific value.
Note: Although most wizards are used to step users through a series of inputs, the ability of a wizard to berun from the Home page (without a starting data object) may open up other opportunities that do not actuallyrequire user input. In this case, you can create a wizard workflow with all of the code in a single action’s Startevent so that it will execute without waiting for user input. You must use StartForm() at the end of the eventto open some data object (specify the form name and object key). To use StartForm to return to the home
page, use any word as the first parameter and either no second parameter, or an empty string for the firstparameter:
`Startform("Home")`
# Script Library
You can create custom scripts using the Script Library form. These scripts are stored in the FSSCRIPTLIBRARYtable and can be called from a Workflow Action, Copy Method, Equation, or Symbol script.
## Creating the script library
You can use the Script Library form to create a script. After you specify a unique name for the script, thescript framework defaults automatically. The Class Name is added to the form by the system. The Class Nameis the same as the Script Name.
You can change the Class Name, but it must be unique. Ensure that the Class Name is the same in both placesin the form.
The Script Class allows for organization of scripts. Infor Standard and Infor IA are scripts provided with theproducts and should not be edited
Infor Scripting Samples are samples for use to create Custom Scripts. Customer class can be used for customscripts.
In the Script Library field, you must add the custom code. The script library requires the context (co = context) of the calling script in the constructor to execute the system functions.
The Script Library > Reports tab displays the script description and dependent classes where the script hasbeen used in the Where Used in Scripts section.
Note: The informational and error message text has been corrected to pinpoint the proper script library wherethe errors occur. You must specify the value for the SCRIPTLIBRARY.IGNOREERRORS to 1. See the profile attributetopic in the Infor PLM for Process Application Configuration Guide.
## Calling the script library
The Script Library must be created before it can be called. It can be called using this syntax:
`Dim scr As <LibraryScriptName> = new <LibraryScriptName>(Me)`
Example:
`Dim Addset As ADDSETS = new ADDSETS(Me)`
In this example, a Formula Copy Method calls the ADDSETS Script Library. The ADDSETS script adds the SNACKS Classification set to the newly copied formula
```
Function execute( ) As Long
Dim _returnvalue As Long
‘Call the library script function Addset.AddSnacks( )
Dim Addset As ADDSETS = new ADDSETS(Me)
addset.AddSnacks( )
Return _returnValue
```
Then copy a formula using the Copy Method. The classification sets for the new formula includes the sets from the original formula and the additionalSNACKS set. The ADDSETS Script Library was called to add the SNACKS set.
You can use a script library function within another script library function. This enables you to reuse evenmore functions without having to place them in one set of coding.

## Constant Values
Constant values, such as FDA recommended values, can be used in product calculations. These values canbe:
- Hard-coded into a Script Library form with a declaration such as:
`Public CONST dailyval As Integer`
- Called from Equations or Workflow scripts in a fashion similar to:
`MessageList(libraryname.dailyval)`
Defined in a custom database table or in the Daily Value table (FSDVRULETECHPARAM) and returned tothe Script Library using a Table Lookup query.
- Maintained on an extension table with access to both the single object (e.g., = @CONST) and the extensiontable. Values can be returned to the Script Library using the ObjProperty function.
- Maintained in the Daily Value Rules form, if the Analysis module is installed. The value is returned to theScript Library using a script such as:
`Dim DV As Double = DailyValueRule("", "CANADA")`
- Retrieved as a value from the Daily Value Rule table using syntax similar to:
`Dim val As Double = ObjProperty("DAILYVALUE.PARAM", "DVRULE", "VIT_DEGRADE", "VITAMIN A", 2)`
This assumes that you have a Daily Value Rule called VIT_DEGRADE. It also assumes that you added a valuefor VITAMIN A to that rule.

# Secure Scripting
You can execute Optiva Scripting functions in a secured mode. This feature limits the amount of damage amalicious user can inflict should they gain access to your scripts. You run all scripting in a sandbox environmentthat has limited permissions.
Customers who are concerned about security or who are exposing Optiva to the internet should considerrunning in this mode.
In the Web.config file, a parameter controls whether you run Scripting in a secured or unsecured environment.See the Infor PLM for Process Installation Guide.
Script hook considerations: DataRowCurrent and DataRowProposed are not available in the secured scripting environment. They have beenreplaced by GetCurrentRowValue, GetProposedRowValue, and SetProposedRowValue. The new script hooks allowaccess to a single value on the row, not the entire row.
## Secured scripting with objectdataset
In secured scripting mode, the ObjectDataSet command cannot work simultaneously with other Optivascripting API’s. To work directly with the data set of an object, you must call the ObjectDataSet command andthen make changes to the DataSet object.
You can call any of the other Optiva scripting API’s to change or view information that is stored in the system.In this case, the information in the DataSet object is flushed back to the system. The DataSet is destroyed.Any attempts to access this DataSet result in .Net exceptions being thrown.
You can make a new call to ObjectDataSet to retrieve a new copy of the DataSet. Then, you can make additionalchanges.
## Secured Scripting for RowUpdate()
Only use RowUpdate for symbols or details that support rowupdate events in hook scripts. For example,ingrpre(and post)rowupdate on formula Item lines, tppre(and post)rowupdate on Spec Parameters.
## Secured Scripting for creating new whitelist entries
You must use the prefix CUSTOMREQUESTREGEX when creating new whitelist entries for your configurationdatabase. Entries with this prefix will be also added to the Secured Scripting whitelist. These entries muststart at 1 (for example, CUSTOMREQUESTREGEX1) and must be incremented by one for each entry. If an entry ismissing, subsequent entries will be ignored.