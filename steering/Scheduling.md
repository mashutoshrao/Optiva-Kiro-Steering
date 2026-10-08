---
inclusion: auto
name: Scheduling
description: Use this file if asked for Scheduling of tasks, Running a workflow via XML.
---
<!------------------------------------------------------------------------------------
   Add rules to this file or a short description and have Kiro refine them for you.
   
   Learn about inclusion modes: https://kiro.dev/docs/steering/#inclusion-modes
-------------------------------------------------------------------------------------> 

# Scheduling a script

You can schedule a workflow task to run daily or weekly. For example, use a VBScript script to pass an Action SetStart XML file to the XML Listener.

ActionSetStart XML does not only import an object into Optiva. It launches a workflow. Using the general task scheduling capability in Windows, the script can be scheduled to run at specific times. For example, the script can run the same time every day or every week.

The XML Listener requires a valid entry for ActionSetStart in the AUTOIMPORT definition.

*   Symbol=ACTIONSETSTART
*   Detail Codes=HEADER;PARAM;OBJECTLIST
*   Option=Replace Existing Data.

## Example of an ActionSetStart XML

In this example, the Workflow Action Set is TEST_ACTION_SET The User Code to launch the workflow is ADMIN. The workflow is launched on formula FML048\0001. The Object Symbol (i.e., object type) is FORMULA. The SRVROW_ID node is required and must match the **Input No** column in the **Action Set > Input grid**.

```xml
<?xml version="1.0" encoding="utf-8" standalone="yes"?>
<fsxml>
  <FSACTIONSETSTART>
    <KEYCODE>TEST_ACTION_SET</KEYCODE>
    <FSACTIONSETSTARTOBJECTLIST>
      <OBJ_LINE_ID>1</OBJ_LINE_ID>
      <TARGET_OBJECT_KEY>FML048\0001</TARGET_OBJECT_KEY>
      <USER_CODE>ADMIN</USER_CODE>
      <ACTIONWIP_ID />
      <TARGET_OBJECT_SYMBOL>FORMULA</TARGET_OBJECT_SYMBOL>
    </FSACTIONSETSTARTOBJECTLIST>
  </FSACTIONSETSTART>
</fsxml>
```
# Running a workflow with inputs via XML

You can include values for the input of a workflow via XML. You must provide a FSACTIONSETSTARTPARAM node for each input in the XML.

This example workflow has a single input:

```xml
<FSACTIONSETSTARTPARAM>
  <ACTIONSET_CODE>TEST_ACTION_SET</ACTIONSET_CODE>
  <SRVROW_ID>1</SRVROW_ID>
  <PARAM_CODE>REPORTFORMAT</PARAM_CODE>
  <LANGUAGE_CODE>EN-US</LANGUAGE_CODE>
  <REQUIRED_IND>1</REQUIRED_IND>
  <PVALUE>XML</PVALUE>
</FSACTIONSETSTARTPARAM>
```

You must list the ACTIONSET_CODE, SRVROW_ID, PARAM_CODE and PVALUE nodes. The ACTIONSET_CODE node must match the KEYCODE node in the Workflow Action Set. The LANGUAGE_CODE and REQUIRED_IND nodes are not necessary. The other nodes will assume default values.

