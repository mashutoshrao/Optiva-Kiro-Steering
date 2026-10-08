---
inclusion: always
---
<!------------------------------------------------------------------------------------
   Add rules to this file or a short description and have Kiro refine them for you.
   
   Learn about inclusion modes: https://kiro.dev/docs/steering/#inclusion-modes
-------------------------------------------------------------------------------------> 
Infor PLM Optiva is a Process Lifecycle Management system for Food & Beverage and Chemical industries. This guide covers scripting conventions, object models, and best practices for Optiva customization.

# Objects

There are many types of Objects in Optiva to store particular type of data

## NPDI objects
NPDI stands for "New product Development" in Optiva.
|Object|Description|_OBJECTSYMBOL|
|-------------------------------|-----------------------------|-----------------------------|
|Items|Supports the definition of raw materials, intermediate items, packaging items|ITEM|
|Formula|Supports experimental,bulk, intermediate formulas, sub-formulas,packaging formulas, and their properties and technical parameters.|FORMULA|
|Specification|Defines the goal of the formula or item. Details thepermittedquantities, acceptable ranges andtarget values, suchas Viscosity, Color, Alcohol content, or Unit cost.|SPECIFICATION|
|Project|A compilation of information from all departments thatis involved with the development of a product, keeping it accessible. Information thatis stored includes:time frame, key people, formulas, specifications, intended market, product brand, production and selling locations. Advanced Program management is used to manage project stages and tasks and deliverable dates.|PROJECT|


## Labelling objects
|Object|Description|_OBJECTSYMBOL|
|-------------------------------|-----------------------------|-----------------------------|
|Ingredient Statements|A listing of ingredients and their quantities in a formula. This can be formatted in sections and to show or hide sub-ingredients and quantities. The statement can be used as part of a product label.|LABEL|
|Label Content|A container for Optiva data thatis used to create an actual label. Itincludes ingredient statements, nutrition values from Analysis,text statements and produces a package label report.|LABELCONTENT|


## Configuration components
|Object|Description|_OBJECTSYMBOL|
|-------------------------------|-----------------------------|-----------------------------|
|Action Sets (or) Workflow processes|Business tasks to be performed. Actions can be performed in series or in parallel.|
|Actions|These are individual tasks of the action sets. Scripts are written in this object to automate the processes of a workflow.|
|Parameters|Characteristics, for example, Viscosity, Color, Alcohol content, or Unit cost that is defined by the customer’s equations or methods.|TECHPARAM|
|UOM|Unit of measures usually added to the parameters, Item lines, etc. |UOM|

- Optiva objects build on each other

In Optiva, one object can act as a building block for another object:
- Units of Measure are used to define Parameters.
- Parameters are used to define Items.
- Items are used to define Formulas.
- Formulas are grouped into Projects.

# Managing objects
Optiva objects include formulas, items, projects, specifications, ingredient statements, guidelines, tests, etc.

Optiva provides search capability. The administrator can configure components such as enumerated lists,
symbols, and users.

This table shows you how to open, create, save, and delete objects in Optiva:
|Type|How to?|
|-------------------------------|-----------------------------|-----------------------------|
|Function Method|Menu > Object Type > Object Type > object or Open an object and click the Lookup on the object toolbar. or Select an object from the Recently Viewed list on the home page or the side-bar of an open object. or Open Formula form or object in a new tab or window.|
|Open an object|Menu > Object Type > Object Explorer > New Object button Open an object. Toolbar > New or Save As buttons or References tab > Create New or Create From buttons This creates the objects and adds the new object to the References tab.|
|Create an object|Toolbar > Save button Refresh the object to see changes made by other users. You can create new versions of objects by saving and renaming the current object. Depending on how the system administrator has implemented security, you can make changes to a new object repeatedly. But the default record does not give you security privileges until you log back into the system.|
|Save an Object|Object > Toolbar > Delete button. Object > Lookup dialog > Delete button Deleted objects remain in the database, tagged as deleted, although they cannot be seen in the lookups. Your database administrator can restore deleted objects and those objects can show in lookup lists.|

## Rules control creating new objects
Rules control the names and properties of new objects. A Create dialog can be used to enter the attributes for a new object. Your system administrator configures the rules to control the naming of the new object according to your business practices.
In addition to naming new objects, rules can also include copy methods that configure properties of new objects. For example, a Rule can assign a new object to a set or remove context information. Sometimes, new objects inherit data from the default record only.

## Default record data copied to a new empty object
When you generate a new object, information that is contained in the default record is inherited by the new object (formula, item, specification).
The only data not copied from a default record to a new record is are: Views, System Status Information (Creation Date, Created By, Modify Date, Modify By), Master Formula check box

## Copying objects
When you copy an Optiva object, most information is brought into the new object. Some information is copied from the default record instead.

By default, this information is copied from another object:
- Data from all tabs
- Sets that are assigned to the object
- Context attributes
- Document attachments
- Data that is not copied is inherited from the default record, set by the system or added as work is completed with the objects.

# Where to Save configuration files
- Action Set (Workflow) needs to be saved folder inside build. Each Action sets will have separate folder. For each action inside workflow individual actions needs to be stored as vb scripts which should prefix <Task_Number>_<Action_Name>.vb (TaskNumber should be 00,01,02,...).
- Copy Method: build\Copy Method
- Query: build\Query
- Script Library: build\Script Library
- Order of Saving the files: Query > Script Library> Copy Method > Action 