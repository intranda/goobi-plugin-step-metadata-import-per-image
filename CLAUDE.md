# Critical rules
Perform all actions in small chunks that can be reviewed.
Prepare good commits with meaningful commit messages. Commit messages may not contain information about Claude or the way the code was created.

# Context
This project is a step plugin for the Goobi workflow web application. In this webapp, digitization is done in processes (representing books or documents)
that follow a pre-defined digitization workflow. A workflow consists of multiple steps. Each step is either manual or automatic. 
Steps could perform shell actions or reference step plugins. A step plugin can have a UI for the user to interact with or only backend
Java code to be executed. When a step is done, the next step is opened. Processes store metadata for the digitized object in MODS/METS format, one XML file
per process called `meta.xml`. In this metadata file, a logical structure of the object can be defined (e. g. a book consists of chapters and sections). Each
structure level can have nested levels and metadata assigned to it. The possible metadata and structure levels are defined in a 
so-called ruleset. Each project may have a different ruleset. The main application is located in `../workflow-core`.

## Step plugins
In the `install/` folder is the configuration file for the plugin.
Everything configurable will be done there in XML. In `module-base/` is the Java source code of the plugin. The plugin 
is implemented with an interface with an `initialize` and `run` method. This is the entry point for the plugin. You have some placeholder
logic for reference there.

# Scope
This plugin will read an Excel file (xslx) and create a logical metadata structure based on the file. The Excel sheet contains
multiple columns, not all are relevant. Each row, except for the header, represent one image to be digitized. At the moment
of execution, all images are already present in the `master` folder. Based on values in the columns per row, the plugin needs to
generate or update the metadata structure of the process. The plugin needs to be very flexible and configurable in its behavior.
In the current use case, we have the following structure: `WholeItem` > `ArchivalObject` > `Folder` > `Item`. These metadata types
need to be configurable and may vary in their number. The top-level structure type is used for the digitized object. The rest of the
structure hierarchy needs to be derived by some configurable rules. In this use case, the Excel sheet contains a colum `URI` with some
identifiers. If the identifier stays the same, all rows belong to the same `ArchivalObject`. If it changes, we start a new
`ArchivalObject`. The column `Structure` contains information about the `Folder`. If no value is given, we need a fallback to
a `Folder` named `Unnamed folder` (configurable). Otherwise, the same logic applies. Same value is the same `Folder`, new
value represents a new `Folder`. Column `Label` contains pagination information and needs to be copied as is to the image pagination.
This is saved in the `meta.xml` file as well. The `Caption` column contains titles for the `Item`s. Again, the same logic applies.
If the `Caption` stays the same, the image belongs to the same `Item`. If it changes, the `Item` is a new one.

## Validation
Before applying any changes to the `meta.xml` file, the plugin needs to make sure that all given information is valid and expected.
For example, the number of rows in the Excel sheet (excluding the header), matches the number of images. Other issues should be
avoided as well. Make good suggestions.

## Backup
When the validation passes and the plugin executes, it creates a backup copy of the `meta.xml` file beforehand. The core application
contains logic to do this.

# General
Create a plan for the implementation prior to execution.
If you find additional information that is relevant for the context, suggest to update the context accordingly for future sessions.
Keep the context up-to-date with implementation progress to ease future sessions.
During implementation, use teams when appropriate. Always stop agents that are no longer needed.