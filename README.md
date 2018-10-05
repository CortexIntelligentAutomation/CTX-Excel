<a href="https://www.cortex-ia.co.uk/" target="_blank"><img src="https://github.com/CortexIATest/CTXImages/blob/master/Cortex-350-120.png" alt="Welcome to Cortex!" width="350" height="120" border="0"></a>

# CTX-Excel
Cortex Subtasks which interact with Microsoft Excel without requiring Microsoft Excel to be installed on the Cortex application server.

The module allows users to perform the following functionality:
* Add Conditional Formatting
* Create Pivot Charts
* Create Pivot Tables
* Create Workbook
* Create Worksheet
* Insert Headers
* Insert Rows
* Set Cell
* Set Cell Where


## Table of Contents
1) [Dependencies](#dependencies)
    * [Cortex Version](#cortex-version)
    * [OCIs](#ocis)
    * [Files](#files)
    * [Other](#other)
1) [Installation](#installation)
1) [How to use](#how-to-use)
1) [How you can contribute](#how-you-can-contribute)
1) [Versioning](#versioning)
1) [Licensing](#licensing)

## Dependencies
### Cortex Version
This version of the CTX-Excel module was developed in Cortex v6.3.0. Some functionality may not be available in earlier verions of Cortex.

### OCIs
The CTX-Excel module requires the following Cortex OCIs:
* PowerShell

### Files
The CTX-Excel module requires the following files:
* [CTX-Excel Studio Package](https://github.com/CortexIntelligentAutomation/CTX-Excel/releases/download/untagged-735f460df6f7d65c9d19/Cortex.Studio.Package.-.V2.3.studiopkg)

### Other Requisites
The CTX-Excel module has the following additional requirements which are explained in detail in the [Deployment Plan](#Installation):
* PowerShell v5 to be installed on the application server
* PSExcel PowerShell module installed
* ImportExcel PowerShell module installed

## Installation
Details of how the module should be imported into Cortex can be found in the [Deployment Plan](#Installation).

## How to use
A detailed User Guide has been provided with instructions on how to use the module, available [here](https://github.com/CortexIATest/CTXExcel/blob/master/CTX-Excel%20-%20User%20Guide.docx). Configuration of each flow/subtask's inputs and outputs are detailed in notes on the flow/subtask workspace.

## How you can contribute
Unfortunately, we cannot offer pull requests at this time or accept any improvements.

## Versioning
The CTX-Excel module has the following versions, starting with the most recent:

Version | Release | Functionality | Notes
------------ | ------------- | ----------- | -----------
v2.3 | 25/09/2018 | Create Pivot Tables | Fixed Bugs
v2.3 | 25/09/2018 | Create Pivot Charts | Fixed Bugs
v2.2 | 05/01/2018 | Add Conditional Formatting | Created
v2.1 | 21/11/2017 | Create Pivot Charts | Created
v2.0 | 16/11/2017 | Create Pivot Tables | Created
v1.0 | 12/05/2017 | Create Workbook | Created
v1.0 | 12/05/2017 | Create Worksheet | Created
v1.0 | 12/05/2017 | Insert Headers | Created
v1.0 | 12/05/2017 | Insert Rows | Created
v1.0 | 12/05/2017 | Set Cell | Created
v1.0 | 12/05/2017 | Set Cell Where | Created

## Licensing
All functionality within this module is licensed under the [Apache 2.0 License](https://www.apache.org/licenses/LICENSE-2.0).

Copyright 2018 Cortex Limited

Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

    http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.
