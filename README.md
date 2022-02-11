
<a href="https://www.cortex-ia.co.uk/" target="_blank"><img src="https://github.com/CortexIATest/CTXImages/blob/master/Cortex-350-120.png" alt="Welcome to Cortex!" width="350" height="120" border="0"></a>

# Excel Cortex Module
The current Cortex Excel module focuses on interactions with .xslx files.

It offers the following functionality:
1.	Read from file
2.	Write to file
3.	Format file contents
4.	Create SQL insert statements from file content
5.	Write data to file from a SQL statement
6.	Retrieve information about file
7.	Convert file to a csv



## Table of Contents
1) [Dependencies](#dependencies)
    * [Cortex Version](#cortex-version)
    * [OCIs](#ocis)
    * [Files](#files)
    * [Other](#other)
1) [Support and Warranty](#support-and-warranty)
2) [Installation](#installation)
3) [How to use](#how-to-use)
4) [How you can contribute](#how-you-can-contribute)
5) [Versioning](#versioning)
6) [Licensing](#licensing)


## Dependencies
### Cortex Version
This version of the Excel module was developed in Cortex version 7.1. Some functionality may not be available on different versions of Cortex.

### OCIs
The  module requires the following Cortex OCIs:
* PowerShell

### Files
The Excel module requires the following files:
* [CTX-Excel.studiopkg](https://github.com/CortexIntelligentAutomation/CTX-Excel/releases/download/v1.0/CTX-Excel.studiopkg)

### Other
The Excel module has the following additional requirements:
*  PowerShell v5.0 or later installed on the same machine as the PowerShell OCI

## Support and Warranty 
This module is supplied as a template that you can amend and extend to fit your requirements, as such it is not supported as part of the Cortex Product suite under the Cortex product support agreement.

## Installation
Details of the installation can be found in the [CTX-Excel Deployment Plan](https://github.com/CortexIntelligentAutomation/CTX-Excel/blob/master/CTX-Excel%20-%20Deployment%20Plan.pdf).
## How to use
A detailed User Guide has been provided with instructions on how to use the flows/subtasks, available [here](https://github.com/CortexIntelligentAutomation/CTX-Excel/blob/master/CTX-Excel%20-%20User%20Guide.pdf). Configuration of subtask inputs and outputs are detailed in notes on the subtask workspace.

## How you can contribute
Unfortunately, we cannot offer pull requests at this time or accept any improvements.

## Versioning
CTX-Excel has the following versions, starting with the most recent:

Version | Release | Functionality | Notes
------------ | ------------- | ----------- | -----------
v1.0 | 17/11/2021 | PEPR-Process-Excel-PowerShell-Response | Created 
v1.0 | 17/11/2021 | Import-Excel | Created 
v1.0 | 17/11/2021 | Export-Excel | Created 
v1.0 | 17/11/2021 | Set-ExcelRange | Created 
v1.0 | 17/11/2021 | ConvertFrom-ExcelToSQLInsert | Created 
v1.0 | 17/11/2021 | Send-SQLDataToExcel | Created 
v1.0 | 17/11/2021 | Get-ExcelSheetInfo | Created 
v1.0 | 17/11/2021 | Get-ExcelWorkbookInfo | Created 
v1.0 | 17/11/2021 | ConvertFrom-ExcelSheet | Created 

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


