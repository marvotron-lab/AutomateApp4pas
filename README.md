# AutomateApp4pas
Delphi solution that allows to automate several office software programs (eg. Ms Office Excel, LibreOffice Calc, Outlook etc.) via a unified interface.

This lib tries to establish a unified way to access and control these programs from Delphi. It is a solution that mostly uses OLE automation (late binding) to communicate with the target programs. 

## How does it work?
For each kind of application (eg. spreadsheet software) a common interface is available, that can be used to automate the program. For each target application (eg. Excel) exists an interface that implements this functionality.

## State 
This is a work in progress (alpha). We are trying to migrate existing code to this framework. The code will change the API is not stable yet.

## Examples
TODO

## References
* `TIHojaSpreadsheetApp` uses `UHojaCalc.pas` from https://github.com/sergio-hcsoft/Delphi-SpreadSheets (with a public-domain / unrestricted license) - this file is contained in this repository. This is an existing lib for accessing Excel and LibreOffice Calc.

## License
Copyright 2025 yonojoy@gmail.com.

This project (all files in this repository) are subject to the terms of the Mozilla Public License, v. 2.0.
