# Export-VBA-Code
Code to automate export VBA code out Word and Excel documents so the code can be reviewed by version control software.

# Introduction

I started to make complicated Office automation files. To help remembering all the small changes I made I came across version control software. I started out with Bazaar and loved the aspect of documenting and retreiving all necesary information for development. To effectively use bazaar i needed to export all code modules outside `*`.doc, `*`.xls, `*`.ect from Office files into text files. This was a tedious and error-prone exercise, so I wanted to automate the procces.

# Important

For this code to work the following needs te be accounted for:
- Office (This was made for Office 2010, code should however transfer to other version with slight modifications)
  - Word (required)
  - Excel (required)
- References should be set to the following modules inside VBA
  - Microsoft Excel 14.0 Object Library
  - Visual Basic Extensibility Version 5.3
  - Microsoft Visual Basic Extensibility Version 5.3
  - Microsoft Scripting Runtime
  - Microsoft VBScript Regular Expressions 5.5

# How-to-use-it?

1. Open the Export Import VBA.doc file
2. Press the Export VBA button (Call to VBA_Export procedure)
3. Select the folder containing VBA coded `*`.doc, `*`.xls, `*`.ect (Word-, Excel-Files)
4. A folder will be created next to a file with "Filename VBA Code"

# Note

Not all procedures and/or references are used during exporting. They where used exploring all available methodes and kept for future enhancements.
