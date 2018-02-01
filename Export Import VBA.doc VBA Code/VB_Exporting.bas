Attribute VB_Name = "VB_Exporting"
'Reference set to Microsoft Excel 14.0 Object Library
'Reference set to Visual Basic Extensibility Version 5.3
'Reference set to Microsoft Visual Basic Extensibility Version 5.3
'Reference set to Microsoft Scripting Runtime
'Reference set to Microsoft VBScript Regular Expressions 5.5
Public myWord As Word.Application
Public myExcel As Excel.Application

Public Type ProcInfo
    ProcName As String
    ProcKind As VBIDE.vbext_ProcKind
    ProcStartLine As Long
    ProcBodyLine As Long
    ProcCountLines As Long
    ProcScope As ProcScope
    ProcDeclaration As String
    ProcComments As String
End Type

Public Enum ProcScope
    ScopePrivate = 1
    ScopePublic = 2
    ScopeFriend = 3
    ScopeDefault = 4
End Enum

Public Enum LineSplits
    LineSplitRemove = 0
    LineSplitKeep = 1
    LineSplitConvert = 2
End Enum

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal ClassName As String, ByVal WindowName As String) As Long        'Windows API
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hWndLock As Long) As Long                                                     'Windows API
    
Sub VBA_Export()
Dim FileSystem As New FileSystemObject                  'Microsoft Scripting Runtime
Dim strFolder As String
Dim myFile As File                                      'Microsoft Scripting Runtime
Dim myFolder As Folder                                  'Microsoft Scripting Runtime
    
Dim OutputWb As Workbook                                'Microsoft Excel 14.0 Object Library
Dim OutputFileRg As Excel.Range                         'Microsoft Excel 14.0 Object Library
Dim OutputFileProceduresRg As Excel.Range               'Microsoft Excel 14.0 Object Library
Dim ws As Excel.Worksheet                               'Microsoft Excel 14.0 Object Library
Dim OutputFilews As Excel.Worksheet                     'Microsoft Excel 14.0 Object Library
Dim OutputFileProceduresWs As Excel.Worksheet           'Microsoft Excel 14.0 Object Library
    
'=======================================================================
'This procedure starts the exporting of Word and Excel Files containing _
VBA files. This procedure creates new instances of the Word.Application _
and the Excel.Application. If the procedure is terminated midway then _
the User needs to terminate the WinWord.Exe and Excel.exe processes in the _
Taskmanager by hand.
'=======================================================================
    
    On Error GoTo ErrHndl:                              'Errors are handled in ErrHndl label
    'On Error GoTo 0
    Set myWord = New Word.Application                   'Create new Word.Application instance
    Set myExcel = New Excel.Application                 'Create new Excel.Application instance
    
    'Suppress Auto_Macros from firing when document is openend. This prevents pop-ups from appearing when opening files.
    myWord.WordBasic.DisableAutoMacros 1                'Set security setting to 1 (High disables Auto_Macros)
    myExcel.EnableEvents = False                        'Suppres Events
    
    strFolder = GetWordFolderContainingVBA              'Ask users for a folder input though a FolderPicker dialog. Can use some error handling in case user cancels operation
    Set myFolder = FileSystem.GetFolder(strFolder)      'Set myFolder as input folder from the user
    
    
    'Create output workbook. This summary will contain a list of all found files and according procedure names and comments inside the code.
    Set OutputWb = myExcel.Workbooks.Add                'Create a New Workbook inside Excel.Application
    
    'Format output ws
    For i = 1 To OutputWb.Worksheets.Count                                                  'Loop over all Worksheets inside output workbook
        Set ws = OutputWb.Worksheets(i)
        Select Case i
        Case 1: ws.Name = "File Names"                                                      'Format first worksheet
                Set OutputFilews = ws
                Set OutputFileRg = OutputFilews.Range("A1")
                OutputFileRg.Offset(0, 0) = "BuildFileName"
                OutputFileRg.Offset(0, 1) = "File Path"
                OutputFileRg.Offset(0, 2) = "File Name"
                OutputFileRg.Offset(0, 3) = "Protection"
                OutputFileRg.Offset(0, 4) = "Count VBComponents"
                Set OutputFileRg = OutputFileRg.Offset(1, 0)
                
        Case 2: ws.Name = "File Procedures"                                                 'Format second worksheet
                Set OutputFileProceduresWs = ws
                Set OutputFileProceduresRg = OutputFileProceduresWs.Range("A1")
                OutputFileProceduresRg.Offset(0, 0) = "File Name"
                OutputFileProceduresRg.Offset(0, 1) = "Module Name"
                OutputFileProceduresRg.Offset(0, 2) = "Type Module"
                OutputFileProceduresRg.Offset(0, 3) = "Procedure naam"
                OutputFileProceduresRg.Offset(0, 4) = "Type Procedure"
                OutputFileProceduresRg.Offset(0, 5) = "BodyLine"
                OutputFileProceduresRg.Offset(0, 6) = "Count Lines"
                OutputFileProceduresRg.Offset(0, 7) = "Start Line"
                OutputFileProceduresRg.Offset(0, 8) = "Procedure Scope"
                OutputFileProceduresRg.Offset(0, 9) = "Procedure Declaration"
                OutputFileProceduresRg.Offset(0, 10) = "Comments"
                OutputFileProceduresRg.Offset(0, 11) = "Length Comments"
                                
                Set OutputFileProceduresRg = OutputFileProceduresRg.Offset(1, 0)
                
        Case Else: ws.Delete                                                                'Delete all other worksheet
        End Select
    Next i
    
    'Export all VBA code from Word / Excel files inside given folder
    Call ExportVBAFolder(myFolder, OutputFileRg, OutputFileProceduresRg, OutputFilews, OutputFileProceduresWs)
    
    
    'Format output file
    With OutputWb.Worksheets("File Procedures").Range("A2").CurrentRegion
        .HorizontalAlignment = xlGeneral                                                    'Align text left side
        .VerticalAlignment = xlTop                                                          'Align text to top of cell.
        .Columns.AutoFit                                                                    'Autofit all columns to content
        .Rows.AutoFit                                                                       'Autofit all rows to content
    End With
    OutputWb.Worksheets("File Procedures").Columns("J:J").ColumnWidth = 60.86               'Format column J to fixed size
    
    'Save output file
    OutputWb.SaveAs strFolder & "\VBA_Export_" & myFolder.Name & ".xlsx"
    
    
    'Enable Auto_Macros when document is openend
    'myWord.WordBasic.DisableAutoMacros 0                'Set security setting to 1 (High disables Auto_Macros)
    myExcel.EnableEvents = True                         'Suppres Events
    
    'Give user a feedback of the operation. In this case operation was succesfull.
    MsgBox "Succesfully exported all VBA from Word and Excel files in folder:" & vbCrLf & strFolder & vbCrLf & vbCrLf & _
            "Summary is given in file:= " & vbCrLf & _
            strFolder & "\VBA_Export_" & myFolder.Name & ".xlsx", vbOKOnly + vbExclamation
            
    'Procedure Cleanup
    myWord.Quit                                         'Close Word instant in memmory
    myExcel.Quit                                        'Close Excel instant in memmory
    Set myWord = Nothing                                'Close Word.Application instance
    Set myExcel = Nothing                               'Close Excel.Application instance
    Set myFolder = Nothing                              'Close myFolder
    Set OutputWb = Nothing                              'Close output file
    Set OutputRg = Nothing                              'Close output range
    Set OutputFilews = Nothing                          'Close output worksheet
    Set OutputFileProcedureWs = Nothing                 'Close output worksheet

    Exit Sub
ErrHndl:
    'Error Procedure Cleanup
    'Enable Auto_Macros when document is openend
    'myWord.WordBasic.DisableAutoMacros 0                'Set security setting to 1 (High disables Auto_Macros)
    myExcel.EnableEvents = True                         'Enable Events
    
    myWord.Quit                                         'Close Word instant in memmory
    myExcel.Quit                                        'Close Excel instant in memmory
    Set myWord = Nothing                                'Close Word.Application instance
    Set myExcel = Nothing                               'Close Excel.Application instance
    Set myFolder = Nothing                              'Close myFolder
    Set OutputWb = Nothing                              'Close output file
    Set OutputRg = Nothing                              'Close output range
    Set OutputFilews = Nothing                          'Close output worksheet
    Set OutputFileProcedureWs = Nothing                 'Close output worksheet

    'Give user feedback of the operation. In this case operation was unsuccesfull.
    MsgBox "Error during export procedure for folder:" & vbCrLf & strFolder, vbOKOnly + vbCritical
End Sub

Sub ExportVBAFolder(Folder As Folder, ByRef OutputFileRg As Excel.Range, ByRef OutputFileProceduresRg As Excel.Range, ByRef OutputFilews As Excel.Worksheet, ByRef OutputFileProceduresWs As Excel.Worksheet)
    Dim SubFolder As Folder                             'Microsoft Scripting Runtime
    Dim strFile As File                                 'Microsoft Scripting Runtime
    
    '=======================================================================
    'This procedure loops over the user given folder. The procedure calls itself _
    when an subfolder is found. When all folders are processed this way, then it _
    will loop over all files inside the found folders. It will try to export VBA _
    code when it finds the following files:
    'Word:      *.doc, *.dot, *.docm, *.dotm, *.dotx, *.docx
    'Excel:     *.xls, *.xlsm, *.xltm, *.xlt, *.xltx, *xlsm, *.xlam, *.xla, *.xlsx
    '=======================================================================

    'Loop over all subfolders. Procedure calls itself
    For Each SubFolder In Folder.SubFolders
        Call ExportVBAFolder(SubFolder, OutputFileRg, OutputFileProceduresRg, OutputFilews, OutputFileProceduresWs)
    Next SubFolder
    
    'Check each File inside Folder. Extract VBA code from Word / Excel from found files
    For Each strFile In Folder.Files
        'Debug.Print strFile.Type
        Select Case strFile.Type
        'Word Files
        Case "Microsoft Word 97 - 2003 Document":       Debug.Print "Case W1 " & strFile.Path & " " & strFile.Name                                              '*.doc
                                                        Call ExportCodeFile(strFile)                                                                            'Export VBA code
                                                        Call ListModule(strFile, OutputFileRg, OutputFileProceduresRg, OutputFilews, OutputFileProceduresWs)    'Export Summary
        Case "Microsoft Word 97 - 2003 Template":       Debug.Print "Case W2 " & strFile.Path & " " & strFile.Name                                              '*.dot
                                                        Call ExportCodeFile(strFile)                                                                            'Export VBA code
                                                        Call ListModule(strFile, OutputFileRg, OutputFileProceduresRg, OutputFilews, OutputFileProceduresWs)    'Export Summary
        Case "Microsoft Word Macro-Enabled Document":   Debug.Print "Case W3 " & strFile.Path & " " & strFile.Name                                              '*.docm
                                                        Call ExportCodeFile(strFile)                                                                            'Export VBA code
                                                        Call ListModule(strFile, OutputFileRg, OutputFileProceduresRg, OutputFilews, OutputFileProceduresWs)    'Export Summary
        Case "Microsoft Word Macro-Enabled Template":   Debug.Print "Case W4 " & strFile.Path & " " & strFile.Name                                              '*.dotm
                                                        Call ExportCodeFile(strFile)                                                                            'Export VBA code
                                                        Call ListModule(strFile, OutputFileRg, OutputFileProceduresRg, OutputFilews, OutputFileProceduresWs)    'Export Summary
        Case "Microsoft Word Template":                 Debug.Print "Case W5 " & strFile.Path & " " & strFile.Name                                              '*.dotx
                                                        Call ExportCodeFile(strFile)                                                                            'Export VBA code
                                                        Call ListModule(strFile, OutputFileRg, OutputFileProceduresRg, OutputFilews, OutputFileProceduresWs)    'Export Summary
        Case "Microsoft Word Document":                 Debug.Print "Case W6 " & strFile.Path & " " & strFile.Name                                              '*.docx
                                                        Call ExportCodeFile(strFile)                                                                            'Export VBA code
                                                        Call ListModule(strFile, OutputFileRg, OutputFileProceduresRg, OutputFilews, OutputFileProceduresWs)    'Export Summary
        'Excel Files
        Case "Microsoft Excel 97-2003 Worksheet":       Debug.Print "Case E1 " & strFile.Path & " " & strFile.Name                                              '*.xls
                                                        Call ExportCodeFile(strFile)                                                                            'Export VBA code
                                                        Call ListModule(strFile, OutputFileRg, OutputFileProceduresRg, OutputFilews, OutputFileProceduresWs)    'Export Summary
        Case "Microsoft Excel Macro-Enabled Worksheet": Debug.Print "Case E2 " & strFile.Path & " " & strFile.Name                                              '*.xlsm
                                                        Call ExportCodeFile(strFile)                                                                            'Export VBA code
                                                        Call ListModule(strFile, OutputFileRg, OutputFileProceduresRg, OutputFilews, OutputFileProceduresWs)    'Export Summary
        Case "Microsoft Excel Macro-Enabled Template":  Debug.Print "Case E3 " & strFile.Path & " " & strFile.Name                                              '*.xltm
                                                        Call ExportCodeFile(strFile)                                                                            'Export VBA code
                                                        Call ListModule(strFile, OutputFileRg, OutputFileProceduresRg, OutputFilews, OutputFileProceduresWs)    'Export Summary
        Case "Microsoft Excel Template":                Debug.Print "Case E4 " & strFile.Path & " " & strFile.Name                                              '*.xlt / *.xltx
                                                        Call ExportCodeFile(strFile)                                                                            'Export VBA code
                                                        Call ListModule(strFile, OutputFileRg, OutputFileProceduresRg, OutputFilews, OutputFileProceduresWs)    'Export Summary
        Case "Microsoft Excel Macro-Enabled Worksheet": Debug.Print "Case E5 " & strFile.Path & " " & strFile.Name                                              '*.xlsm
                                                        Call ExportCodeFile(strFile)                                                                            'Export VBA code
                                                        Call ListModule(strFile, OutputFileRg, OutputFileProceduresRg, OutputFilews, OutputFileProceduresWs)    'Export Summary
        Case "Microsoft Excel Add-In":                  Debug.Print "Case E6 " & strFile.Path & " " & strFile.Name                                              '*.xlam / *.xla
                                                        Call ExportCodeFile(strFile)                                                                            'Export VBA code
                                                        Call ListModule(strFile, OutputFileRg, OutputFileProceduresRg, OutputFilews, OutputFileProceduresWs)    'Export Summary
        Case "Microsoft Excel Worksheet":               Debug.Print "Case E7 " & strFile.Path & " " & strFile.Name                                              '*.xlsx
                                                        Call ExportCodeFile(strFile)                                                                            'Export VBA code
                                                        Call ListModule(strFile, OutputFileRg, OutputFileProceduresRg, OutputFilews, OutputFileProceduresWs)    'Export Summary
        Case Else: 'Do Nothing
        End Select
    Next strFile
End Sub

Sub ExportCodeFile(myFile As File)
Dim VBAEditor As VBIDE.VBE          'Microsoft VBScript Regular Expressions 5.5
Dim VBProj As VBIDE.VBProject       'Microsoft VBScript Regular Expressions 5.5
Dim VBComp As VBIDE.VBComponent     'Microsoft VBScript Regular Expressions 5.5
Dim CodeMod As VBIDE.CodeModule     'Microsoft VBScript Regular Expressions 5.5
Dim myWb As Excel.Workbook          'Microsoft Excel 14.0 Object Library
Dim myDoc As Document
Dim myDefaultEF As String
Dim myCodMod As VBComponent         'Microsoft Visual Basic Extensibility Version 5.3
Dim myExportStr As String

    '=======================================================================
    'This procedure loops over the user given folder. The procedure calls itself _
    when an subfolder is found. When all folders are processed this way, then it _
    will loop over all files inside the found folders. It will try to export VBA _
    code when it finds the following files:
    'Word:      *.doc, *.dot, *.docm, *.dotm, *.dotx, *.docx
    'Excel:     *.xls, *.xlsm, *.xltm, *.xlt, *.xltx, *xlsm, *.xlam, *.xla, *.xlsx
    '=======================================================================

    'Set Reference to VBProj
    Select Case myFile.Type
    Case "Microsoft Word 97 - 2003 Document", _
            "Microsoft Word 97 - 2003 Template", _
            "Microsoft Word Macro-Enabled Document", _
            "Microsoft Word Macro-Enabled Template", _
            "Microsoft Word Template", _
            "Microsoft Word Document":
                                                    Set myDoc = myWord.Documents.Open(myFile.Path)                              'Open Word file
                                                    Set VBProj = myDoc.VBProject                                                'Reference VBA Project of Word file
    Case "Microsoft Excel 97-2003 Worksheet", _
            "Microsoft Excel Macro-Enabled Worksheet", _
            "Microsoft Excel Macro-Enabled Worksheet", _
            "Microsoft Excel Add-In", _
            "Microsoft Excel Worksheet":
                                                    Set myWb = myExcel.Workbooks.Open(myFile.Path)                              'Open Excel file
                                                    Set VBProj = myWb.VBProject                                                 'Reference VBA Project of Excel file
                                                        
    Case "Microsoft Excel Macro-Enabled Template", _
            "Microsoft Excel Template":
                                                    Set myWb = myExcel.Workbooks.Open(myFile.Path, , , , , , , , , True)        'Open Excel Template file
                                                    Set VBProj = myWb.VBProject                                                 'Reference VBA Project of Excel file
                                                    
    Case Else:                                      Set VBProj = Nothing                                                        'Procedure doesn't support found filetype
                                                    MsgBox "Could not reference VBProj", vbCritical + vbOKOnly
    End Select

    'Check if the found file containts VBA code
    For i = 1 To VBProj.VBComponents.Count                                                                                      'Loop over all VBA Project Components
        Set VBComp = VBProj.VBComponents.Item(i)                                                                                'Set Reference to found VBA Component
        If (VBComp.CodeModule.CountOfDeclarationLines + 1) <= VBComp.CodeModule.CountOfLines Then                               'Check if VBA Component contains code
            Call CheckFolder(myFile)                                                                                            'Code is found, create a export folder
            Exit For
        End If
    Next i
   
    'Loop over all VBA Project Components
    For i = 1 To VBProj.VBComponents.Count
        Set myCodMod = VBProj.VBComponents.Item(i)                                                                              'Set Reference to found VBA Component
        
        Select Case myCodMod.Type                                                                                               'Check VBA Component Type
        Case 1:     'VBA Code Module
                    myExportStr = myFile.ParentFolder & "\" & myFile.Name & " VBA Code\" & myCodMod.Name & ".bas"               'Export path string *.bas
        Case 2:     'VBA Class Module
                    myExportStr = myFile.ParentFolder & "\" & myFile.Name & " VBA Code\" & myCodMod.Name & ".cls"               'Export path string *.cls
        Case 3:     'VBA Form Module
                    myExportStr = myFile.ParentFolder & "\" & myFile.Name & " VBA Code\" & myCodMod.Name & ".frm"               'Export path string *.frm
        Case 100:   'VBA Document Module
                    myExportStr = myFile.ParentFolder & "\" & myFile.Name & " VBA Code\" & myCodMod.Name & ".cls"               'Export path string *.cls
        Case Else:   'Procedure doesn't support found VBA Comonent Type
                     GoTo SkipExport
        End Select
        
        'Check if VBA Component contains code, if so export file with Export path string
        If myCodMod.CodeModule.CountOfLines > 0 Then myCodMod.Export (myExportStr)
    
SkipExport:
    Next i

ExitSub:
    'Procedure Cleanup
    Select Case myFile.Type
    Case "Microsoft Word 97 - 2003 Document", _
            "Microsoft Word 97 - 2003 Template", _
            "Microsoft Word Macro-Enabled Document", _
            "Microsoft Word Macro-Enabled Template", _
            "Microsoft Word Template", _
            "Microsoft Word Document":
                                                    myDoc.Close False                                                           'Close Word file
                                                    Set myDoc = Nothing                                                         'Release Word from memmory
                                                    Set VBProj = Nothing                                                        'Release VBProj from memmory
    Case "Microsoft Excel 97-2003 Worksheet", _
            "Microsoft Excel Macro-Enabled Worksheet", _
            "Microsoft Excel Macro-Enabled Worksheet", _
            "Microsoft Excel Macro-Enabled Template", _
            "Microsoft Excel Template", _
            "Microsoft Excel Add-In", _
            "Microsoft Excel Worksheet":
                                                    myWb.Close False                                                            'Close Excel file
                                                    Set myWb = Nothing                                                          'Release Excel file from memmory
                                                    Set VBProj = Nothing                                                        'Release VBProj from memmory
                                                        
    Case Else:                                      Set VBProj = Nothing                                                        'Release VBProj from memmory
    End Select
End Sub

Sub CheckFolder(myFile As File)
Dim myExportPath As String
Dim myDefaultEF As String
Dim FileSystem As New FileSystemObject              'Microsoft Scripting Runtime

'=======================================================================
'This procedure checks if for a given file if a folder exists where _
File.Name & " VBA Code\" exists. If it exists then the folder gets deleted _
then it gets created again. IF it doesn't exist then create the folder.
'=======================================================================

    myDefaultEF = "\" & myFile.Name & " VBA Code\"                                 'Default Export Folder Name
    myExportPath = myFile.ParentFolder & myDefaultEF                               'Export Folder path string
    
    If FileSystem.FolderExists(myExportPath) Then                                   'Check if a Folder exists in the Export Folder path string
        FileSystem.DeleteFolder Left(myExportPath, Len(myExportPath) - 1)           'Remove found Folder. Make sure that path string doesn't end with \
        FileSystem.CreateFolder myExportPath                                        'Create Folder in Export Folder path string location
    Else
        FileSystem.CreateFolder myExportPath                                        'Create Folder in Export Folder path string location
    End If
    
    'Procedure Cleanup
    Set FileSystem = Nothing                                                        'Release FileSystem from memmory
End Sub

Sub ListModule(myFile As File, ByRef OutputFileRg As Excel.Range, ByRef OutputFileProceduresRg As Excel.Range, ByRef OutputFilews As Excel.Worksheet, ByRef OutputFileProceduresWs As Excel.Worksheet)
Dim VBAEditor As VBIDE.VBE                  'Microsoft VBScript Regular Expressions 5.5
Dim VBProj As VBIDE.VBProject               'Microsoft VBScript Regular Expressions 5.5
Dim VBComp As VBIDE.VBComponent             'Microsoft VBScript Regular Expressions 5.5
Dim CodeMod As VBIDE.CodeModule             'Microsoft VBScript Regular Expressions 5.5
Dim myDoc As Document
Dim myWb As Workbook                        'Microsoft Excel 14.0 Object Library
Dim LineNum As Long
Dim ProcName As String
Dim ProcKind As VBIDE.vbext_ProcKind        'Microsoft VBScript Regular Expressions 5.5
Dim SlashPos As Long
Dim FileName As String
Dim myPInfo As ProcInfo

'=======================================================================
'This procedure extracts programming data per procedure from a VB Project CodeModule. _
The data is exported to an external workbook file.
'=======================================================================

    'Get the reference to VBProj of the given file
    Select Case myFile.Type
    Case "Microsoft Word 97 - 2003 Document", _
            "Microsoft Word 97 - 2003 Template", _
            "Microsoft Word Macro-Enabled Document", _
            "Microsoft Word Macro-Enabled Template", _
            "Microsoft Word Template", _
            "Microsoft Word Document":
                                                    Set myDoc = myWord.Documents.Open(myFile.Path)                              'Open Word File
                                                    Set VBProj = myDoc.VBProject                                                'Reference Word VB Project
    Case "Microsoft Excel 97-2003 Worksheet", _
            "Microsoft Excel Macro-Enabled Worksheet", _
            "Microsoft Excel Macro-Enabled Worksheet", _
            "Microsoft Excel Add-In", _
            "Microsoft Excel Worksheet":
                                                    Set myWb = myExcel.Workbooks.Open(myFile.Path)                              'Open Excel File
                                                    Set VBProj = myWb.VBProject                                                 'Reference Excel VB Project
                                                        
    Case "Microsoft Excel Macro-Enabled Template", _
            "Microsoft Excel Template":
                                                    Set myWb = myExcel.Workbooks.Open(myFile.Path, , , , , , , , , True)        'Open Excel Template File
                                                    Set VBProj = myWb.VBProject                                                 'Reference Excel VB Project
                                                    
    Case Else:                                      Set VBProj = Nothing                                                        'Procedure doesn't support found File.Type
                                                    MsgBox "Could not reference VBProj", vbCritical + vbOKOnly
    End Select
    
    'Output the found file to export workbook file (sheet 1)
    OutputFileRg.Offset(0, 0) = VBProj.BuildFileName                                                                            'Export Build File Name
    OutputFileRg.Offset(0, 1) = VBProj.FileName                                                                                 'Export File Name (full path name)
    SlashPos = InStrRev(VBProj.FileName, "\") + 1                                                                               'Determine last \ in File Name
    FileName = Mid(VBProj.FileName, SlashPos, Len(VBProj.FileName) - SlashPos + 1)                                              'Destil only the File Name path string
    OutputFileRg.Offset(0, 2) = FileName                                                                                        'Export File Name (only File Name)
    OutputFileRg.Offset(0, 3) = VBProj.Protection                                                                               'Export VB Project Protection setting
    OutputFileRg.Offset(0, 4) = VBProj.VBComponents.Count                                                                       'Export VB Project number of VB Project components
    Set OutputFileRg = OutputFileRg.Offset(1, 0)                                                                                'Set Output up for Next output File
    
    'Output the found file to export workbook file (sheet 2)
    For Each VBComp In VBProj.VBComponents                                                                                      'Loop over all VB Components in VB Project
        With VBComp.CodeModule                                                                                                  'For every VB Components CodeModule
            LineNum = .CountOfDeclarationLines + 1                                                                              'Set LineNumber
            Do While LineNum <= .CountOfLines                                                                                   'Loop while LineNumer is smaller then total amount of codelines
                OutputFileProceduresRg.Offset(0, 0) = FileName                                                                  'Export FileName
                OutputFileProceduresRg.Offset(0, 1) = VBComp.Name                                                               'Export VB Component Name
                OutputFileProceduresRg.Offset(0, 2) = ComponentTypeToString(VBComp.Type)                                        'Export string of VB Component Type
                ProcName = .ProcOfLine(LineNum, ProcKind)                                                                       'Get ProcName and ProcKind
                OutputFileProceduresRg.Offset(0, 3) = ProcName                                                                  'Export ProcName
                OutputFileProceduresRg.Offset(0, 4) = ProcKindString(ProcKind)                                                  'Export ProcKind
                myPInfo = ProcedureInfo(ProcName, ProcKind, VBComp.CodeModule)                                                  'Get ProcInfo
                With myPInfo
                    OutputFileProceduresRg.Offset(0, 5) = .ProcBodyLine                                                         'Export Proc Body Lines
                    OutputFileProceduresRg.Offset(0, 6) = .ProcCountLines                                                       'Export Proc Count Lines
                    OutputFileProceduresRg.Offset(0, 7) = .ProcStartLine                                                        'Export Proc Start Lines
                    OutputFileProceduresRg.Offset(0, 8) = ProcScopeString(.ProcScope)                                           'Export string of Proc Scope
                    OutputFileProceduresRg.Offset(0, 9) = .ProcDeclaration                                                      'Export Proc Declaration
                    OutputFileProceduresRg.Offset(0, 10) = .ProcComments                                                        'Export Proc Comments
                    OutputFileProceduresRg.Offset(0, 11) = Len(.ProcComments)                                                   'Export Proc Comments Length
                End With
                
                LineNum = .ProcStartLine(ProcName, ProcKind) + _
                            .ProcCountLines(ProcName, ProcKind) + 1                                                             'Get Next LineNumber : Next ProcName StartLine
                Set OutputFileProceduresRg = OutputFileProceduresRg.Offset(1, 0)                                                'Set Output up for Next output Procedure
            Loop
        End With
        'Set OutputRg = OutputRg.Offset(1, 0)
    Next VBComp
    
    'Procedure Cleanup
    Select Case myFile.Type
    Case "Microsoft Word 97 - 2003 Document", _
            "Microsoft Word 97 - 2003 Template", _
            "Microsoft Word Macro-Enabled Document", _
            "Microsoft Word Macro-Enabled Template", _
            "Microsoft Word Template", _
            "Microsoft Word Document":
                                                    myDoc.Close False                                                           'Close Word File
    Case "Microsoft Excel 97-2003 Worksheet", _
            "Microsoft Excel Macro-Enabled Worksheet", _
            "Microsoft Excel Macro-Enabled Template", _
            "Microsoft Excel Template", _
            "Microsoft Excel Macro-Enabled Worksheet", _
            "Microsoft Excel Add-In", _
            "Microsoft Excel Worksheet":
                                                    myWb.Close False                                                            'Close Excel File
                                                    
    Case Else:
    End Select
    
    Set myDoc = Nothing                                                                                                         'Release Word File from memmory
    Set myWb = Nothing                                                                                                          'Release Excel File from memmory
    Set VBProj = Nothing                                                                                                        'Release VB Project from memmory
    Set VBComp = Nothing                                                                                                        'Release VB Component from memmory
End Sub

Function ProcedureInfo(ProcName As String, ProcKind As VBIDE.vbext_ProcKind, CodeMod As VBIDE.CodeModule) As ProcInfo
Dim PInfo As ProcInfo
Dim BodyLine As Long
Dim Declaration As String
Dim FirstLine As String

'=======================================================================
'This procedure extracts declaration type of the given procedure from a VB Project CodeModule.
'=======================================================================

    BodyLine = CodeMod.ProcStartLine(ProcName, ProcKind)                                                                        'Get BodyLine from Code Module for given ProcName
    If BodyLine > 0 Then                                                                                                        'Check if BodyLine > 0
        With CodeMod
            PInfo.ProcName = ProcName                                                                                           'Set ProcName
            PInfo.ProcKind = ProcKind                                                                                           'Set ProcKind
            PInfo.ProcBodyLine = .ProcBodyLine(ProcName, ProcKind)                                                              'Set Proc Body Line
            PInfo.ProcCountLines = .ProcCountLines(ProcName, ProcKind)                                                          'Set Proc Count Lines
            PInfo.ProcStartLine = .ProcStartLine(ProcName, ProcKind)                                                            'Set Proc Start Line
            
            FirstLine = .Lines(PInfo.ProcBodyLine, 1)                                                                           'Get First Line of Code from Proc Body Lines
            'Check First Line of Code from Proc Body Lines, determine Proc Scope
            If StrComp(Left(FirstLine, Len("Public")), "Public", vbBinaryCompare) = 0 Then                                      'Check if Public
                PInfo.ProcScope = ScopePublic
            ElseIf StrComp(Left(FirstLine, Len("Private")), "Private", vbBinaryCompare) = 0 Then                                'Check if Private
                PInfo.ProcScope = ScopePrivate
            ElseIf StrComp(Left(FirstLine, Len("Friend")), "Friend", vbBinaryCompare) = 0 Then                                  'Check if Friend
                PInfo.ProcScope = ScopeFriend
            Else
                PInfo.ProcScope = ScopeDefault                                                                                  'Else it's Default
            End If
            PInfo.ProcDeclaration = GetProcedureDeclaration(CodeMod, ProcName, ProcKind, LineSplitRemove)                       'Get Proc Declaration
            PInfo.ProcComments = GetCommentLinesProcedure(CodeMod, ProcName, ProcKind, LineSplitRemove)                         'Get Proc Comments
        End With
    End If
    
    ProcedureInfo = PInfo                                                                                                       'Return ProcedureInfo
End Function

Public Function GetProcedureDeclaration(CodeMod As VBIDE.CodeModule, _
                    ProcName As String, ProcKind As VBIDE.vbext_ProcKind, _
                    Optional LineSplitBehavior As LineSplits = LineSplitRemove) As String
'====================================================================================
' GetProcedureDeclaration
' This return the procedure declaration of ProcName in CodeMod. The LineSplitBehavior
' determines what to do with procedure declaration that span more than one line using
' the "_" line continuation character. If LineSplitBehavior is LineSplitRemove, the
' entire procedure declaration is converted to a single line of text. If
' LineSplitBehavior is LineSplitKeep the "_" characters are retained and the
' declaration is split with vbNewLine into multiple lines. If LineSplitBehavior is
' LineSplitConvert, the "_" characters are removed and replaced with vbNewLine.
' The function returns vbNullString if the procedure could not be found.
'=====================================================================================
Dim LineNum As Long
Dim S As String
Dim Declaration As String

    On Error Resume Next
    LineNum = CodeMod.ProcBodyLine(ProcName, ProcKind)
    
    If Err.Number <> 0 Then
        Exit Function
    End If
    
    S = CodeMod.Lines(LineNum, 1)
    Do While Right(S, 1) = "_"
        Select Case True
            Case LineSplitBehavior = LineSplitConvert:  S = Left(S, Len(S) - 1) & vbNewLine
            Case LineSplitBehavior = LineSplitKeep:     S = S & vbNewLine
            Case LineSplitBehavior = LineSplitRemove:   S = Left(S, Len(S) - 1) & " "
        End Select
        Declaration = Declaration & S
        LineNum = LineNum + 1
        S = CodeMod.Lines(LineNum, 1)
    Loop
    Declaration = SingleSpace(Declaration & S)
    GetProcedureDeclaration = Declaration
End Function

Public Function GetCommentLinesProcedure(CodeMod As VBIDE.CodeModule, _
                    ProcName As String, ProcKind As VBIDE.vbext_ProcKind, _
                    Optional LineSplitBehavior As LineSplits = LineSplitRemove) As String
Dim LineNum As Long
Dim S As String
Dim strLine As String
Dim strComment As String
Dim bQoute As Long

'====================================================================================
' GetProcedureDeclaration
' This return the procedure declaration of ProcName in CodeMod. The LineSplitBehavior
' determines what to do with procedure declaration that span more than one line using
' the "_" line continuation character. If LineSplitBehavior is LineSplitRemove, the
' entire procedure declaration is converted to a single line of text. If
' LineSplitBehavior is LineSplitKeep the "_" characters are retained and the
' declaration is split with vbNewLine into multiple lines. If LineSplitBehavior is
' LineSplitConvert, the "_" characters are removed and replaced with vbNewLine.
' The function returns vbNullString if the procedure could not be found.
'=====================================================================================

    On Error Resume Next
    LineNum = CodeMod.ProcStartLine(ProcName, ProcKind)
    
    If Err.Number <> 0 Then
        Exit Function
    End If
    
    For i = LineNum To (LineNum + CodeMod.ProcCountLines(ProcName, ProcKind))
        S = CodeMod.Lines(i, 1)
        Do While Right(S, 1) = "_"
            Select Case True
                Case LineSplitBehavior = LineSplitConvert:  S = Left(S, Len(S) - 1) & vbNewLine
                Case LineSplitBehavior = LineSplitKeep:     S = S & vbNewLine
                Case LineSplitBehavior = LineSplitRemove:   S = Left(S, Len(S) - 1) & " "
            End Select
            strLine = strLine & S
            i = i + 1
            S = CodeMod.Lines(i, 1)
        Loop
        
        If strLine <> "" Then
            strLine = strLine & S
            strLine = SingleSpace(strLine)
        Else
            strLine = SingleSpace(S)
        End If
        
        bQoute = InStr(1, strLine, "'", vbBinaryCompare)
        Select Case bQoute
            Case 0:         'No Qoute found. Don't add line to strComment
            Case 1:         strComment = strComment & Mid(strLine, bQoute, Len(strLine) - bQoute + 1) & vbCrLf
            Case Is > 1:    strComment = strComment & "        " & Mid(strLine, bQoute, Len(strLine) - bQoute + 1) & vbCrLf     'vbTab doesn't work inside Excel cell, so simulate using 8 space characters
            Case Else:      'Do Nothing
        End Select
        strLine = ""
    Next i
    GetCommentLinesProcedure = strComment
End Function

Private Function SingleSpace(ByVal Text As String) As String
'====================================================================================
'This procedure removes all double spaces with single spaces from a given string.
'=====================================================================================
Dim Pos As String
    Pos = InStr(1, Text, Space(2), vbBinaryCompare)                                 'Determine position of double space
    Do Until Pos = 0                                                                'Loop untill no double space is found
        Text = Replace(Text, Space(2), Space(1))                                    'Replace double space with single space
        Pos = InStr(1, Text, Space(2), vbBinaryCompare)                             'Find next double space
    Loop
    SingleSpace = Text                                                              'Return replaced string
End Function

Function ComponentTypeToString(ComponentType As VBIDE.vbext_ComponentType) As String
'====================================================================================
'This procedure translates VBIDE.vbext_ComponentType to a string value for export purpose.
'=====================================================================================
    Select Case ComponentType
        Case vbext_ct_ActiveXDesigner:  ComponentTypeToString = "ActiveX Designer"
        Case vbext_ct_ClassModule:      ComponentTypeToString = "Class Module"
        Case vbext_ct_Document:         ComponentTypeToString = "Document Module"
        Case vbext_ct_MSForm:           ComponentTypeToString = "UserForm"
        Case vbext_ct_StdModule:        ComponentTypeToString = "Code Module"
        Case Else:                      ComponentTypeToString = "Unknown Type: " & CStr(ComponentType)
    End Select
End Function
Function ProcKindString(ProcKind As VBIDE.vbext_ProcKind) As String
'====================================================================================
'This procedure translates VBIDE.vbext_ProcKind to a string value for export purpose.
'=====================================================================================
    Select Case ProcKind
        Case vbext_pk_Get:  ProcKindString = "Property Get"
        Case vbext_pk_Let:  ProcKindString = "Property Let"
        Case vbext_pk_Set:  ProcKindString = "Property Set"
        Case vbext_pk_Proc: ProcKindString = "Sub Or Function"
        Case Else:          ProcKindString = "Unknown Type: " & CStr(ProcKind)
    End Select
End Function

Function ProcScopeString(ScopeString As ProcScope) As String
'====================================================================================
'This procedure translates ScopeString to a string value for export purpose.
'=====================================================================================
    Select Case ScopeString
        Case ScopePrivate:  ProcScopeString = "Private"
        Case ScopePublic:   ProcScopeString = "Public"
        Case ScopeFriend:   ProcScopeString = "Friend"
        Case ScopeDefault:  ProcScopeString = "Default"
        Case Else:          ProcScopeString = "Unknown Type: " & CStr(ScopeString)
    End Select
End Function

Function GetWordFolderContainingVBA() As String
Dim fd As FileDialog                                                            'Declare a variable as a FileDialog object.
Dim vrtSelectedItem As Variant                                                  'Declare a variable to contain the path of each selected item. _
                                                                                Even though the path is a String, the variable must be a Variant because _
                                                                                of the For Each ... Next routines only work with Variant and Objects.
'====================================================================================
'This procedure prompts the user with a msoFileDialogFolderPicker. Where the user can _
input a Folder. The Folder.path is given back as string.
'=====================================================================================
Set fd = Application.FileDialog(msoFileDialogFolderPicker)                      'Create a FileDialog object as a File Picker dialog.


With fd                                                                         'Use a With...End With block to reference the FileDialog object.

    'Add a filter that includes GIF and JPEG images and make it the first item in the list.
    '.Filters.Add "Word VBA Files", "*.doc;*.docx;*.docm", 1
    
    
    .AllowMultiSelect = False                                                   'Can not select more then 1 item

    'Use the Show method to display the File Picker dialog box and return the user's action.
    'If the user presses the action button...
    If .Show = -1 Then

        'Step through each string in the FileDialogSelectedItems collection.
        For Each vrtSelectedItem In .SelectedItems

            'vrtSelectedItem is a String that contains the path of each selected item.
            'You can use any file I/O functions that you want to work with this path.
            'This example simply displays the path in a message box.
            'MsgBox "Selected item's path: " & vrtSelectedItem
            GetWordFolderContainingVBA = vrtSelectedItem
            
        Next vrtSelectedItem
    'If the user presses Cancel...
    Else
    End If
End With

'Set the object variable to Nothing.
Set fd = Nothing
End Function

Function GetWordFileContainingVBA() As String
Dim fd As FileDialog                                                            'Declare a variable as a FileDialog object.
Dim vrtSelectedItem As Variant                                                  'Declare a variable to contain the path of each selected item. _
                                                                                Even though the path is a String, the variable must be a Variant because _
                                                                                of the For Each ... Next routines only work with Variant and Objects.
'====================================================================================
'This procedure prompts the user with a msoFileDialogFilePicker. Where the user can _
input a File. The File.path is given back as string.
'=====================================================================================
Set fd = Application.FileDialog(msoFileDialogFolderPicker)                      'Create a FileDialog object as a File Picker dialog.

With fd                                                                         'Use a With...End With block to reference the FileDialog object.

    'Add a filter that includes GIF and JPEG images and make it the first item in the list.
    '.Filters.Add "Word VBA Files", "*.doc;*.docx;*.docm", 1
        
    .AllowMultiSelect = False                                                   'Can not select more then 1 item

    'Use the Show method to display the File Picker dialog box and return the user's action.
    'If the user presses the action button...
    If .Show = -1 Then

        'Step through each string in the FileDialogSelectedItems collection.
        For Each vrtSelectedItem In .SelectedItems

            'vrtSelectedItem is a String that contains the path of each selected item.
            'You can use any file I/O functions that you want to work with this path.
            'This example simply displays the path in a message box.
            'MsgBox "Selected item's path: " & vrtSelectedItem
            GetWordFileContainingVBA = vrtSelectedItem
            
        Next vrtSelectedItem
    'If the user presses Cancel...
    Else
    End If
End With

'Set the object variable to Nothing.
Set fd = Nothing
End Function

Function IsEditorInSync() As Boolean
'=======================================================================
' IsEditorInSync
' This tests if the VBProject selected in the Project window, and
' therefore the ActiveVBProject is the same as the VBProject associated
' with the ActiveCodePane. If these two VBProjects are the same,
' the editor is in sync and the result is True. If these are not the
' same project, the editor is out of sync and the result is True.
'=======================================================================
    With Application.VBE
    IsEditorInSync = .ActiveVBProject Is _
        .ActiveCodePane.CodeModule.Parent.Collection.Parent
    End With
End Function

Sub SyncVBAEditor()
'=======================================================================
' SyncVBAEditor
' This syncs the editor with respect to the ActiveVBProject and the
' VBProject containing the ActiveCodePane. This makes the project
' that conrains the ActiveCodePane the ActiveVBProject.
'=======================================================================
    With Application.VBE
    If Not .ActiveCodePane Is Nothing Then
        Set .ActiveVBProject = .ActiveCodePane.CodeModule.Parent.Collection.Parent
    End If
    End With
End Sub

Sub EliminateScreenFlicker()
Dim VBEHwnd As Long
'=======================================================================
' EliminateScreenFlicker
' This disables the VBA Editor Screen Flickering
'=======================================================================
   
    On Error GoTo ErrH:
    
    Application.VBE.MainWindow.Visible = False
    
    VBEHwnd = FindWindow("wndclass_desked_gsk", _
        Application.VBE.MainWindow.Caption)
    
    If VBEHwnd Then
        LockWindowUpdate VBEHwnd
    End If
    
    '''''''''''''''''''''''''
    ' your code here
    '''''''''''''''''''''''''
    
    Application.VBE.MainWindow.Visible = False
ErrH:
    LockWindowUpdate 0&
End Sub
