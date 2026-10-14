Attribute VB_Name = "AsisFlowAccessExport"
' access_export.bas
'
' The same tool as access_export.vbs, for a PC where script files (.vbs) cannot be run.
' Writes the design of an Access file (.accdb / .mdb) out as text files. See access_export.vbs for what is written.
'
' How to use (on a PC that has Microsoft Access):
'   1. Make a copy of the Access file that is to be read.
'   2. Start Access and make a new, empty database (NOT the file that is to be read).
'   3. Open the Visual Basic editor (Alt+F11), choose File > Import File and pick this file.
'      (If you paste the text into a new module instead, leave out the first line "Attribute VB_Name ...".)
'   4. Put the cursor inside "AsisFlowAccessExport" at the bottom and press F5. Type the full path of the copy.
'   5. A folder "<file name>_export" is made next to the copy. Hand over the whole folder.
'
' The file that is read is never opened as the current database, so its start-up form, its AutoExec macro
' and its code do not run. This file is plain ASCII on purpose.

Option Compare Database
Option Explicit

'==== core begin (the same text is in access_export.vbs and access_export.bas) ====
Const TOOL_NAME = "asis-flow-skill access_export 1"
Const AC_FORM = 2
Const AC_REPORT = 3
Const AC_MACRO = 4
Const AC_MODULE = 5

Dim gFso
Dim gLog
Dim gOut
Dim gDone
Dim gFail

Function ExportAll(sSrc, sOutDir)
    Dim app, db, sTmp, sForms, sReports, sMacros, sModules
    ExportAll = False
    Set gFso = CreateObject("Scripting.FileSystemObject")
    If Not gFso.FileExists(sSrc) Then
        Say "File not found: " & sSrc
        Exit Function
    End If
    If Not gFso.FolderExists(sOutDir) Then gFso.CreateFolder sOutDir
    gOut = sOutDir
    gDone = 0
    gFail = 0
    Set gLog = gFso.CreateTextFile(gFso.BuildPath(gOut, "access_export.txt"), True, True)
    gLog.WriteLine "# access_export 1"
    gLog.WriteLine "tool" & vbTab & TOOL_NAME
    gLog.WriteLine "source" & vbTab & gFso.GetFileName(sSrc)
    gLog.WriteLine "date" & vbTab & Now
    sTmp = gFso.BuildPath(gOut, "_access_export_tmp.accdb")
    If gFso.FileExists(sTmp) Then gFso.DeleteFile sTmp, True

    On Error Resume Next
    Set app = CreateObject("Access.Application")
    If Err.Number <> 0 Then
        Fatal "Microsoft Access was not found on this PC (" & Err.Description & ")"
        Exit Function
    End If
    gLog.WriteLine "access" & vbTab & app.Version
    app.Visible = True
    Err.Clear
    app.NewCurrentDatabase sTmp
    If Err.Number <> 0 Then
        Fatal "Could not make the temporary database (" & Err.Number & " " & Err.Description & ")"
        app.Quit 2
        Exit Function
    End If
    Set db = app.DBEngine.OpenDatabase(sSrc, False, True)
    If Err.Number <> 0 Then
        Fatal "Could not open the file read-only. Files that need a password are not read (" & Err.Number & " " & Err.Description & ")"
        app.CloseCurrentDatabase
        app.Quit 2
        Exit Function
    End If
    On Error GoTo 0

    WriteTables db
    WriteQueries db
    sForms = DocNames(db, "Forms")
    sReports = DocNames(db, "Reports")
    sMacros = DocNames(db, "Scripts")
    sModules = DocNames(db, "Modules")
    db.Close
    Set db = Nothing

    SaveObjects app, sSrc, AC_FORM, sForms, "form", "F_", ".form.txt"
    SaveObjects app, sSrc, AC_REPORT, sReports, "report", "R_", ".report.txt"
    SaveObjects app, sSrc, AC_MACRO, sMacros, "macro", "X_", ".macro.txt"
    SaveObjects app, sSrc, AC_MODULE, sModules, "module", "M_", ".bas"

    On Error Resume Next
    app.CloseCurrentDatabase
    app.Quit 2
    Set app = Nothing
    If gFso.FileExists(sTmp) Then gFso.DeleteFile sTmp, True
    Err.Clear
    On Error GoTo 0
    gLog.WriteLine "done" & vbTab & gDone & vbTab & "failed" & vbTab & gFail
    gLog.Close
    ExportAll = True
    Say "Finished.  Written: " & gDone & "   Failed: " & gFail & vbCrLf & "Folder: " & gOut & vbCrLf & "Hand over the whole folder. The list is in access_export.txt."
End Function

Sub Fatal(sText)
    gLog.WriteLine "error" & vbTab & vbTab & vbTab & sText
    gLog.Close
    Say "Stopped: " & sText
End Sub

Function DocNames(db, sContainer)
    Dim doc, sText
    sText = ""
    On Error Resume Next
    For Each doc In db.Containers(sContainer).Documents
        If Left(doc.Name, 1) <> "~" Then
            If Len(sText) > 0 Then sText = sText & vbLf
            sText = sText & doc.Name
        End If
    Next
    Err.Clear
    On Error GoTo 0
    DocNames = sText
End Function

Sub SaveObjects(app, sSrc, nType, sNames, sKind, sPrefix, sSuffix)
    Dim aList, i, sName, sFile
    If Len(sNames) = 0 Then Exit Sub
    aList = Split(sNames, vbLf)
    For i = 0 To UBound(aList)
        sName = aList(i)
        sFile = sPrefix & SafeName(sName) & sSuffix
        On Error Resume Next
        Err.Clear
        app.DoCmd.TransferDatabase 0, "Microsoft Access", sSrc, nType, sName, sName, False
        If Err.Number = 0 Then app.SaveAsText nType, sName, gFso.BuildPath(gOut, sFile)
        If Err.Number = 0 Then
            gDone = gDone + 1
            gLog.WriteLine sKind & vbTab & sName & vbTab & sFile & vbTab & "ok"
        Else
            gFail = gFail + 1
            gLog.WriteLine sKind & vbTab & sName & vbTab & sFile & vbTab & "FAILED " & Err.Number & " " & OneLine(Err.Description)
            Err.Clear
        End If
        On Error GoTo 0
    Next
End Sub

Sub WriteTables(db)
    Dim f, td, fld, rel, i, sLine, sPk, n
    Set f = gFso.CreateTextFile(gFso.BuildPath(gOut, "access_tables.txt"), True, True)
    f.WriteLine "# access_tables 1"
    f.WriteLine "# Table definitions only. No rows are written."
    n = 0
    On Error Resume Next
    sLine = ""
    sLine = db.Properties("StartupForm").Value
    If Err.Number <> 0 Then
        sLine = ""
        Err.Clear
    End If
    If Len(sLine) > 0 Then f.WriteLine "[startup]" & vbTab & "form=" & sLine
    For Each td In db.TableDefs
        If Not IsSystemName(td.Name) Then
            If Len(td.Connect) > 0 Then
                f.WriteLine "[linked]" & vbTab & td.Name & vbTab & "source=" & td.SourceTableName & vbTab & "connect=" & MaskSecret(td.Connect)
            Else
                f.WriteLine "[table]" & vbTab & td.Name & vbTab & "rows=" & td.RecordCount
            End If
            sPk = PrimaryKey(td)
            For Each fld In td.Fields
                sLine = "  " & fld.Name & vbTab & FieldType(fld)
                If InStr(sPk, vbLf & fld.Name & vbLf) > 0 Then sLine = sLine & vbTab & "pk"
                f.WriteLine sLine
            Next
            n = n + 1
            Err.Clear
        End If
    Next
    For Each rel In db.Relations
        If Not IsSystemName(rel.Table) Then
            For i = 0 To rel.Fields.Count - 1
                f.WriteLine "[relation]" & vbTab & rel.Table & "." & rel.Fields(i).Name & vbTab & rel.ForeignTable & "." & rel.Fields(i).ForeignName
            Next
        End If
    Next
    Err.Clear
    On Error GoTo 0
    f.Close
    gLog.WriteLine "tables" & vbTab & n & vbTab & "access_tables.txt" & vbTab & "ok"
End Sub

Sub WriteQueries(db)
    Dim f, qd, n
    Set f = gFso.CreateTextFile(gFso.BuildPath(gOut, "access_queries.sql"), True, True)
    f.WriteLine "-- access_queries 1"
    n = 0
    On Error Resume Next
    For Each qd In db.QueryDefs
        If Left(qd.Name, 1) <> "~" Then
            f.WriteLine ""
            f.WriteLine "-- query: " & qd.Name & vbTab & "type=" & QueryType(qd.Type)
            If Len(qd.Connect) > 0 Then f.WriteLine "-- connect: " & MaskSecret(qd.Connect)
            f.WriteLine MaskSecret(qd.SQL)
            n = n + 1
            Err.Clear
        End If
    Next
    Err.Clear
    On Error GoTo 0
    f.Close
    gLog.WriteLine "queries" & vbTab & n & vbTab & "access_queries.sql" & vbTab & "ok"
End Sub

Function PrimaryKey(td)
    Dim idx, fld, sText
    sText = vbLf
    On Error Resume Next
    For Each idx In td.Indexes
        If idx.Primary Then
            For Each fld In idx.Fields
                sText = sText & fld.Name & vbLf
            Next
        End If
    Next
    Err.Clear
    On Error GoTo 0
    PrimaryKey = sText
End Function

Function FieldType(fld)
    Dim sText
    Select Case fld.Type
        Case 1
            sText = "yesno"
        Case 2
            sText = "byte"
        Case 3
            sText = "integer"
        Case 4
            sText = "long"
            If (fld.Attributes And 16) <> 0 Then sText = "autonumber"
        Case 5
            sText = "currency"
        Case 6
            sText = "single"
        Case 7
            sText = "double"
        Case 8
            sText = "datetime"
        Case 9
            sText = "binary"
        Case 10
            sText = "text(" & fld.Size & ")"
        Case 11
            sText = "ole"
        Case 12
            sText = "memo"
        Case 15
            sText = "guid"
        Case 16
            sText = "bigint"
        Case 20
            sText = "decimal"
        Case 101
            sText = "attachment"
        Case Else
            sText = "type" & fld.Type
    End Select
    FieldType = sText
End Function

Function QueryType(nType)
    Dim sText
    Select Case nType
        Case 0
            sText = "select"
        Case 16
            sText = "crosstab"
        Case 32
            sText = "delete"
        Case 48
            sText = "update"
        Case 64
            sText = "append"
        Case 80
            sText = "maketable"
        Case 96
            sText = "ddl"
        Case 112
            sText = "passthrough"
        Case 128
            sText = "union"
        Case Else
            sText = "type" & nType
    End Select
    QueryType = sText
End Function

Function IsSystemName(sName)
    Dim sUp
    sUp = UCase(sName)
    IsSystemName = (Left(sUp, 4) = "MSYS") Or (Left(sUp, 4) = "USYS") Or (Left(sUp, 1) = "~")
End Function

Function SafeName(sName)
    Dim re
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.Pattern = "[\\/:*?""<>|]"
    SafeName = re.Replace(sName, "_")
End Function

Function MaskSecret(sText)
    Dim re
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.IgnoreCase = True
    re.Pattern = "(PWD|PASSWORD)\s*=\s*[^;]*"
    MaskSecret = re.Replace(sText, "$1=***")
End Function

Function OneLine(sText)
    OneLine = Replace(Replace(sText, vbCr, " "), vbLf, " ")
End Function
'==== core end ====

Sub Say(sText)
    MsgBox sText, 0, "access_export"
End Sub

Public Sub AsisFlowAccessExport()
    Dim fso, sSrc, sOutDir
    sSrc = InputBox("Full path of the Access file to write out as text (.accdb / .mdb)." & vbCrLf & "Use a copy of the file.", "access_export")
    sSrc = Trim(Replace(sSrc, Chr(34), ""))
    If Len(sSrc) = 0 Then Exit Sub
    Set fso = CreateObject("Scripting.FileSystemObject")
    sOutDir = fso.BuildPath(fso.GetParentFolderName(sSrc), fso.GetBaseName(sSrc) & "_export")
    ExportAll sSrc, sOutDir
End Sub
