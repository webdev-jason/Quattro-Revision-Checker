Attribute VB_Name = "Module1"
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" ( _
        ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
    Private Declare PtrSafe Function FindWindowEx Lib "user32" Alias "FindWindowExA" ( _
        ByVal hWndParent As LongPtr, ByVal hWndChildAfter As LongPtr, _
        ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
    Private Declare PtrSafe Function GetWindowText Lib "user32" Alias "GetWindowTextA" ( _
        ByVal hWnd As LongPtr, ByVal lpString As String, ByVal cch As Long) As Long
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" ( _
        ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" ( _
        ByVal hWndParent As Long, ByVal hWndChildAfter As Long, _
        ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" ( _
        ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

Private Const LOAD_WAIT_SECONDS As Long = 6
Private Const MAX_REV_CHARS As Long = 4

'=============================================================
' MAIN PROCEDURE
'=============================================================
Sub ExtractChromeTabRev()
    Dim ws As Worksheet
    Dim cell As Range, r As Variant
    Dim windowTitle As String, revText As String
    Dim rangesToCheck As Variant
    Dim shp As Shape, h As Hyperlink
    Dim rngBorders As Range, rngSpec As Variant
    Dim savePath As String, fileName As String
    Dim tempCols As Variant
    Dim rCheck As Variant, c As Range
    Dim newCell As Range
    Dim oldVal As String, newVal As String
    Dim i As Long

    Set ws = ActiveSheet
    rangesToCheck = Array("A3:A32", "G7:G8", "G12:G22")

    ' --- TEMP COLUMNS (one per range) to avoid overwriting when rows overlap ---
    tempCols = Array(ws.Columns("Z").Column, ws.Columns("AA").Column, ws.Columns("AB").Column)

    ' --- BACKUP EXISTING REV VALUES (one temp column per range) ---
    For i = LBound(rangesToCheck) To UBound(rangesToCheck)
        For Each c In ws.Range(rangesToCheck(i))
            ws.Cells(c.Row, tempCols(i)).Value = c.Offset(0, 1).Value
        Next c
    Next i

    ' --- CLEAR PREVIOUS RESULTS (target cells where revs go) ---
    For Each r In rangesToCheck
        ws.Range(r).Offset(0, 1).ClearContents
    Next r

    ' --- PROCESS LINKS: open and extract revs ---
    For Each r In rangesToCheck
        For Each cell In ws.Range(r)
            If cell.Hyperlinks.Count > 0 Then
                cell.Hyperlinks(1).Follow NewWindow:=True
                Sleep LOAD_WAIT_SECONDS * 1000

                windowTitle = GetChromeWindowTitle()
                If Len(windowTitle) > 0 Then
                    revText = ExtractRevFromTitle(windowTitle)
                    If Len(revText) > 0 Then
                        cell.Offset(0, 1).Value = revText
                    Else
                        cell.Offset(0, 1).Value = ""
                    End If
                Else
                    cell.Offset(0, 1).Value = ""
                End If
            End If
        Next cell
    Next r

    ' --- COMPARE OLD AND NEW VALUES; HIGHLIGHT ONLY TRUE CHANGES ---
    For i = LBound(rangesToCheck) To UBound(rangesToCheck)
        For Each newCell In ws.Range(rangesToCheck(i))
            oldVal = CleanForCompare(ws.Cells(newCell.Row, tempCols(i)).Value)
            newVal = CleanForCompare(newCell.Offset(0, 1).Value)

            ' Treat placeholder tokens as empty (in case they show up)
            If oldVal = "REVNOTFOUND" Or oldVal = "NOWINDOW" Then oldVal = ""
            If newVal = "REVNOTFOUND" Or newVal = "NOWINDOW" Then newVal = ""

            ' Now compare: highlight only if different and at least one is non-empty
            If oldVal <> newVal Then
                If Not (oldVal = "" And newVal = "") Then
                    newCell.Offset(0, 1).Interior.Color = vbYellow
                End If
            End If
        Next newCell
    Next i

    ' --- REMOVE TEMPORARY BACKUPS (clear only the rows we used) ---
    For i = LBound(rangesToCheck) To UBound(rangesToCheck)
        For Each c In ws.Range(rangesToCheck(i))
            ws.Cells(c.Row, tempCols(i)).ClearContents
        Next c
    Next i

    ' --- CLEANUP BEFORE SAVE ---
    ' 1. Remove all hyperlinks but preserve visible text
    For Each h In ws.Hyperlinks
        h.Range.Value = h.Range.Value
    Next h
    ws.Hyperlinks.Delete

    ' 1a. Reapply "All Borders" to the specified ranges
    For Each rngSpec In Array("A3:A32", "G7:G8", "G12:G22")
        Set rngBorders = ws.Range(rngSpec)
        With rngBorders.Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With
    Next rngSpec

    ' 2. Unmerge and clear G30 (unmerge BEFORE clear)
    With ws.Range("G30")
        If .MergeCells Then .UnMerge
        .ClearContents
    End With

    ' 3. Delete form control button(s)
    For Each shp In ws.Shapes
        If shp.Type = msoFormControl Then
            If shp.FormControlType = xlButtonControl Then shp.Delete
        End If
    Next shp

    ' --- RECORD DATE AND TIME ---
    ws.Range("K34").Value = Date
    ws.Range("K35").Value = Time

    ' --- SAVE MACRO-FREE COPY (.xlsx) ---
    fileName = "Quattro Revisions.xlsx"
    savePath = ThisWorkbook.Path & "\" & fileName

    Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs fileName:=savePath, FileFormat:=xlOpenXMLWorkbook
    Application.DisplayAlerts = True

    MsgBox "Revisions updated and compared successfully!" & vbCrLf & _
           "Changed cells highlighted." & vbCrLf & _
           "Macro-free copy saved as: " & fileName, vbInformation
End Sub

'=============================================================
' CleanForCompare:
' - removes invisible chars (NBSP, zero-width, BOM)
' - strips CR/LF/tabs
' - trims & collapses spaces
' - returns ONLY alphanumeric characters in UPPERCASE
'=============================================================
Private Function CleanForCompare(ByVal v As Variant) As String
    Dim s As String, outS As String
    Dim i As Long, ch As String

    On Error GoTo CleanFail
    If IsError(v) Or IsNull(v) Then CleanForCompare = "": Exit Function

    s = CStr(v)
    ' Remove common invisible characters
    s = Replace(s, Chr(160), " ")         ' non-breaking space
    s = Replace(s, ChrW(8203), "")        ' zero-width space
    s = Replace(s, ChrW(65279), "")       ' BOM/ZWNBSP
    s = Replace(s, vbCr, "")
    s = Replace(s, vbLf, "")
    s = Replace(s, vbTab, " ")
    s = Trim$(s)

    On Error Resume Next
    s = Application.WorksheetFunction.Trim(s) ' collapse multiple internal spaces
    On Error GoTo CleanFail

    s = UCase$(s)

    ' Build output with only A-Z and 0-9
    outS = ""
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch Like "[A-Z0-9]" Then outS = outS & ch
    Next i

    CleanForCompare = outS
    Exit Function
CleanFail:
    CleanForCompare = ""
End Function

'=============================================================
' Find any Chrome/Edge (or other) window containing "Rev"
'=============================================================
Private Function GetChromeWindowTitle() As String
    Dim hWnd As LongPtr
    Dim sBuffer As String * 255
    Dim lLen As Long
    Dim classNames As Variant
    Dim cls As Variant
    Dim title As String

    classNames = Array("Chrome_WidgetWin_1", "Edge_WidgetWin_1", "MozillaWindowClass", "IEFrame")

    For Each cls In classNames
        hWnd = FindWindow(cls, vbNullString)
        Do While hWnd <> 0
            lLen = GetWindowText(hWnd, sBuffer, 255)
            If lLen > 0 Then
                title = Left$(sBuffer, lLen)
                If InStr(1, title, "Rev", vbTextCompare) > 0 Then
                    GetChromeWindowTitle = title
                    Exit Function
                End If
            End If
            hWnd = FindWindowEx(0, hWnd, cls, vbNullString)
        Loop
    Next cls
End Function

'=============================================================
' Extract revision letters/numbers following "Rev"
'=============================================================
Private Function ExtractRevFromTitle(ByVal title As String) As String
    Dim pos As Long, revPart As String, i As Long
    pos = InStr(1, title, "Rev", vbTextCompare)
    If pos = 0 Then Exit Function

    revPart = Mid$(title, pos + 3)
    revPart = Replace(revPart, ":", "")
    revPart = Trim$(revPart)

    For i = 1 To Len(revPart)
        If Mid$(revPart, i, 1) Like "[A-Za-z0-9]" Then
            ExtractRevFromTitle = ExtractRevFromTitle & Mid$(revPart, i, 1)
            If Len(ExtractRevFromTitle) >= MAX_REV_CHARS Then Exit For
        ElseIf Len(ExtractRevFromTitle) > 0 Then
            Exit For
        End If
    Next i
End Function


