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

Private Const LOAD_WAIT_SECONDS As Long = 5     ' Time to wait for the browser to load
Private Const MAX_REV_CHARS As Long = 4         ' Max characters to capture after "Rev"

'=============================================================
' MAIN PROCEDURE
'=============================================================
Sub ExtractChromeTabRev()
    Dim ws As Worksheet
    Dim cell As Range
    Dim windowTitle As String
    Dim revText As String
    Dim r As Variant
    Dim rangesToCheck As Variant
    Dim shp As Shape
    Dim savePath As String
    Dim fileName As String
    Dim h As Hyperlink
    Dim rngBorders As Range
    Dim rngSpec As Variant
    Dim borderIndex As Long

    Set ws = ActiveSheet
    rangesToCheck = Array("A3:A32", "G7:G8", "G12:G22")

    ' --- CLEAR PREVIOUS RESULTS ---
    For Each r In rangesToCheck
        ws.Range(r).Offset(0, 1).ClearContents
    Next r

    ' --- PROCESS LINKS ---
    For Each r In rangesToCheck
        For Each cell In ws.Range(r)
            If cell.Hyperlinks.Count > 0 Then
                ' Open hyperlink
                cell.Hyperlinks(1).Follow NewWindow:=True
                Sleep LOAD_WAIT_SECONDS * 1000

                ' Find Chrome/Edge window containing "Rev"
                windowTitle = GetChromeWindowTitle()

                If Len(windowTitle) > 0 Then
                    revText = ExtractRevFromTitle(windowTitle)
                    If Len(revText) > 0 Then
                        cell.Offset(0, 1).Value = revText
                    Else
                        cell.Offset(0, 1).Value = "RevNotFound"
                    End If
                Else
                    cell.Offset(0, 1).Value = "NoWindow"
                End If
            End If
        Next cell
    Next r

    ' --- CLEANUP BEFORE SAVE ---
    ' 1. Remove all hyperlinks but preserve displayed text and formatting as much as possible
    For Each h In ws.Hyperlinks
        With h.Range
            .Value = .Value  ' keeps the visible text (removes hyperlink)
        End With
    Next h
    ' Remove hyperlink objects (no visual change to values)
    ws.Hyperlinks.Delete

    ' 1a. Reapply "All Borders" to the specified ranges to ensure borders persist
    For Each rngSpec In Array("A3:A32", "G7:G8", "G12:G22")
        Set rngBorders = ws.Range(rngSpec)
        For borderIndex = xlEdgeLeft To xlInsideHorizontal
            With rngBorders.Borders(borderIndex)
                .LineStyle = xlContinuous
                .Weight = xlThin
            End With
        Next borderIndex
    Next rngSpec

    ' 2. Unmerge and clear G30 (unmerge BEFORE clearing)
    With ws.Range("G30")
        If .MergeCells Then .UnMerge
        .ClearContents
    End With

    ' 3. Delete form control button(s) (Form Controls only)
    For Each shp In ws.Shapes
        If shp.Type = msoFormControl Then
            If shp.FormControlType = xlButtonControl Then
                shp.Delete
            End If
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

    MsgBox "All done extracting revisions!" & vbCrLf & _
           "Date recorded in K34." & vbCrLf & _
           "Macro-free copy saved as: " & fileName, vbInformation
End Sub

'=============================================================
' Find any Chrome/Edge window containing "Rev"
'=============================================================
Private Function GetChromeWindowTitle() As String
    Dim hWnd As LongPtr
    Dim sBuffer As String * 255
    Dim lLen As Long
    Dim classNames As Variant
    Dim cls As Variant

    classNames = Array("Chrome_WidgetWin_1", "Edge_WidgetWin_1", "MozillaWindowClass", "IEFrame")

    For Each cls In classNames
        hWnd = FindWindow(cls, vbNullString)
        Do While hWnd <> 0
            lLen = GetWindowText(hWnd, sBuffer, 255)
            If lLen > 0 Then
                Dim title As String
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
    Dim pos As Long
    Dim revPart As String
    Dim i As Long

    pos = InStr(1, title, "Rev", vbTextCompare)
    If pos = 0 Then Exit Function

    revPart = Mid$(title, pos + 3)
    revPart = Replace(revPart, ":", "")
    revPart = Trim$(revPart)

    ' Capture up to MAX_REV_CHARS or until space or dash
    For i = 1 To Len(revPart)
        If Mid$(revPart, i, 1) Like "[A-Za-z0-9]" Then
            ExtractRevFromTitle = ExtractRevFromTitle & Mid$(revPart, i, 1)
            If Len(ExtractRevFromTitle) >= MAX_REV_CHARS Then Exit For
        ElseIf Len(ExtractRevFromTitle) > 0 Then
            Exit For
        End If
    Next i
End Function


