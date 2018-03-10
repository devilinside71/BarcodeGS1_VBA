Attribute VB_Name = "Module1"
Option Explicit
Sub GS1()
    'Description
    'Parameters:
    'Created by: Laszlo Tamas


    On Error GoTo PROC_ERR

   'Code here

    '---------------
PROC_EXIT:
    On Error GoTo 0
    Exit Sub
PROC_ERR:
    Debug.Print "Error in Procedure GS1"
    If Err.Number Then
        Debug.Print Err.Description
    End If
    Resume PROC_EXIT
End Sub
Private Sub GS1Test()
    'Test procedure for GS1
    Dim dtmStartTime As Date



    dtmStartTime = Now()
    Call GS1
End Sub
'----------------
'Columns and Rows
'----------------
Private Function Col_Letter(lngCol As Long) As String
    'Get letter from column number
    Dim vArr
    
    '  On Error Resume Next
    vArr = Split(Cells(1, lngCol).Address(True, False), "$")
    Col_Letter = vArr(0)
End Function
Private Function Col_LetterHeader(sheetName As String, headText As String, Optional headRow = 1) As String
    'Get column letter from header text
    Dim lngColNumber As Long
    
    lngColNumber = Col_NumberHeader(sheetName, headText, headRow)
    Col_LetterHeader = Col_Letter(lngColNumber)
End Function
Private Function Col_Number(colLetter) As Long
    'Get column number from column letter
    Col_Number = Range(colLetter & "1").Column
End Function
Private Function Col_NumberHeader(sheetName As String, headText As String, Optional headRow = 1) As Long
    'Get column number from header text
    Dim i As Long
    Dim strCellString As String
    
    Col_NumberHeader = 0
    For i = 1 To 400
        strCellString = Trim(CStr(Sheets(sheetName).Cells(headRow, i)))
        If strCellString = headText Then
            Col_NumberHeader = i
            Exit Function
        End If
    Next i
End Function
Private Sub ColLetterTests()
    'Test for Col_Letter, Col_LetterHeader, Col_Number and Col_NumberHeader
    Debug.Print Col_Letter(12)
    Debug.Print Col_LetterHeader("Hogyallunk", "Any.csop.")
    Debug.Print Col_Number("H")
    Debug.Print Col_NumberHeader("Hogyallunk", "Any.csop.")
End Sub
Private Function GetLastRow(sheetName As String, checkColumn As Long, _
    Optional firstrow = 2, Optional lastrow = 600000, _
        Optional backwardCheck = True) As Long
    'Adott f�l utols� sora
    Dim i As Long
    Dim curSheet As Worksheet
    Dim strCell As String
    
    Set curSheet = ActiveWorkbook.ActiveSheet
    Sheets(sheetName).Activate
    GetLastRow = 0
    If backwardCheck Then
        For i = lastrow To firstrow Step -1
            strCell = Trim(CStr(Cells(i, checkColumn)))
            If strCell <> "" Then
                GetLastRow = i
                Exit For
            End If
        Next i
    Else
        For i = firstrow To lastrow
            strCell = Trim(CStr(Cells(i, checkColumn)))
            If strCell = "" Then
                GetLastRow = i - 1
                Exit For
            End If
        Next i
    End If
    curSheet.Activate
    Set curSheet = Nothing
    Debug.Print "LastRow of " & sheetName & ": " & GetLastRow & " ChkCol:" & checkColumn
End Function


Private Sub GS1Class_ClassTest()
    Dim clGS1Class As New GS1Class
    
    clGS1Class.Barcode = "01059965271763401020141719073121280122804"
    Debug.Print "clGS1Class.Barcode: " & clGS1Class.Barcode
    clGS1Class.Catalognumber = "280122804"
    Debug.Print "clGS1Class.Catalognumber: " & clGS1Class.Catalognumber
    clGS1Class.EANnumber = "05996527176340"
    Debug.Print "clGS1Class.EANnumber: " & clGS1Class.EANnumber
    clGS1Class.LOTnumber = "2014"
    Debug.Print "clGS1Class.LOTnumber: " & clGS1Class.LOTnumber
    clGS1Class.ExpirationDate = "190731"
    Debug.Print "clGS1Class.ExpirationDate: " & clGS1Class.ExpirationDate
    
    clGS1Class.Barcode = "01059965271763401020141719073121280122804"
    Debug.Print "Function CheckGTINID test: >> " & clGS1Class.CheckGTINID()
    
    clGS1Class.Barcode = "(01)05996527176340(10)2014(17)190731(21)280122804"
    Debug.Print "Function FormatBarcode test: >> " & clGS1Class.FormatBarcode()
    
    clGS1Class.Barcode = "01059965271763401020141719073121280122804"
    Debug.Print "Function Verify test: >> " & clGS1Class.Verify()
    
    clGS1Class.Barcode = "01059965271763401020141719073121280122804"
    Debug.Print "Function GetEANnumber test: >> " & clGS1Class.GetEANnumber()
    
    clGS1Class.Barcode = "01059965271763401020141719073121280122804"
    Debug.Print "Function GetLOTnumber test: >> " & clGS1Class.GetLOTnumber()
    
    Set clGS1Class = Nothing
End Sub




