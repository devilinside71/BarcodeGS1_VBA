VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GS1Class"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Reference needed: Microsoft VBScript Regula Expessions 5.5
Private m_strBarcode As String
Private Const cmstrBarcode As String = "01059965271763401020141719073121280122804"
Private m_strCatalognumber As String
Private Const cmstrCatalognumber As String = "280122804"
Private m_strEANnumber As String
Private Const cmstrEANnumber As String = "05996527176340"
Private m_strLOTnumber As String
Private Const cmstrLOTnumber As String = "2014"
Private m_strExpirationDate As String
Private Const cmstrExpirationDate As String = "190731"
Private Const cmstrGtinID As String = "01"
Private Const cmstrLotID As String = "10"
Private Const cmstrExpirationDateID As String = "17"
Private Const cmstrCatalogNumberID As String = "21"
Public Property Let Barcode(parBarcode As String)
    
    On Error GoTo PROC_ERR
    
    m_strBarcode = parBarcode
    Debug.Print "GS1Class.Barcode has been set to: " & m_strBarcode
    
PROC_EXIT:
    Exit Property
    
PROC_ERR:
    Err.Raise Err.Number
    Resume PROC_EXIT
End Property

Public Property Get Barcode() As String
    
    On Error GoTo PROC_ERR
    
    Barcode = m_strBarcode
    
PROC_EXIT:
    Exit Property
    
PROC_ERR:
    Err.Raise Err.Number
    Resume PROC_EXIT
End Property
Public Property Let Catalognumber(parCatalognumber As String)
    
    On Error GoTo PROC_ERR
    
    m_strCatalognumber = parCatalognumber
    Debug.Print "GS1Class.Catalognumber has been set to: " & m_strCatalognumber
    
PROC_EXIT:
    Exit Property
    
PROC_ERR:
    Err.Raise Err.Number
    Resume PROC_EXIT
End Property

Public Property Get Catalognumber() As String
    
    On Error GoTo PROC_ERR
    
    Catalognumber = m_strCatalognumber
    
PROC_EXIT:
    Exit Property
    
PROC_ERR:
    Err.Raise Err.Number
    Resume PROC_EXIT
End Property
Public Property Let EANnumber(parEANnumber As String)
    
    On Error GoTo PROC_ERR
    
    m_strEANnumber = parEANnumber
    Debug.Print "GS1Class.EANnumber has been set to: " & m_strEANnumber
    
PROC_EXIT:
    Exit Property
    
PROC_ERR:
    Err.Raise Err.Number
    Resume PROC_EXIT
End Property

Public Property Get EANnumber() As String
    
    On Error GoTo PROC_ERR
    
    EANnumber = m_strEANnumber
    
PROC_EXIT:
    Exit Property
    
PROC_ERR:
    Err.Raise Err.Number
    Resume PROC_EXIT
End Property
Public Property Let LOTnumber(parLOTnumber As String)
    
    On Error GoTo PROC_ERR
    
    m_strLOTnumber = parLOTnumber
    Debug.Print "GS1Class.LOTnumber has been set to: " & m_strLOTnumber
    
PROC_EXIT:
    Exit Property
    
PROC_ERR:
    Err.Raise Err.Number
    Resume PROC_EXIT
End Property

Public Property Get LOTnumber() As String
    
    On Error GoTo PROC_ERR
    
    LOTnumber = m_strLOTnumber
    
PROC_EXIT:
    Exit Property
    
PROC_ERR:
    Err.Raise Err.Number
    Resume PROC_EXIT
End Property
Public Property Let ExpirationDate(parExpirationDate As String)
    
    On Error GoTo PROC_ERR
    
    m_strExpirationDate = parExpirationDate
    Debug.Print "GS1Class.ExpirationDate has been set to: " & m_strExpirationDate
    
PROC_EXIT:
    Exit Property
    
PROC_ERR:
    Err.Raise Err.Number
    Resume PROC_EXIT
End Property

Public Property Get ExpirationDate() As String
    
    On Error GoTo PROC_ERR
    
    ExpirationDate = m_strExpirationDate
    
PROC_EXIT:
    Exit Property
    
PROC_ERR:
    Err.Raise Err.Number
    Resume PROC_EXIT
End Property
Private Sub Class_Initialize()
    Debug.Print "Class GS1Class initialized"
    
    m_strBarcode = cmstrBarcode
    Debug.Print "GS1Class Default value for Barcode: " & m_strBarcode
    m_strCatalognumber = cmstrCatalognumber
    Debug.Print "GS1Class Default value for Catalognumber: " & m_strCatalognumber
    m_strEANnumber = cmstrEANnumber
    Debug.Print "GS1Class Default value for EANnumber: " & m_strEANnumber
    m_strLOTnumber = cmstrLOTnumber
    Debug.Print "GS1Class Default value for LOTnumber: " & m_strLOTnumber
    m_strExpirationDate = cmstrExpirationDate
    Debug.Print "GS1Class Default value for ExpirationDate: " & m_strExpirationDate
End Sub
Private Sub Class_Terminate()
    Debug.Print "Class GS1Class terminated"
End Sub
Sub Reset()
    
    m_strBarcode = cmstrBarcode
    Debug.Print "GS1Class Default value for Barcode: " & m_strBarcode
    m_strCatalognumber = cmstrCatalognumber
    Debug.Print "GS1Class Default value for Catalognumber: " & m_strCatalognumber
    m_strEANnumber = cmstrEANnumber
    Debug.Print "GS1Class Default value for EANnumber: " & m_strEANnumber
    m_strLOTnumber = cmstrLOTnumber
    Debug.Print "GS1Class Default value for LOTnumber: " & m_strLOTnumber
    m_strExpirationDate = cmstrExpirationDate
    Debug.Print "GS1Class Default value for ExpirationDate: " & m_strExpirationDate
End Sub
Function CheckGTINID() As Boolean
    'Check GTIN ID
    'Returns: wether barcode starts with 01
    'Created by: Laszlo Tamas
    
    Dim blnRes As Boolean
    
    On Error GoTo FUNC_ERR
    
    blnRes = False
    'Code here
    If Left(m_strBarcode, 2) = "01" Then
        blnRes = True
    End If
    CheckGTINID = blnRes
    '---------------
FUNC_EXIT:
    On Error GoTo 0
    Exit Function
FUNC_ERR:
    Debug.Print "Error in Function CheckGTINID"
    If Err.Number Then
        Debug.Print Err.Description
    End If
    Resume FUNC_EXIT
End Function
Function FormatBarcode() As String
    'Format barcode
    'Returns: formatted barcode
    'Created by: Laszlo Tamas

    Dim strRes As String

    On Error GoTo FUNC_ERR

    strRes = m_strBarcode
    'Code here
    strRes = Replace(strRes, "(", "")
    strRes = Replace(strRes, ")", "")
    FormatBarcode = strRes
    '---------------
FUNC_EXIT:
    On Error GoTo 0
    Exit Function
FUNC_ERR:
    Debug.Print "Error in Function FormatBarcode"
    If Err.Number Then
        Debug.Print Err.Description
    End If
    Resume FUNC_EXIT
End Function
Function Verify() As Boolean
    'Description Verify GS1 barcode
    'Returns: wether the code is in 01..10..17..21.. format
    'Created by: Laszlo Tamas
    
    Dim blnRes As Boolean
    Dim rExp As New RegExp
    Dim rMatch As Object
    Dim rItem As Object
    
    On Error GoTo FUNC_ERR
    
    blnRes = False
    With rExp
        .Global = True
        .MultiLine = False
        .IgnoreCase = True
        .Pattern = "^(01)(\d{14})10(\d*)17(\d{6})21(\d{9})$"
    End With
    Set rMatch = rExp.Execute(m_strBarcode)
    'Debug.Print rMatch.Count
    If rMatch.Count > 0 Then
'        For Each rItem In rMatch
'            'code here
'        Next rItem
        blnRes = True
    End If
    
    Verify = blnRes
    '---------------
FUNC_EXIT:
    On Error GoTo 0
    Set rMatch = Nothing
    Exit Function
FUNC_ERR:
    Debug.Print "Error in Function Verify"
    If Err.Number Then
        Debug.Print Err.Description
    End If
    Resume FUNC_EXIT
End Function
Function GetEANnumber() As String
    'Get EAN number
    'Returns: EAN number
    'Created by: Laszlo Tamas
    
    Dim strRes As String
    Dim rExp As New RegExp
    Dim rMatch As Object
    Dim rItem As Object
    
    On Error GoTo FUNC_ERR
    
    strRes = m_strBarcode
    
    With rExp
        .Global = True
        .MultiLine = False
        .IgnoreCase = True
        .Pattern = "^(01)(\d{14})10(\d*)17(\d{6})21(\d{9})$"
    End With
    Set rMatch = rExp.Execute(m_strBarcode)
    If rMatch.Count > 0 Then
        strRes = rMatch.Item(0).SubMatches(1)
    End If
    
    
    GetEANnumber = strRes
    '---------------
FUNC_EXIT:
    On Error GoTo 0
    Exit Function
FUNC_ERR:
    Debug.Print "Error in Function GetEANnumber"
    If Err.Number Then
        Debug.Print Err.Description
    End If
    Resume FUNC_EXIT
End Function
Function GetLOTnumber() As String
    'Get LOT number
    'Parameters:
    'Returns:LOT number
    'Created by: Laszlo Tamas

    Dim strRes As String
    Dim rExp As New RegExp
    Dim rMatch As Object
    Dim rItem As Object
    
    On Error GoTo FUNC_ERR
    
    strRes = m_strBarcode
    
    With rExp
        .Global = True
        .MultiLine = False
        .IgnoreCase = True
        .Pattern = "^(01)(\d{14})10(\d*)17(\d{6})21(\d{9})$"
    End With
    Set rMatch = rExp.Execute(m_strBarcode)
    If rMatch.Count > 0 Then
        strRes = rMatch.Item(0).SubMatches(2)
    End If
    'Code here

    GetLOTnumber = strRes
    '---------------
FUNC_EXIT:
    On Error GoTo 0
    Exit Function
FUNC_ERR:
    Debug.Print "Error in Function GetLOTnumber"
    If Err.Number Then
        Debug.Print Err.Description
    End If
    Resume FUNC_EXIT
End Function







