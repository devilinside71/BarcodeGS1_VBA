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
Private mDictC As Collection

Public Property Let Barcode(parBarcode As String)
    
    On Error GoTo PROC_ERR
    
    m_strBarcode = parBarcode
    If InStr(m_strBarcode, "(") > 0 Then
        m_strBarcode = FormatBarcode()
    End If
    Debug.Print "GS1Class.Barcode has been set to: " & m_strBarcode
    
PROC_EXIT:
    Exit Property
    
PROC_ERR:
    Err.Raise Err.Number
    Resume PROC_EXIT
End Property

Public Property Get Barcode() As String
    
    On Error GoTo PROC_ERR
    
    Barcode = FormatBarcode()
    
PROC_EXIT:
    Exit Property
    
PROC_ERR:
    Err.Raise Err.Number
    Resume PROC_EXIT
End Property
Public Property Let CatalogNumber(parCatalognumber As String)
    
    On Error GoTo PROC_ERR
    
    m_strCatalognumber = parCatalognumber
    Debug.Print "GS1Class.Catalognumber has been set to: " & m_strCatalognumber
    
PROC_EXIT:
    Exit Property
    
PROC_ERR:
    Err.Raise Err.Number
    Resume PROC_EXIT
End Property

Public Property Get CatalogNumber() As String
    
    On Error GoTo PROC_ERR
    m_strCatalognumber = GetCatalogNumber
    CatalogNumber = m_strCatalognumber
    
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
    m_strEANnumber = GetEANnumber()
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
    m_strLOTnumber = GetLOTnumber()
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
    m_strExpirationDate = GetExpirationDate()
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
    If Left(m_strBarcode, 2) = cmstrGtinID Then
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
Function GetExpirationDate() As String
    'Get expiration date
    'Returns: expiration date
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
        strRes = rMatch.Item(0).SubMatches(3)
    End If

    GetExpirationDate = strRes
    '---------------
FUNC_EXIT:
    On Error GoTo 0
    Exit Function
FUNC_ERR:
    Debug.Print "Error in Function GetExpirationDate"
    If Err.Number Then
        Debug.Print Err.Description
    End If
    Resume FUNC_EXIT
End Function
Function GetCatalogNumber() As String
    'Get catalog number
    'Returns: catalog number
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
        strRes = rMatch.Item(0).SubMatches(4)
    End If

    GetCatalogNumber = strRes
    '---------------
FUNC_EXIT:
    On Error GoTo 0
    Exit Function
FUNC_ERR:
    Debug.Print "Error in Function GetCatalogNumber"
    If Err.Number Then
        Debug.Print Err.Description
    End If
    Resume FUNC_EXIT
End Function

Function CreateGS1(EANnumber As String, LOTnumber As String, ExpirationDate As String, CatalogNumber As String, Optional OutputStyle As String = "Normal")
    'Create GS1 barcode
    'Parameters:
    '           EanNumber EAN number
    '           LotNumber LOT number
    '           ExpirationDate expiration date
    '           CatalogNumber catalog number
    '           OutputStyle output style Normal, Brackets, ZPL, Character
    'Returns: GS1 barcode
    'Created by: Laszlo Tamas
    
    Dim strRes
    Dim bracketBefore As String
    Dim bracketAfter As String
    Dim tempStr As String
    Dim i
    Dim tt
    
    On Error GoTo FUNC_ERR
    bracketBefore = ""
    bracketAfter = ""
    If OutputStyle = "Normal" Or OutputStyle = "Brackets" Then
        If OutputStyle = "Brackets" Then
            bracketBefore = "("
            bracketAfter = ")"
        End If
        strRes = bracketBefore + cmstrGtinID + bracketAfter + EANnumber
        strRes = strRes + bracketBefore + cmstrLotID + bracketAfter + LOTnumber
        strRes = strRes + bracketBefore + cmstrExpirationDateID + bracketAfter + ExpirationDate
        strRes = strRes + bracketBefore + cmstrCatalogNumberID + bracketAfter + CatalogNumber
        
    End If
    If OutputStyle = "ZPL" Then
        '>;>80105996527176340102014>6AA>5>8171907312128012280>64
        strRes = "^BCN,,N,N^FD>;>8" + cmstrGtinID + EANnumber + cmstrLotID + Left(LOTnumber, 4)
        If Len(LOTnumber) > 4 Then
            strRes = strRes + ">6" + Right(LOTnumber, Len(LOTnumber) - 4) + ">5"
        End If
        strRes = strRes + ">8" + cmstrExpirationDateID + ExpirationDate
        strRes = strRes + cmstrCatalogNumberID + Left(CatalogNumber, 8) + ">6" + Right(CatalogNumber, 1) + "^FS"
    End If
    If OutputStyle = "Character" Then
        
        strRes = ChrW(&HCD) + ChrW(&HCA) + ChrW(&H21)
        strRes = strRes + GetUniGS1(Mid(EANnumber, 1, 2)) + GetUniGS1(Mid(EANnumber, 3, 2)) + GetUniGS1(Mid(EANnumber, 5, 2)) + GetUniGS1(Mid(EANnumber, 7, 2))
        strRes = strRes + GetUniGS1(Mid(EANnumber, 9, 2)) + GetUniGS1(Mid(EANnumber, 11, 2)) + GetUniGS1(Mid(EANnumber, 13, 2))
        strRes = strRes + ChrW(&H2A) + GetUniGS1(Mid(LOTnumber, 1, 2)) + GetUniGS1(Mid(LOTnumber, 3, 2))
        If Len(LOTnumber) > 4 Then
            strRes = strRes + ChrW(&HC8)
            
            For i = 5 To Len(LOTnumber)
                tt = GetUniGS1(Trim(CStr(Asc(Mid(LOTnumber, i, 1)) - 32)))
                
                strRes = strRes + tt
            Next i
            strRes = strRes + ChrW(&HC7)
            
        End If
        strRes = strRes + ChrW(&HCA) + ChrW(&H31)
        strRes = strRes + GetUniGS1(Mid(ExpirationDate, 1, 2)) + GetUniGS1(Mid(ExpirationDate, 3, 2)) + GetUniGS1(Mid(ExpirationDate, 5, 2))
        strRes = strRes + ChrW(&H35)
        strRes = strRes + GetUniGS1(Mid(CatalogNumber, 1, 2)) + GetUniGS1(Mid(CatalogNumber, 3, 2)) + GetUniGS1(Mid(CatalogNumber, 5, 2)) + GetUniGS1(Mid(CatalogNumber, 7, 2))
        strRes = strRes + ChrW(&HC8)
        strRes = strRes + GetUniGS1(Trim(CStr(Asc(Right(CatalogNumber, 1)) - 32)))
        'Checkdigit and Stop
        strRes = strRes + ChrW(&H70) + ChrW(&HCE)
    End If
    
    
    
    CreateGS1 = strRes
    '---------------
FUNC_EXIT:
    On Error GoTo 0
    Exit Function
FUNC_ERR:
    Debug.Print "Error in Function CreateGS1"
    If Err.Number Then
        Debug.Print Err.Description
    End If
    Resume FUNC_EXIT
End Function

Function GetUniGS1(codeC)
    If codeC = "00" Then GetUniGS1 = ChrW(194)
    If codeC = "01" Then GetUniGS1 = ChrW(33)
    If codeC = "02" Then GetUniGS1 = ChrW(34)
    If codeC = "03" Then GetUniGS1 = ChrW(35)
    If codeC = "04" Then GetUniGS1 = ChrW(36)
    If codeC = "05" Then GetUniGS1 = ChrW(37)
    If codeC = "06" Then GetUniGS1 = ChrW(38)
    If codeC = "07" Then GetUniGS1 = ChrW(39)
    If codeC = "08" Then GetUniGS1 = ChrW(40)
    If codeC = "09" Then GetUniGS1 = ChrW(41)
    If codeC = "10" Then GetUniGS1 = ChrW(42)
    If codeC = "11" Then GetUniGS1 = ChrW(43)
    If codeC = "12" Then GetUniGS1 = ChrW(44)
    If codeC = "13" Then GetUniGS1 = ChrW(45)
    If codeC = "14" Then GetUniGS1 = ChrW(46)
    If codeC = "15" Then GetUniGS1 = ChrW(47)
    If codeC = "16" Then GetUniGS1 = ChrW(48)
    If codeC = "17" Then GetUniGS1 = ChrW(49)
    If codeC = "18" Then GetUniGS1 = ChrW(50)
    If codeC = "19" Then GetUniGS1 = ChrW(51)
    If codeC = "20" Then GetUniGS1 = ChrW(52)
    If codeC = "21" Then GetUniGS1 = ChrW(53)
    If codeC = "22" Then GetUniGS1 = ChrW(54)
    If codeC = "23" Then GetUniGS1 = ChrW(55)
    If codeC = "24" Then GetUniGS1 = ChrW(56)
    If codeC = "25" Then GetUniGS1 = ChrW(57)
    If codeC = "26" Then GetUniGS1 = ChrW(58)
    If codeC = "27" Then GetUniGS1 = ChrW(59)
    If codeC = "28" Then GetUniGS1 = ChrW(60)
    If codeC = "29" Then GetUniGS1 = ChrW(61)
    If codeC = "30" Then GetUniGS1 = ChrW(62)
    If codeC = "31" Then GetUniGS1 = ChrW(63)
    If codeC = "32" Then GetUniGS1 = ChrW(64)
    If codeC = "33" Then GetUniGS1 = ChrW(65)
    If codeC = "34" Then GetUniGS1 = ChrW(66)
    If codeC = "35" Then GetUniGS1 = ChrW(67)
    If codeC = "36" Then GetUniGS1 = ChrW(68)
    If codeC = "37" Then GetUniGS1 = ChrW(69)
    If codeC = "38" Then GetUniGS1 = ChrW(70)
    If codeC = "39" Then GetUniGS1 = ChrW(71)
    If codeC = "40" Then GetUniGS1 = ChrW(72)
    If codeC = "41" Then GetUniGS1 = ChrW(73)
    If codeC = "42" Then GetUniGS1 = ChrW(74)
    If codeC = "43" Then GetUniGS1 = ChrW(75)
    If codeC = "44" Then GetUniGS1 = ChrW(76)
    If codeC = "45" Then GetUniGS1 = ChrW(77)
    If codeC = "46" Then GetUniGS1 = ChrW(78)
    If codeC = "47" Then GetUniGS1 = ChrW(79)
    If codeC = "48" Then GetUniGS1 = ChrW(80)
    If codeC = "49" Then GetUniGS1 = ChrW(81)
    If codeC = "50" Then GetUniGS1 = ChrW(82)
    If codeC = "51" Then GetUniGS1 = ChrW(83)
    If codeC = "52" Then GetUniGS1 = ChrW(84)
    If codeC = "53" Then GetUniGS1 = ChrW(85)
    If codeC = "54" Then GetUniGS1 = ChrW(86)
    If codeC = "55" Then GetUniGS1 = ChrW(87)
    If codeC = "56" Then GetUniGS1 = ChrW(88)
    If codeC = "57" Then GetUniGS1 = ChrW(89)
    If codeC = "58" Then GetUniGS1 = ChrW(90)
    If codeC = "59" Then GetUniGS1 = ChrW(91)
    If codeC = "60" Then GetUniGS1 = ChrW(92)
    If codeC = "61" Then GetUniGS1 = ChrW(93)
    If codeC = "62" Then GetUniGS1 = ChrW(94)
    If codeC = "63" Then GetUniGS1 = ChrW(95)
    If codeC = "64" Then GetUniGS1 = ChrW(96)
    If codeC = "65" Then GetUniGS1 = ChrW(97)
    If codeC = "66" Then GetUniGS1 = ChrW(98)
    If codeC = "67" Then GetUniGS1 = ChrW(99)
    If codeC = "68" Then GetUniGS1 = ChrW(100)
    If codeC = "69" Then GetUniGS1 = ChrW(101)
    If codeC = "70" Then GetUniGS1 = ChrW(102)
    If codeC = "71" Then GetUniGS1 = ChrW(103)
    If codeC = "72" Then GetUniGS1 = ChrW(104)
    If codeC = "73" Then GetUniGS1 = ChrW(105)
    If codeC = "74" Then GetUniGS1 = ChrW(106)
    If codeC = "75" Then GetUniGS1 = ChrW(107)
    If codeC = "76" Then GetUniGS1 = ChrW(108)
    If codeC = "77" Then GetUniGS1 = ChrW(109)
    If codeC = "78" Then GetUniGS1 = ChrW(110)
    If codeC = "79" Then GetUniGS1 = ChrW(111)
    If codeC = "80" Then GetUniGS1 = ChrW(112)
    If codeC = "81" Then GetUniGS1 = ChrW(113)
    If codeC = "82" Then GetUniGS1 = ChrW(114)
    If codeC = "83" Then GetUniGS1 = ChrW(115)
    If codeC = "84" Then GetUniGS1 = ChrW(116)
    If codeC = "85" Then GetUniGS1 = ChrW(117)
    If codeC = "86" Then GetUniGS1 = ChrW(118)
    If codeC = "87" Then GetUniGS1 = ChrW(119)
    If codeC = "88" Then GetUniGS1 = ChrW(120)
    If codeC = "89" Then GetUniGS1 = ChrW(121)
    If codeC = "90" Then GetUniGS1 = ChrW(122)
    If codeC = "91" Then GetUniGS1 = ChrW(123)
    If codeC = "92" Then GetUniGS1 = ChrW(124)
    If codeC = "93" Then GetUniGS1 = ChrW(125)
    If codeC = "94" Then GetUniGS1 = ChrW(126)
    If codeC = "95" Then GetUniGS1 = ChrW(195)
    If codeC = "96" Then GetUniGS1 = ChrW(196)
    If codeC = "97" Then GetUniGS1 = ChrW(197)
    If codeC = "98" Then GetUniGS1 = ChrW(198)
    If codeC = "99" Then GetUniGS1 = ChrW(199)
End Function
Function ParseGS1() As String()
    'Parse GS1 code
    'Returns: array with barcode elements, (01) EAN, (10) LOT, (17) expiration date, (21) catalog number
    'Created by: Laszlo Tamas

    Dim strRes(3) As String

    On Error GoTo FUNC_ERR

    strRes(0) = GetEANnumber()
    strRes(1) = GetLOTnumber()
    strRes(2) = GetExpirationDate()
    strRes(3) = GetCatalogNumber()
    'Code here

    ParseGS1 = strRes
    '---------------
FUNC_EXIT:
    On Error GoTo 0
    Exit Function
FUNC_ERR:
    Debug.Print "Error in Function ParseGS1"
    If Err.Number Then
        Debug.Print Err.Description
    End If
    Resume FUNC_EXIT
End Function








