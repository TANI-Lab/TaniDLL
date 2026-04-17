Attribute VB_Name = "Module1"
' =========================
' DLL 宣言
' =========================
#If VBA7 Then
    #If Win64 Then
        Private Declare PtrSafe Function TANI_OpenSimple Lib "TaniDLL64.dll" ( _
            ByVal id As Long, _
            ByVal name As String, _
            ByVal BAUD As Long) As Long
    
        Private Declare PtrSafe Function TANI_OpenEx Lib "TaniDLL64.dll" ( _
            ByVal id As Long, _
            ByVal name As String, _
            ByVal BAUD As Long, _
            ByVal databits As Long, _
            ByVal parity As Long, _
            ByVal stopbits As Long, _
            ByVal flow As Long) As Long
    
        Private Declare PtrSafe Sub TANI_Close Lib "TaniDLL64.dll" ( _
            ByVal id As Long)
    
        Private Declare PtrSafe Function TANI_SendText Lib "TaniDLL64.dll" ( _
            ByVal id As Long, _
            ByVal txt As String, _
            ByVal TERM As String) As Long
    
        Private Declare PtrSafe Function TANI_SendBytes Lib "TaniDLL64.dll" ( _
            ByVal id As Long, _
            ByRef data As Byte, _
            ByVal length As Long) As Long
    
        Private Declare PtrSafe Function TANI_RecvLine Lib "TaniDLL64.dll" ( _
            ByVal id As Long, _
            ByRef buf As Byte, _
            ByVal bufSize As Long, _
            ByVal timeout As Long, _
            ByVal TERM As Byte) As Long
    
        Private Declare PtrSafe Function TANI_RecvBytesUntil Lib "TaniDLL64.dll" _
            (ByVal id As Long, _
             ByRef buf As Byte, _
             ByVal bufSize As Long, _
             ByVal termByte As Byte, _
             ByVal inclTerm As Long, _
             ByVal timeout As Long) As Long
             
        Private Declare PtrSafe Function TANI_RecvBytes Lib "TaniDLL64.dll" ( _
            ByVal id As Long, _
            ByRef buf As Byte, _
            ByVal size As Long, _
            ByVal timeout As Long) As Long
            
        Private Declare PtrSafe Function TANI_CheckLicense Lib "TaniDLL64.dll" ( _
         ByVal LicenseKey As String) As Long
    
    #Else
        ' 32bit用
        
        Private Declare Function TANI_OpenSimple Lib "TaniDLL32.dll" ( _
            ByVal id As Long, _
            ByVal name As String, _
            ByVal baud As Long _
        ) As Long
    
        Private Declare Function TANI_OpenEx Lib "TaniDLL32.dll" ( _
            ByVal id As Long, _
            ByVal name As String, _
            ByVal baud As Long, _
            ByVal dataBits As Long, _
            ByVal parity As Long, _
            ByVal stopBits As Long, _
            ByVal flow As Long _
        ) As Long
    
        Private Declare Function TANI_SendText Lib "TaniDLL32.dll" ( _
            ByVal id As Long, _
            ByVal txt As String, _
            ByVal term As String _
        ) As Long
    
        Private Declare Function TANI_SendBytes Lib "TaniDLL32.dll" ( _
            ByVal id As Long, _
            ByRef buf As Byte, _
            ByVal size As Long) As Long
    
        Private Declare Function TANI_RecvLine Lib "TaniDLL32.dll" ( _
            ByVal id As Long, _
            ByRef buf As Byte, _
            ByVal bufSize As Long, _
            ByVal timeout As Long, _
            ByVal term As Byte) As Long
    
        Private Declare Function TANI_RecvBytes Lib "TaniDLL32.dll" ( _
            ByVal id As Long, _
            ByRef buf As Byte, _
            ByVal size As Long, _
            ByVal timeout As Long _
        ) As Long
    
        Private Declare Function TANI_RecvBytesUntil Lib "TaniDLL32.dll" _
        (ByVal id As Long, _
         ByRef buf As Byte, _
         ByVal bufSize As Long, _
         ByVal termByte As Byte, _
         ByVal inclTerm As Long, _
         ByVal timeout As Long) As Long
         
        Private Declare Sub TANI_Close Lib "TaniDLL32.dll" ( _
            ByVal id As Long _
        )
        Private Declare Function TANI_CheckLicense Lib "TaniDLL32.dll" ( _
         ByVal LicenseKey As String) As Long
    
    #End If
#Else
    ' 32bit用
    
    Private Declare Function TANI_OpenSimple Lib "TaniDLL32.dll" ( _
        ByVal id As Long, _
        ByVal name As String, _
        ByVal baud As Long _
    ) As Long

    Private Declare Function TANI_OpenEx Lib "TaniDLL32.dll" ( _
        ByVal id As Long, _
        ByVal name As String, _
        ByVal baud As Long, _
        ByVal dataBits As Long, _
        ByVal parity As Long, _
        ByVal stopBits As Long, _
        ByVal flow As Long _
    ) As Long

    Private Declare Function TANI_SendText Lib "TaniDLL32.dll" ( _
        ByVal id As Long, _
        ByVal txt As String, _
        ByVal term As String _
    ) As Long

    Private Declare Function TANI_SendBytes Lib "TaniDLL32.dll" ( _
        ByVal id As Long, _
        ByRef buf As Byte, _
        ByVal size As Long) As Long

    Private Declare Function TANI_RecvLine Lib "TaniDLL32.dll" ( _
        ByVal id As Long, _
        ByRef buf As Byte, _
        ByVal bufSize As Long, _
        ByVal timeout As Long, _
        ByVal term As Byte) As Long

    Private Declare Function TANI_RecvBytes Lib "TaniDLL32.dll" ( _
        ByVal id As Long, _
        ByRef buf As Byte, _
        ByVal size As Long, _
        ByVal timeout As Long _
    ) As Long

    Private Declare Function TANI_RecvBytesUntil Lib "TaniDLL32.dll" _
    (ByVal id As Long, _
     ByRef buf As Byte, _
     ByVal bufSize As Long, _
     ByVal termByte As Byte, _
     ByVal inclTerm As Long, _
     ByVal timeout As Long) As Long
     
    Private Declare Sub TANI_Close Lib "TaniDLL32.dll" ( _
        ByVal id As Long _
    )
    Private Declare Function TANI_CheckLicense Lib "TaniDLL32.dll" ( _
     ByVal LicenseKey As String) As Long

#End If


Option Explicit

' =========================
' 設定
' =========================
Const PORT_ID As Long = 0
Const COM_NAME As String = "COM5"
Const BAUD As Long = 115200

Const LF As Byte = 10
Const CR As Byte = 13
Const TERM As Byte = 35 ' "#"

' =========================
' メイン
' =========================
Sub TaniDLL_AutoTest()

    Debug.Print "==== AUTO TEST START ===="

    'If Not Test_License Then Exit Sub
    If Not Test_Open Then Exit Sub

    Test_Text
    Test_Binary
    Test_Terminator
    Test_Timeout
    Test_BufferOverflow
    Test_CRLF
    Test_Continuous
    Test_Invalid

    TANI_Close PORT_ID

    Debug.Print "==== AUTO TEST END ===="

End Sub

' =========================
' 個別テスト
' =========================

Function Test_License() As Boolean
    Dim ret As Long
    ret = TANI_CheckLicense("YOUR_LICENSE_KEY")
    PrintResult "License", (ret = 1), ret
    Test_License = (ret = 1)
End Function

Function Test_Open() As Boolean
    Dim ret As Long
    ret = TANI_OpenSimple(PORT_ID, COM_NAME, BAUD)
    PrintResult "Open", (ret = 0), ret
    Test_Open = (ret = 0)
End Function

Sub Test_Text()
    Dim buf(0 To 255) As Byte
    Dim len1 As Long

    TANI_SendText PORT_ID, "PING", vbLf
    len1 = TANI_RecvLine(PORT_ID, buf(0), 256, 2000, LF)

    PrintResult "TextRecv", (len1 > 0), len1
    Debug.Print "  Data:", BytesToString(buf, len1)
End Sub

Sub Test_Binary()
    Dim buf(0 To 255) As Byte
    Dim len1 As Long

    TANI_SendText PORT_ID, "BINTEST", vbLf
    len1 = TANI_RecvBytes(PORT_ID, buf(0), 4, 2000)

    PrintResult "BinaryRecv", (len1 = 4), len1
End Sub

Sub Test_Terminator()
    Dim buf(0 To 255) As Byte
    Dim len1 As Long

    TANI_SendText PORT_ID, "TERMTEST", vbLf
    len1 = TANI_RecvLine(PORT_ID, buf(0), 256, 2000, TERM)

    PrintResult "Terminator", (len1 > 0), len1
End Sub

Sub Test_Timeout()
    Dim buf(0 To 255) As Byte
    Dim len1 As Long

    len1 = TANI_RecvLine(PORT_ID, buf(0), 256, 10, LF)
    PrintResult "Timeout", (len1 = 0), len1
End Sub

Sub Test_BufferOverflow()
    Dim buf(0 To 9) As Byte
    Dim len1 As Long

    TANI_SendText PORT_ID, String(50, "A"), vbLf
    len1 = TANI_RecvLine(PORT_ID, buf(0), 10, 2000, LF)

    PrintResult "BufferLimit", (len1 <= 10), len1
End Sub

Sub Test_CRLF()
    Dim buf(0 To 255) As Byte
    Dim len1 As Long

    TANI_SendText PORT_ID, "PING", vbLf
    len1 = TANI_RecvLine(PORT_ID, buf(0), 256, 2000, LF)

    If len1 > 0 Then
        PrintResult "CR Check", (buf(len1 - 1) = CR), buf(len1 - 1)
    Else
        PrintResult "CR Check", False, len1
    End If
End Sub

Sub Test_Continuous()
    Dim i As Long
    Dim buf(0 To 255) As Byte
    Dim len1 As Long
    Dim ok As Boolean

    ok = True

    For i = 1 To 100
        TANI_SendText PORT_ID, "PING", vbLf
        len1 = TANI_RecvLine(PORT_ID, buf(0), 256, 1000, LF)
        If len1 = 0 Then ok = False
    Next

    PrintResult "Continuous", ok, 100
End Sub

Sub Test_Invalid()
    Dim ret As Long
    ret = TANI_SendText(99, "TEST", vbLf)
    PrintResult "InvalidID", (ret = -1 Or ret = 0), ret
End Sub

' =========================
' ユーティリティ
' =========================

Sub PrintResult(name As String, ok As Boolean, val As Variant)
    If ok Then
        Debug.Print "[PASS]", name, val
    Else
        Debug.Print "[FAIL]", name, val
    End If
End Sub

Function BytesToString(buf() As Byte, length As Long) As String
    Dim i As Long, s As String

    For i = 0 To length - 1
        If buf(i) <> 13 Then ' CR除去
            s = s & Chr(buf(i))
        End If
    Next

    BytesToString = s
End Function
