Attribute VB_Name = "Module2"
' =========================
' DLL êÈåæ
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
        ' 32bitóp
        
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
    ' 32bitóp
    
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

Sub TestSerial()

  Dim ret As Long
  Dim buf(255) As Byte

  ' COM open
  ret = TANI_OpenEx(0, "COM5", 115200, 8, 0, 1, 0)
  If ret <> 0 Then
    MsgBox "open error", vbCritical
    Exit Sub
  End If

  ' sned
  ret = TANI_SendText(0, "HELLO", vbCr)
  If ret <= 0 Then
    MsgBox "send error", vbCritical
    GoTo EXIT_PROC
  End If

  ' Recv
  ret = TANI_RecvLine(0, buf(0), 256, 2000, 13)
  If ret > 0 Then
    Dim s As String
    s = StrConv(buf, vbUnicode)
    MsgBox "Recv: " & Left(s, ret)
  ElseIf ret = 0 Then
    MsgBox "timeout"
  Else
    MsgBox "Recv error", vbCritical
  End If

EXIT_PROC:
  TANI_Close 0

End Sub

