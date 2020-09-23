VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsConcat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private sBuffer As String
Private Const ciIncriment As Integer = 15000
Private lOffset As Long

Public Sub ReInit()
sBuffer = ""
lOffset = 0
End Sub

Public Function GetString() As String
GetString = Left$(sBuffer, lOffset)
sBuffer = ""  'reset
End Function

'This function lets you assign a string to the concating buffer.
Public Sub SetString(ByRef Source As String)
sBuffer = Source & String$(ciIncriment, 0)
End Sub

Public Sub SConcat(ByRef Source As String)
Dim lBufferLen As Long
lBufferLen = Len(Source)
'Allocate more space in buffer if needed
If (lOffset + lBufferLen) >= Len(sBuffer) Then
   If lBufferLen > lOffset Then
      sBuffer = sBuffer & String$(lBufferLen, 0)
   Else
      sBuffer = sBuffer & String$(ciIncriment, 0)
   End If
End If
Mid$(sBuffer, lOffset + 1, lBufferLen) = Source
lOffset = lOffset + lBufferLen
End Sub
