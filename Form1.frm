VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "String Concatenation Demonstration"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&Run Test"
      Height          =   495
      Left            =   1800
      TabIndex        =   0
      Top             =   1320
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function timeGetTime Lib "winmm.dll" () As Long

Private Sub Command1_Click()
Dim StartTime As Long
Dim i As Long
Dim n As Long
Dim s1 As String
Dim s2 As String
Dim s3 As String
Dim s4 As String
Dim var1 As clsConcat
Dim var2 As clsConcat
Const NUM_ITEMS = 5000
s1 = "add me" & vbCrLf
s2 = "me too" & vbCrLf

'Buffer Concatenation
'Note: If you want to concatenate to 2 variables then
'you should instantiate 2 instances of the class
'as this example demonstrates
Screen.MousePointer = vbHourglass
Set var1 = New clsConcat 'start new instance of class
Set var2 = New clsConcat 'start another new instance of class
StartTime = timeGetTime
For i = 1 To NUM_ITEMS
    var1.SConcat s1
    var1.SConcat s2
    var2.SConcat s1
    var2.SConcat s2
Next
'Get the strings!
s3 = var1.GetString
s4 = var2.GetString
'Clear objects from memory
Set var1 = Nothing
Set var2 = Nothing
Screen.MousePointer = vbDefault
MsgBox "Buffer concatenation took " & (timeGetTime - StartTime) / 1000 & " secs for " & CStr(Len(s3) + Len(s4)) & " characters."

'Regular concatenation
Screen.MousePointer = vbHourglass
s3 = ""
s4 = ""
StartTime = timeGetTime
For i = 1 To NUM_ITEMS
    s3 = s3 & s1 & s2
    s4 = s4 & s1 & s2
Next
Screen.MousePointer = vbDefault
MsgBox "Regular concatenation took " & (timeGetTime - StartTime) / 1000 & " secs for " & CStr(Len(s3) + Len(s4)) & " characters."

End Sub

