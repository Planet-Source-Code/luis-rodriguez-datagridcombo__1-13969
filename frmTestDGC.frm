VERSION 5.00
Object = "{21886BFE-DC13-11D4-B0D1-00C04F29F4F9}#2.0#0"; "DGridCombo.ocx"
Begin VB.Form Form1 
   Caption         =   "Test"
   ClientHeight    =   4635
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6705
   LinkTopic       =   "Form1"
   ScaleHeight     =   4635
   ScaleWidth      =   6705
   StartUpPosition =   3  'Windows Default
   Begin DGridCombo.GridCombo GridCombo1 
      Height          =   315
      Index           =   1
      Left            =   960
      TabIndex        =   0
      Top             =   960
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   556
      BackStyle       =   0
      DGStyle         =   1
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub GridCombo1_DropDown(Index As Integer)
    Dim Rs As ADODB.Recordset, i As Integer
    If GridCombo1(Index).ListCount <= 0 Then
        Set Rs = New ADODB.Recordset
        Rs.Fields.Append "Field1", adChar, 4, adFldFixed
        Rs.Fields.Append "Field2", adChar, 4, adFldFixed
        Rs.Fields.Append "Field3", adChar, 4, adFldFixed
        Rs.Fields.Append "Field4", adChar, 4, adFldFixed
        Rs.Fields.Append "Field5", adChar, 4, adFldFixed
        Rs.Fields.Append "Field6", adChar, 4, adFldFixed
        Rs.Open
        For i = 0 To 20
            Rs.AddNew
            Rs(0) = CStr(i) & "AA"
            Rs(1) = CStr(i) & "BB"
            Rs(2) = CStr(i) & "CC"
            Rs(3) = CStr(i) & "DD"
            Rs(4) = CStr(i) & "EE"
            Rs(5) = CStr(i) & "FF"
        Next i
        Set GridCombo1(Index).RowSource = Rs
        Set Rs = Nothing
        
    End If
End Sub
