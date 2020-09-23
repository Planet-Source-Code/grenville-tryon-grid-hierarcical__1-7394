VERSION 5.00
Object = "*\ATS.vbp"
Begin VB.Form frmDemo 
   Caption         =   "Form1"
   ClientHeight    =   3540
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5055
   LinkTopic       =   "Form1"
   ScaleHeight     =   3540
   ScaleWidth      =   5055
   StartUpPosition =   3  'Windows Default
   Begin TS.TSFlexGrid TS 
      Height          =   3495
      Left            =   90
      TabIndex        =   0
      Top             =   0
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   6165
   End
End
Attribute VB_Name = "frmDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Dim MSF As MSFlexGrid
Set MSF = TS.GETMSF
MSF.FormatString = "#                     |NAME OF THE ELEMENT   |SONS     "
TS.SetColor 0, RGB(0, 128, 128)
TS.SetColor 1, RGB(0, 164, 164)
TS.SetColor 2, RGB(0, 192, 192)
TS.SetColor 3, RGB(0, 218, 218)
TS.SetColor 4, RGB(0, 255, 255)
TS.SetSeparator "."
TS.Add "A001", "A001" + Chr(9) + "A" + Chr(9) + "0"
TS.Add "A001.001", "A001001" + Chr(9) + "C" + Chr(9) + "0"
TS.Add "A001.001.001", "A001001001" + Chr(9) + "D" + Chr(9) + "0"
TS.Add "A001.002", "A001002" + Chr(9) + "E" + Chr(9) + "0"
TS.Add "A001.002.001", "A001002001" + Chr(9) + "F" + Chr(9) + "0"
TS.Add "A001.002.002", "A001002001" + Chr(9) + "G" + Chr(9) + "0"
TS.Add "A001.002.002.001", "A001002001001" + Chr(9) + "H" + Chr(9) + "0"
TS.Add "A002", "A002" + Chr(9) + "I" + Chr(9) + "0"
TS.Add "A003", "A003" + Chr(9) + "J" + Chr(9) + "0"
TS.Add "A003.001", "A003001" + Chr(9) + "K" + Chr(9) + "0"
TS.Add "A003.001.001", "A003001001" + Chr(9) + "L" + Chr(9) + "0"
TS.Add "A003.002", "A003002" + Chr(9) + "M" + Chr(9) + "0"
TS.Add "A003.002.001", "A003002001" + Chr(9) + "N" + Chr(9) + "0"
TS.Add "A003.002.002", "A003002001" + Chr(9) + "O" + Chr(9) + "0"
TS.Add "A003.002.002.001", "A003002001001" + Chr(9) + "P" + Chr(9) + "0"
TS.Add "A003.002.002.001.001", "A003002001001001" + Chr(9) + "Q" + Chr(9) + "0"
TS.Draw
Status False
MSF.SelectionMode = flexSelectionByRow
MSF.Row = 1
MSF.Col = 1
End Sub

Private Sub TS_DblClick()
If TS.IsHide Then
    TS.Show True
Else
    TS.Hide
End If
Status True
End Sub

Private Sub TS_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
Case 45
    TS.Hide
Case 43
    TS.Show False
End Select
Status True
End Sub

Private Sub Status(OnScreen As Boolean)
Dim MSF As MSFlexGrid, Cont As Integer, ActRow As Integer
Set MSF = TS.GETMSF
ActRow = MSF.Row
MSF.Visible = False
For Cont = 1 To MSF.Rows - 1
    MSF.Row = Cont
    MSF.Col = 1
    If TS.IsHide Then
        MSF.CellFontBold = True
    Else
        MSF.CellFontBold = False
    End If
    MSF.TextMatrix(Cont, 2) = IIf(TS.Sons = 0, "", CStr(TS.Sons))
Next
MSF.Visible = True
If OnScreen Then
    MSF.SetFocus
End If
MSF.Row = ActRow
End Sub
