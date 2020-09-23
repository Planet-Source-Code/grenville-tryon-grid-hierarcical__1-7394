VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.UserControl TSFlexGrid 
   ClientHeight    =   2310
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2925
   ScaleHeight     =   2310
   ScaleWidth      =   2925
   Begin MSFlexGridLib.MSFlexGrid MSF 
      Height          =   1575
      Left            =   180
      TabIndex        =   0
      Top             =   420
      Width           =   2445
      _ExtentX        =   4313
      _ExtentY        =   2778
      _Version        =   65541
      Rows            =   20
      Cols            =   20
      ScrollTrack     =   -1  'True
      FocusRect       =   0
   End
End
Attribute VB_Name = "TSFlexGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Public cSeparator As String
Public ArrData As String
Public ArrColor As String
Public nHeight As Double

Event Click()
Event DblClick()
Event KeyPress(KeyAscii As Integer)

'DEFINE THE COLOR OF THE LEVELS
Public Function SetColor(Level As Integer, Color As Double) As Boolean
Dim Arr As Variant, Cont As Integer
Arr = TSStrArr(ArrColor, "|")
ArrColor = ""
For Cont = 0 To UBound(Arr, 1) - 1
    If Cont <> Level Then
        ArrColor = ArrColor + CStr(Arr(Cont)) + "|"
    Else
        ArrColor = ArrColor + CStr(Color) + "|"
    End If
Next
End Function

'DRAWS WITH THE CHARACTERISTICS
Public Function Draw() As Boolean
Dim Cont2 As Integer, Cont As Integer, ArrD As Variant, Level As Integer, ArrC As Variant, ActCol As Integer
ActCol = MSF.Col
MSF.Cols = MSF.Cols + 1
ArrD = TSStrArr(ArrData, "|")
ArrC = TSStrArr(ArrColor, "|")
For Cont = 1 To MSF.Rows - 1
    MSF.TextMatrix(Cont, MSF.Cols - 1) = CStr(ArrD(Cont - 1))
    Level = UBound(TSStrArr(CStr(ArrD(Cont - 1)), cSeparator), 1)
    MSF.Row = Cont
    For Cont2 = MSF.FixedCols To MSF.Cols - 1
        MSF.Col = Cont2
        MSF.CellBackColor = CDbl(Val(ArrC(Level)))
    Next
Next
MSF.Col = MSF.Cols - 1
MSF.Sort = 0
MSF.Cols = MSF.Cols - 1
MSF.Col = ActCol
End Function

'DEFINE THE SEPARATOR ITEM
Public Function SetSeparator(Separator As String) As Boolean
cSeparator = Separator
End Function

'RETURN THE FLEX REFERENCE
Public Function GETMSF() As Variant
Set GETMSF = MSF
End Function

'ADD AN ELEMENT TO THE GRID
Public Function ADD(Chain As String, Text As String) As Boolean
ArrData = ArrData + Chain + "|"
MSF.AddItem Text
End Function

'HIDED THE SONS?
Public Function IsHide() As Boolean
IsHide = False
If MSF.Row + 1 <> MSF.Rows Then
    IsHide = IIf(MSF.RowHeight(MSF.Row + 1) = 0, True, False)
End If
End Function

'SHOW THE MINOR LEVELS
Public Function Show(ShowAll As Boolean)
Dim ArrD As Variant, Chain As String, Cont As Integer, ActRow As Integer, Level As Integer
ArrD = TSStrArr(ArrData, "|")
ActRow = MSF.Row
Chain = ArrD(ActRow - 1)
Level = UBound(TSStrArr(Chain, cSeparator), 1)
Do While Mid(CStr(ArrD(ActRow)), 1, Len(Chain)) = Chain
    If Level + 1 = UBound(TSStrArr(CStr(ArrD(ActRow)), cSeparator), 1) Or ShowAll Then
        MSF.RowHeight(ActRow + 1) = nHeight
    End If
    ActRow = ActRow + 1
Loop
End Function

'HIDE THE MINOR LEVELS
Public Function Hide()
Dim ArrD As Variant, Chain As String, Cont As Integer, ActRow As Integer, Level As Integer
ArrD = TSStrArr(ArrData, "|")
ActRow = MSF.Row
Chain = ArrD(ActRow - 1)
Level = UBound(TSStrArr(Chain, cSeparator), 1)
Do While Mid(CStr(ArrD(ActRow)), 1, Len(Chain)) = Chain ' And (Level + 1 = UBound(TSStrArr(CStr(ArrD(ActRow)), cSeparator), 1) Or Level = 0)
    MSF.RowHeight(ActRow + 1) = 0
    ActRow = ActRow + 1
Loop
End Function

'COUNT NUMBER OF SONS
Public Function Sons() As Integer
Dim ArrD As Variant, Chain As String, Cont As Integer, ActRow As Integer, Level As Integer
Sons = 0
ArrD = TSStrArr(ArrData, "|")
ActRow = MSF.Row
Chain = ArrD(ActRow - 1)
Level = UBound(TSStrArr(Chain, cSeparator), 1)
Do While Mid(CStr(ArrD(ActRow)), 1, Len(Chain)) = Chain And ActRow < MSF.Rows
    If Level + 1 = UBound(TSStrArr(CStr(ArrD(ActRow)), cSeparator), 1) Then
        Sons = Sons + 1
    End If
    ActRow = ActRow + 1
Loop
End Function

Private Sub MSF_Click()
RaiseEvent Click
End Sub

Private Sub MSF_DblClick()
RaiseEvent DblClick
End Sub

Private Sub MSF_KeyPress(KeyAscii As Integer)
RaiseEvent KeyPress(KeyAscii)
End Sub

'INITIALIZE SETTINGS
Private Sub UserControl_Initialize()
MSF.Rows = 1
MSF.Cols = 1
cSeparator = "."
ArrColor = String(20, "|")
nHeight = MSF.RowHeight(0)
End Sub

'ADJUST THE SIZE OF THE GRID
Private Sub UserControl_Resize()
MSF.Left = 0
MSF.Top = 0
MSF.Height = UserControl.Height
MSF.Width = UserControl.Width
End Sub

