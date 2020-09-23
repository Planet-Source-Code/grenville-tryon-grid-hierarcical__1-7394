Attribute VB_Name = "PubFun"

'CONVIERTE UNA Chain EN ARRAY
Public Function TSStrArr(ByVal Chain As String, ByVal Separator As String)
Dim Rpta() As Variant, Contador As String, Numero As Integer, Buffer As String
Numero = 0
If IsEmpty(Separator) Then
 Separator = "|"
End If
Chain = Chain + Separator
While InStr(1, Chain, Separator) <> 0
    Buffer = Mid(Chain, 1, InStr(1, Chain, Separator) - 1)
    ReDim Preserve Rpta(Numero)
    Rpta(Numero) = Buffer
    Numero = Numero + 1
    Chain = Mid(Chain, InStr(1, Chain, Separator) + 1)
Wend
TSStrArr = Rpta
End Function

