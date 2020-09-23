Attribute VB_Name = "instruct"
Declare Sub InitCommonControls Lib "comctl32.dll" ()
Public M, n, O As Integer
Public Ftrigo As Boolean
Public Function factorielle(ByVal x As Integer) As Double
   'Input: non-negative integer
   'Usage: Factorial (4)
Dim y

   If x < 0 Then
    y = MsgBox("Input value must be a non-negative integer.", vbCritical)
       Exit Function
   End If


   If x = 0 Then
       factorielle = 1
   Else
       factorielle = x * factorielle(x - 1)
   End If
End Function

Sub sort()
With GraphPlot
    .Label1(0).ForeColor = .Label1(1).ForeColor
    .Label1(0).Caption = .Label1(1).Caption
    .Label1(1).ForeColor = .Label1(2).ForeColor
    .Label1(1).Caption = .Label1(2).Caption
    .Label1(2).ForeColor = .Label1(3).ForeColor
    .Label1(2).Caption = .Label1(3).Caption
    .Label1(3).ForeColor = .Label1(4).ForeColor
    .Label1(3).Caption = .Label1(4).Caption
    .Label1(4).ForeColor = .Label1(5).ForeColor
    .Label1(4).Caption = .Label1(5).Caption
    .Label1(5).ForeColor = .Label1(6).ForeColor
    .Label1(5).Caption = .Label1(6).Caption
    .Label1(6).ForeColor = .Label1(7).ForeColor
    .Label1(6).Caption = .Label1(7).Caption
    .Label1(7).ForeColor = .Label1(8).ForeColor
    .Label1(7).Caption = .Label1(8).Caption
    .Label1(8).ForeColor = .Label1(9).ForeColor
    .Label1(8).Caption = .Label1(9).Caption
    .Label1(9).ForeColor = .Label1(10).ForeColor
    .Label1(9).Caption = .Label1(10).Caption
    .Label1(10).ForeColor = .Label1(11).ForeColor
    .Label1(10).Caption = .Label1(11).Caption
    .Label1(11).ForeColor = .Label1(12).ForeColor
    .Label1(11).Caption = .Label1(12).Caption
End With
End Sub
Sub RndColor()
    M = Int(Rnd * 255)
    n = Int(Rnd * 255)
    O = Int(Rnd * 255)
End Sub
'---- Décompose un nombre en facteurs premiers
Public Function Decompose(ByVal Nombre As Long) As Long()
    ' Tableau pour le résultat
    Dim Res() As Long
    ' Nombre d'éléments trouvés
    Dim nb As Integer: nb = 1
    ' Place le 1
    ReDim Res(1 To 1) As Long
    Res(1) = 1
    
    ' Diviseur courant
    Dim Div As Long: Div = 2
    ' Tant que le diviseur est inférieur ou égal au nombre
    Do While Div <= Nombre
        ' Tant que Div divise Nombre
        Do While Nombre Mod Div = 0
            ' Ajoute dans le tableau résultat
            nb = nb + 1
            ReDim Preserve Res(1 To nb) As Long
            Res(nb) = Div
            ' Nombre suivant
            Nombre = Nombre / Div
        Loop
        ' Cherche le prochain premier
        Div = ProchainPremier(Div)
    Loop
    
    ' Résultat
    Decompose = Res
End Function

'---- ProchainPremier : retourne le prochain nombre premier
Public Function ProchainPremier(ByVal Nombre As Long) As Long
    Do
        Nombre = Nombre + 1
    Loop Until EstPremier(Nombre)
    ProchainPremier = Nombre
End Function

'---- EstPremier : indique si un nombre est premier
Public Function EstPremier(ByVal Nombre As Long) As Boolean
    ' Positif seulement
    Nombre = Abs(Nombre)
    
    EstPremier = True
    If Nombre > 3 Then
        Dim i As Long
        For i = 2 To Nombre / 2
            If Nombre Mod i = 0 Then
                EstPremier = False
                Exit For
            End If
        Next
    End If
End Function
Function ConvertToBase(DecNumber As Double, NewBase As Integer) As String
On Error GoTo 5
    Dim ModBase As Double


    Do
        ModBase = CDbl(DecNumber - (Int(DecNumber / NewBase)) * NewBase)
        DecNumber = Int(DecNumber / NewBase)
        If ModBase > 9 Then ModBase = ModBase + 7
        ConvertToBase = Chr(ModBase + 48) & ConvertToBase
    Loop Until DecNumber = 0

5 End Function



Function ConvertFromBase(BaseNumber As String, OldBase As Integer) As Double
On Error GoTo 5
    Dim i As Integer, LetterVal As Integer
    On Error Resume Next


    For i = 1 To Len(BaseNumber)
        LetterVal = Asc(Mid(BaseNumber, Len(BaseNumber) - i + 1, 1)) - 48
        If LetterVal > 9 Then LetterVal = LetterVal - 7
        If LetterVal > OldBase Then GoTo InvalidNumber
        ConvertFromBase = ConvertFromBase + (OldBase ^ (i - 1)) * LetterVal
    Next i

    Exit Function
InvalidNumber:
    ConvertFromBase = 0
5 End Function

