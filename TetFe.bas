Attribute VB_Name = "modTetFe"
Option Explicit

Function GenerateColor(iBlackCol As Long, fLastTarget As Single) As Long
    Dim iRed As Long
    Dim iGreen As Long
    Dim iBlue As Long
    
    'iRed = Int(Rnd(1) * 255)
    'iGreen = Int(Rnd(1) * 255)
    'iBlue = Int(Rnd(1) * 255)
    
    'Dim iBlackCol As Long
    Dim fSatCol As Single
    'iBlackCol = Int(Rnd(1) * 3)
    fSatCol = Int(Rnd(1) * 2)
    If (iBlackCol = 1) Then
        iRed = 0
        If (fSatCol = 1) Then
            iGreen = 255
            iBlue = Int(Rnd(1) * 255)
        Else
            'blue
            iBlue = 255
            iGreen = Int(Rnd(1) * 255)
        End If
    Else
        If (iBlackCol = 2) Then
            iGreen = 0
            If (fSatCol = 1) Then
                'red
                iRed = 255
                iBlue = Int(Rnd(1) * 255)
            Else
                'blue
                iBlue = 255
                iRed = Int(Rnd(1) * 255)
            End If
        Else
            iBlue = 0
            If (fSatCol = 1) Then
                'red
                iRed = 255
                iGreen = Int(Rnd(1) * 255)
            Else
                'green
                iGreen = 255
                iRed = Int(Rnd(1) * 255)
            End If
        End If
    End If
    
'    GenerateColor = RGB(iRed, iGreen, iBlue)
'    Exit Sub
    
    Dim fBrite As Single
    'fBrite = ((Red value X 299) + (Green value X 587) + (Blue value X 114)) / 1000
    fBrite = CSng(iRed) * 0.299 + CSng(iGreen) * 0.587 + CSng(iBlue) * 0.114
    
    Dim fTarget As Single
    
    fTarget = 64# + (Rnd(1) * 128#)
    While (Abs(fTarget - fLastTarget) < 60)
        fTarget = 64# + (Rnd(1) * 128#)
    Wend
    
    Dim fScale As Single
    fScale = fTarget / fBrite
    
    iRed = iRed * fScale
    iBlue = iBlue * fScale
    iGreen = iGreen * fScale
    
    GenerateColor = RGB(iRed, iGreen, iBlue)
    fLastTarget = fTarget
End Function

Function GetName(sNameValuePair As String) As String
    Dim i As Long
    i = InStr(sNameValuePair, "=")
    GetName = Left$(sNameValuePair, i - 1)
End Function

Function GetValue(sNameValuePair As String) As String
    Dim i As Long
    i = InStr(sNameValuePair, "=")
    GetValue = Mid$(sNameValuePair, i + 1)
End Function

