VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPiece"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim m_iBlockCount As Long
Dim m_iBlockRelX() As Long
Dim m_iBlockRelY() As Long
Dim m_iBlockColor() As Long
Dim m_fBlockBrightness() As Single

Sub CreateRandom(iBlocks As Long)
    m_iBlockCount = iBlocks
    ReDim m_iBlockRelX(iBlocks - 1) As Long
    ReDim m_iBlockRelY(iBlocks - 1) As Long
    ReDim m_iBlockColor(iBlocks - 1) As Long
    ReDim m_fBlockBrightness(iBlocks - 1) As Single
    
    Dim i As Long
    Dim x As Long
    Dim y As Long
    
    x = 0
    y = 0
    
    Dim dx As Long
    Dim dy As Long
    Dim r As Long
    
    Dim pendx As Long
    Dim pendy As Long
    
    Dim bNotOK As Boolean
    Dim bNotFound As Boolean
    Dim j As Long
    Dim iStuckLoop As Long
    
    Dim iBlackCol As Long
    Dim fLastBrightness As Single
    
    Dim colornum As Long
    
    iBlackCol = Int(Rnd(1) * 3)
    For i = 0 To iBlocks - 1
        m_iBlockRelX(i) = x
        m_iBlockRelY(i) = y
'        m_iBlockColor(i) = GenerateColor(iBlackCol, fLastBrightness)
        colornum = i - Int(i / 4) * 4
        If (colornum = 0) Then
            m_iBlockColor(i) = RGB(255, 0, 0)
        End If
        If (colornum = 1) Then
            m_iBlockColor(i) = RGB(0, 0, 0)
        End If
        If (colornum = 2) Then
            m_iBlockColor(i) = RGB(255, 128, 0)
        End If
        If (colornum = 3) Then
            m_iBlockColor(i) = RGB(192, 0, 64)
        End If
        m_fBlockBrightness(i) = fLastBrightness
        bNotOK = True
        iStuckLoop = 0
        While (bNotOK)
            r = Int(Rnd(1) * 4)
            If (r = 0) Then
                dy = -1
                dx = 0
            Else
                If (r = 1) Then
                    dy = 0
                    dx = 1
                Else
                    If (r = 2) Then
                        dy = 1
                        dx = 0
                    Else
                        If (r = 3) Then
                            dy = 0
                            dx = -1
                        Else
                            Stop
                        End If
                    End If
                End If
            End If
            
            If (i > 0) Then
                j = Int(Rnd(1) * (i + 1))
                x = m_iBlockRelX(j)
                y = m_iBlockRelY(j)
                fLastBrightness = m_fBlockBrightness(j)
            End If
            
            pendx = x + dx
            pendy = y + dy
            bNotFound = True
            For j = 0 To i
                If ((m_iBlockRelX(j) = pendx) And (m_iBlockRelY(j) = pendy)) Then
                    bNotFound = False
                End If
            Next j
            If (bNotFound) Then
                bNotOK = False
            End If
            iStuckLoop = iStuckLoop + 1
            If (iStuckLoop = 100) Then
                Beep
                bNotOK = False
                pendx = 0
                pendy = 0
                i = -1
            End If
        Wend
            x = pendx
            y = pendy
    Next i
    Dim cmx As Long
    Dim cmy As Long
    For i = 0 To iBlocks - 1
        cmx = cmx + m_iBlockRelX(i)
        cmy = cmy + m_iBlockRelY(i)
    Next i
    cmx = Int(cmx / iBlocks)
    cmy = Int(cmy / iBlocks)
    For i = 0 To iBlocks - 1
        m_iBlockRelX(i) = m_iBlockRelX(i) - cmx
        m_iBlockRelY(i) = m_iBlockRelY(i) - cmy
    Next i
End Sub

Function GetBlockCount() As Long
    GetBlockCount = m_iBlockCount
End Function

Function GetRelX(i As Long) As Long
    GetRelX = m_iBlockRelX(i)
End Function

Function GetRelY(i As Long) As Long
    GetRelY = m_iBlockRelY(i)
End Function

Function GetBlockColor(i As Long) As Long
    GetBlockColor = m_iBlockColor(i)
End Function

Sub CopyPiece(pp As cPiece)
    m_iBlockCount = pp.GetBlockCount()
    ReDim m_iBlockRelX(m_iBlockCount - 1) As Long
    ReDim m_iBlockRelY(m_iBlockCount - 1) As Long
    ReDim m_iBlockColor(m_iBlockCount - 1) As Long
    Dim i As Long
    For i = 0 To m_iBlockCount - 1
        m_iBlockRelX(i) = pp.GetRelX(i)
        m_iBlockRelY(i) = pp.GetRelY(i)
        m_iBlockColor(i) = pp.GetBlockColor(i)
    Next i
End Sub

Function Rotate(a As Long, b As Long)
    Dim i As Long
    Dim nux As Long
    Dim nuy As Long
    
    For i = 0 To m_iBlockCount - 1
        '(y+ix)(a+bi)
        '=ay + ixa + biy + biix
        '=ay-bx + (xa+by)i
        nuy = a * m_iBlockRelY(i) - b * m_iBlockRelX(i)
        nux = a * m_iBlockRelX(i) + b * m_iBlockRelY(i)
        m_iBlockRelY(i) = nuy
        m_iBlockRelX(i) = nux
    Next i
End Function
