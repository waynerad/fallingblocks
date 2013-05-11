VERSION 5.00
Begin VB.Form frmTetFE 
   Caption         =   "WayneTris"
   ClientHeight    =   1815
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   2070
   LinkTopic       =   "Form1"
   ScaleHeight     =   121
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   138
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   1680
      Top             =   2520
   End
End
Attribute VB_Name = "frmTetFE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetTickCount Lib "kernel32" () As Long

Dim m_iWidthBlks As Long
Dim m_iHeightBlks As Long

Dim m_iBlkSizePx As Long
Dim m_iBlocksPerPieceMin As Long
Dim m_iBlocksPerPieceMax As Long

Dim m_pCurrentPiece As cPiece
Dim m_iCurrentX As Long
Dim m_iCurrentY As Long

Dim m_iBoardColor As Long

Dim m_bRotateClockWise As Boolean

Dim m_bSquareOcc() As Boolean
Dim m_iSquareColor() As Long

Dim m_bInPlay As Boolean
Dim m_bInEvent As Boolean

Dim m_iLastTickMove As Long

Dim m_fTickInterval As Single
Dim m_iTickInterval As Long

Dim m_iGravitydx As Long
Dim m_iGravitydy As Long

Dim m_iLastTopBlock As Long

Dim m_iLastTickSpeedAdjust As Long

Dim m_fTargetTickInterval As Single
Dim m_fSpeedGradualizationScalarSlower As Single
Dim m_fSpeedGradualizationScalarFaster As Single

Dim m_bPaused As Boolean

Dim m_fPercentToTop As Single
Dim m_fBottomTickInterval As Single
Dim m_fTopTickInterval As Single

Dim m_iSpeedAdjustTickInterval As Long

Dim m_fSmallPieceSize As Single
Dim m_fBigPieceSize As Single

Dim m_iTicksLeftToGravityReverse As Long
Dim m_iTicksBetweenGravityReversal As Long

Dim m_iSquareTargetColor As Long
Dim m_fSquareTargetGradualizationFactor As Single

Dim m_fPercentToSubTop As Single

Dim m_bOptionsUp As Boolean
Dim m_pOptions As frmOptions

Dim m_iPieceCount As Long

Dim m_iGravityReversals As Long

Dim m_iZigDnom As Long
Dim m_iZigOffset As Long

Dim m_iZagDnom As Long
Dim m_iZagOffset As Long

Dim m_bShowScore As Boolean

Dim m_iMaxScore As Long

Function GetXOffset(x As Long, y As Long) As Long
    Dim dy As Long
    If (m_iZigDnom > 1) Then
        dy = y + m_iZigOffset
        dy = dy - (Int(dy / m_iZigDnom) * m_iZigDnom)
        GetXOffset = (dy * m_iBlkSizePx) / m_iZigDnom
    Else
        GetXOffset = 0
    End If
End Function

Function GetYOffset(x As Long, y As Long) As Long
    Dim dx As Long
    If (m_iZagDnom > 1) Then
        dx = x + m_iZagOffset
        dx = dx - (Int(dx / m_iZagDnom) * m_iZagDnom)
        GetYOffset = (dx * m_iBlkSizePx) / m_iZagDnom
    Else
        GetYOffset = 0
    End If
End Function

Function DrawBlock(x As Long, y As Long, iColor As Long)
    Dim xo As Long
    Dim yo As Long
    xo = GetXOffset(x, y)
    yo = GetYOffset(x, y)
    Me.Line (x * m_iBlkSizePx + xo, y * m_iBlkSizePx + yo)-(x * m_iBlkSizePx + m_iBlkSizePx - 1 + xo, y * m_iBlkSizePx + m_iBlkSizePx - 1 + yo), iColor, BF
End Function

Function EraseBlock(x As Long, y As Long)
    Me.Line (x * m_iBlkSizePx, y * m_iBlkSizePx)-(x * m_iBlkSizePx + m_iBlkSizePx - 1, y * m_iBlkSizePx + m_iBlkSizePx - 1), m_iBoardColor, BF
End Function

Sub TickSleep(n As Long)
    Dim iStartTick As Long
    Dim iNow As Long
    iStartTick = GetTickCount()
    iNow = GetTickCount()
    While (iNow < iStartTick + n)
        iNow = GetTickCount()
        If (iNow < iStartTick) Then
            Stop
            Exit Sub
        End If
    Wend
End Sub

Function ScrambleBoard()
    Dim x As Long
    Dim y As Long
    
    Dim i As Long
    For i = 0 To 9
    
    For y = 0 To m_iHeightBlks - 1
    For x = 0 To m_iWidthBlks - 1
        DrawBlock x, y, GenerateColor(Int(Rnd(1) * 3), -255#)
    Next x
    Next y
    TickSleep 10
    Next i
End Function

Function DrawBoardBorder()
    Me.Width = (m_iBlkSizePx * (m_iWidthBlks + 1)) * Screen.TwipsPerPixelX
    Me.height = (m_iBlkSizePx * (m_iHeightBlks + 1)) * Screen.TwipsPerPixelY + 400
    Me.Line (0, 0)-(m_iWidthBlks * m_iBlkSizePx + 3, m_iHeightBlks * m_iBlkSizePx + 3), 0, BF
    Me.Line (0, 0)-(m_iWidthBlks * m_iBlkSizePx - 1, m_iHeightBlks * m_iBlkSizePx - 1), m_iBoardColor, BF
End Function

Function DrawPiece(pp As cPiece, x As Long, y As Long)
    Dim i As Long
    Dim relx As Long
    Dim rely As Long
    Dim iColor As Long
    
    For i = 0 To pp.GetBlockCount() - 1
        relx = pp.GetRelX(i)
        rely = pp.GetRelY(i)
        iColor = pp.GetBlockColor(i)
        DrawBlock x + relx, y + rely, iColor
    Next i
End Function

Function ErasePiece(pp As cPiece, x As Long, y As Long)
    Dim i As Long
    Dim relx As Long
    Dim rely As Long
    Dim iColor As Long
    
    For i = 0 To pp.GetBlockCount() - 1
        relx = pp.GetRelX(i)
        rely = pp.GetRelY(i)
        iColor = pp.GetBlockColor(i)
        DrawBlock x + relx, y + rely, m_iBoardColor
    Next i
End Function

Function MoveCurrentPiece(dx As Long, dy As Long)
    ErasePiece m_pCurrentPiece, m_iCurrentX, m_iCurrentY
    m_iCurrentX = m_iCurrentX + dx
    m_iCurrentY = m_iCurrentY + dy
    DrawPiece m_pCurrentPiece, m_iCurrentX, m_iCurrentY
End Function

Function IsBlockSpaceEmpty(x As Long, y As Long) As Boolean
    If (x < 0) Then
        IsBlockSpaceEmpty = False
        Exit Function
    End If
    If (x >= m_iWidthBlks) Then
        IsBlockSpaceEmpty = False
        Exit Function
    End If
    If (y < 0) Then
        IsBlockSpaceEmpty = False
        Exit Function
    End If
    If (y >= m_iHeightBlks) Then
        IsBlockSpaceEmpty = False
        Exit Function
    End If
    IsBlockSpaceEmpty = (m_bSquareOcc(y * m_iWidthBlks + x) = False)
End Function

Function CanMovePiece(pp As cPiece, desirex As Long, desirey As Long)
    Dim i As Long
    Dim relx As Long
    Dim rely As Long
    For i = 0 To pp.GetBlockCount() - 1
        relx = pp.GetRelX(i)
        rely = pp.GetRelY(i)
        If (IsBlockSpaceEmpty(desirex + relx, desirey + rely) = False) Then
            CanMovePiece = False
            Exit Function
        End If
    Next i
    CanMovePiece = True
End Function

Sub TryToMoveCurrentPiece(dx As Long, dy As Long)
    If (CanMovePiece(m_pCurrentPiece, m_iCurrentX + dx, m_iCurrentY + dy)) Then
        MoveCurrentPiece dx, dy
    End If
End Sub

Sub TryToRotateCurrentPiece()
    Dim pp As cPiece
    Set pp = New cPiece
    pp.CopyPiece m_pCurrentPiece
    If (m_bRotateClockWise) Then
        pp.Rotate 0, -1
    Else
        pp.Rotate 0, 1
    End If
    If (CanMovePiece(pp, m_iCurrentX, m_iCurrentY)) Then
        ErasePiece m_pCurrentPiece, m_iCurrentX, m_iCurrentY
        Set m_pCurrentPiece = pp
        DrawPiece m_pCurrentPiece, m_iCurrentX, m_iCurrentY
    Else
        If (m_iGravitydy <> 0) Then
            If (CanMovePiece(pp, m_iCurrentX + 1, m_iCurrentY)) Then
                ErasePiece m_pCurrentPiece, m_iCurrentX, m_iCurrentY
                Set m_pCurrentPiece = pp
                m_iCurrentX = m_iCurrentX + 1
                DrawPiece m_pCurrentPiece, m_iCurrentX, m_iCurrentY
            Else
                If (CanMovePiece(pp, m_iCurrentX - 1, m_iCurrentY)) Then
                    ErasePiece m_pCurrentPiece, m_iCurrentX, m_iCurrentY
                    Set m_pCurrentPiece = pp
                    m_iCurrentX = m_iCurrentX - 1
                    DrawPiece m_pCurrentPiece, m_iCurrentX, m_iCurrentY
                End If
            End If
        Else
            If (CanMovePiece(pp, m_iCurrentX, m_iCurrentY + 1)) Then
                ErasePiece m_pCurrentPiece, m_iCurrentX, m_iCurrentY
                Set m_pCurrentPiece = pp
                m_iCurrentY = m_iCurrentY + 1
                DrawPiece m_pCurrentPiece, m_iCurrentX, m_iCurrentY
            Else
                If (CanMovePiece(pp, m_iCurrentX, m_iCurrentY + 1)) Then
                    ErasePiece m_pCurrentPiece, m_iCurrentX, m_iCurrentY
                    Set m_pCurrentPiece = pp
                    m_iCurrentY = m_iCurrentY - 1
                    DrawPiece m_pCurrentPiece, m_iCurrentX, m_iCurrentY - 1
                End If
            End If
        End If
    End If
End Sub

Sub GameOver()
    Dim x As Long
    Dim y As Long
    
    For y = 0 To m_iHeightBlks - 1
    For x = 0 To m_iWidthBlks - 1
        If (m_bSquareOcc(y * m_iWidthBlks + x) = False) Then
            m_iSquareColor(y * m_iWidthBlks + x) = 0
        End If
    Next x
    Next y
    RedrawBoard
    m_bInPlay = False
End Sub

Sub FlipBoard(iDirectionX As Long, iDirectionY As Long)
    Dim x As Long
    Dim y As Long
    
    Dim primey As Long
    Dim fromey As Long
    Dim checkey As Long
    
    Dim bInColumn As Boolean
    
    If (iDirectionX = 0) Then
        For x = 0 To m_iWidthBlks - 1
            bInColumn = False
            For y = 0 To m_iHeightBlks - 1
                If (m_bSquareOcc(y * m_iWidthBlks + x)) Then
                    bInColumn = True
                End If
            Next y
            If (bInColumn = True) Then
                If (iDirectionY > 0) Then
                    checkey = 0
                Else
                    checkey = m_iHeightBlks - 1
                End If
                While (m_bSquareOcc(checkey * m_iWidthBlks + x) = False)
                    y = 0
                    While (y <= m_iHeightBlks - 2)
                        If (iDirectionY > 0) Then
                            primey = y
                            fromey = y + 1
                        Else
                            primey = m_iHeightBlks - 1 - y
                            fromey = m_iHeightBlks - 2 - y
                        End If
                        m_bSquareOcc(primey * m_iWidthBlks + x) = m_bSquareOcc(fromey * m_iWidthBlks + x)
                        m_iSquareColor(primey * m_iWidthBlks + x) = m_iSquareColor(fromey * m_iWidthBlks + x)
                        If (m_bSquareOcc(primey * m_iWidthBlks + x)) Then
                            DrawBlock x, primey, m_iSquareColor(primey * m_iWidthBlks + x)
                        Else
                            EraseBlock x, primey
                        End If
                        y = y + 1
                    Wend
                    If (iDirectionY > 0) Then
                        y = m_iHeightBlks - 1
                    Else
                        y = 0
                    End If
                    m_bSquareOcc(y * m_iWidthBlks + x) = False
                    m_iSquareColor(y * m_iWidthBlks + x) = m_iBoardColor
                    EraseBlock x, y
                    TickSleep 1
                Wend
            End If
        Next x
    Else
        ' psycho reverse x and y here...
        For x = 0 To m_iHeightBlks - 1
            bInColumn = False
            For y = 0 To m_iWidthBlks - 1
                If (m_bSquareOcc(x * m_iWidthBlks + y)) Then
                    bInColumn = True
                End If
            Next y
            If (bInColumn = True) Then
                If (iDirectionX > 0) Then
                    checkey = 0
                Else
                    checkey = m_iWidthBlks - 1
                End If
                While (m_bSquareOcc(x * m_iWidthBlks + checkey) = False)
                    y = 0
                    While (y <= m_iWidthBlks - 2)
                        If (iDirectionX > 0) Then
                            primey = y
                            fromey = y + 1
                        Else
                            primey = m_iWidthBlks - 1 - y
                            fromey = m_iWidthBlks - 2 - y
                        End If
                        m_bSquareOcc(x * m_iWidthBlks + primey) = m_bSquareOcc(x * m_iWidthBlks + fromey)
                        m_iSquareColor(x * m_iWidthBlks + primey) = m_iSquareColor(x * m_iWidthBlks + fromey)
                        If (m_bSquareOcc(x * m_iWidthBlks + primey)) Then
                            DrawBlock primey, x, m_iSquareColor(x * m_iWidthBlks + primey)
                        Else
                            EraseBlock primey, x
                        End If
                        y = y + 1
                    Wend
                    If (iDirectionX > 0) Then
                        y = m_iWidthBlks - 1
                    Else
                        y = 0
                    End If
                    m_bSquareOcc(x * m_iWidthBlks + y) = False
                    m_iSquareColor(x * m_iWidthBlks + y) = m_iBoardColor
                    EraseBlock y, x
                    TickSleep 1
                Wend
            End If
        Next x
    End If
End Sub

Sub ReverseGravity()
    Dim iytemp As Long
    
    If (False) Then
        iytemp = m_iGravitydy
        m_iGravitydy = m_iGravitydx
        m_iGravitydx = -iytemp
        FlipBoard -m_iGravitydx, -m_iGravitydy
        If (m_iGravitydx <> 0) Then
            m_iGravityReversals = 8
        Else
            m_iGravityReversals = 0
        End If
    Else
        FlipBoard m_iGravitydx, m_iGravitydy
        m_iGravitydx = -m_iGravitydx
        m_iGravitydy = -m_iGravitydy
        m_iGravityReversals = m_iGravityReversals + 1
    End If
End Sub

Sub CreateNewCurrentPiece()
    Dim bKeepCreatingPieces As Boolean
    Dim bMakeSmallest As Boolean
    Dim iSize As Long
    Dim pprot1 As cPiece
    Dim pprot2 As cPiece
    Dim pprot3 As cPiece
    Dim bIsGood As Boolean
    Dim nexx As Long
    Dim nexy As Long
    bKeepCreatingPieces = True
    Dim fLBottom As Single
    Dim fLTop As Single
    Dim fLThis As Single
    
    If (m_iTicksLeftToGravityReverse < 0) Then
        m_iTicksLeftToGravityReverse = m_iTicksBetweenGravityReversal
        ReverseGravity
    End If
    
    Dim iCantCount As Long
    
    bMakeSmallest = False
    While (bKeepCreatingPieces)
        Set m_pCurrentPiece = New cPiece
        
        If (bMakeSmallest) Then
            iSize = 1
        Else
            If (m_fPercentToTop > 1#) Then
                iSize = 1
            Else
                fLBottom = Log(m_fSmallPieceSize)
                fLTop = Log(m_fBigPieceSize)
                fLThis = fLBottom * (1 - m_fPercentToTop) + (fLTop * m_fPercentToTop)
                fLThis = m_fBigPieceSize + m_fSmallPieceSize - Exp(fLThis)
            End If
            If (fLThis = Int(fLThis)) Then
                iSize = fLThis
            Else
                If (Rnd(1) < (fLThis - Int(fLThis))) Then
                    iSize = Int(fLThis) + 1
                Else
                    iSize = Int(fLThis)
                End If
            End If
        End If
        
        m_pCurrentPiece.CreateRandom iSize
        
        Set pprot1 = New cPiece
        pprot1.CopyPiece m_pCurrentPiece
        pprot1.Rotate 0, 1
        Set pprot2 = New cPiece
        pprot2.CopyPiece pprot1
        pprot2.Rotate 0, 1
        Set pprot3 = New cPiece
        pprot3.CopyPiece pprot2
        pprot3.Rotate 0, 1
        
        If (m_iGravitydx = 0) Then
            m_iCurrentX = Int(Rnd(1) * m_iWidthBlks)
            If (m_iGravitydy > 0) Then
                m_iCurrentY = m_iBlocksPerPieceMax  'Int(m_iHeightBlks / 2) 'm_iBlocksPerPiece
            Else
                m_iCurrentY = m_iHeightBlks - m_iBlocksPerPieceMax - 1
            End If
        Else
            m_iCurrentY = Int(Rnd(1) * m_iHeightBlks)
            If (m_iGravitydx > 0) Then
                m_iCurrentX = m_iBlocksPerPieceMax
            Else
                m_iCurrentX = m_iWidthBlks - m_iBlocksPerPieceMax - 1
            End If
        End If
        
        bIsGood = False
        nexx = m_iCurrentX - m_iGravitydx
        nexy = m_iCurrentY - m_iGravitydy
        
        While ((nexx > -m_iBlocksPerPieceMax) And (nexx < m_iWidthBlks + m_iBlocksPerPieceMax) And (nexy > -m_iBlocksPerPieceMax) And (nexy < m_iHeightBlks + m_iBlocksPerPieceMax))
            If (CanMovePiece(m_pCurrentPiece, nexx, nexy) And CanMovePiece(pprot1, nexx, nexy) And CanMovePiece(pprot2, nexx, nexy) And CanMovePiece(pprot3, nexx, nexy)) Then
                m_iCurrentX = nexx
                m_iCurrentY = nexy
                bIsGood = True
            End If
            nexx = nexx - m_iGravitydx
            nexy = nexy - m_iGravitydy
        Wend
        If (bIsGood = False) Then
            iCantCount = iCantCount + 1
            If (iSize = 1) Then
                If (iCantCount > 40) Then
                    GameOver
                    m_iCurrentY = -1000
                    Exit Sub
                End If
            Else
                bKeepCreatingPieces = True
                If (iCantCount > 20) Then
                    bMakeSmallest = True
                End If
            End If
        Else
            bKeepCreatingPieces = False
        End If
        m_iLastTickMove = GetTickCount()
    Wend
    m_iPieceCount = m_iPieceCount + 1
    If (m_iPieceCount > 1073741823) Then
        m_iPieceCount = 0
    End If
End Sub

Sub RedrawBoard()
    Dim x As Long
    Dim y As Long
    Dim fScore As Double
    
    If (m_bInPlay) Then
        For y = 0 To m_iHeightBlks - 1
        For x = 0 To m_iWidthBlks - 1
            DrawBlock x, y, m_iSquareColor(y * m_iWidthBlks + x)
        Next x
        Next y
        If (m_bShowScore) Then
            Me.CurrentX = 0
            Me.CurrentY = 0
            fScore = (CDbl(1) / m_fBottomTickInterval) * 10000
            Me.Print CLng(fScore)
            DrawSuccessBar fScore
        End If
    End If
End Sub

Sub ReduceColors()
    Dim x As Long
    Dim y As Long
    
    Dim iTargetRed As Long
    Dim iTargetGreen As Long
    Dim iTargetBlue As Long
    
    Dim iThisRed As Long
    Dim iThisGreen As Long
    Dim iThisBlue As Long
    
    Dim iThisColor As Long
    
    iTargetRed = m_iSquareTargetColor And 255
    iTargetGreen = (m_iSquareTargetColor And 65280) / 256
    iTargetBlue = (m_iSquareTargetColor And 16711680) / 65536
    
    For y = 0 To m_iHeightBlks - 1
    For x = 0 To m_iWidthBlks - 1
        If (m_bSquareOcc(y * m_iWidthBlks + x)) Then
            iThisColor = m_iSquareColor(y * m_iWidthBlks + x)
            iThisRed = iThisColor And 255
            iThisGreen = (iThisColor And 65280) / 256
            iThisBlue = (iThisColor And 16711680) / 65536
            iThisRed = (iThisRed * (1 - m_fSquareTargetGradualizationFactor)) + (iTargetRed * m_fSquareTargetGradualizationFactor)
            iThisGreen = (iThisGreen * (1 - m_fSquareTargetGradualizationFactor)) + (iTargetGreen * m_fSquareTargetGradualizationFactor)
            iThisBlue = (iThisBlue * (1 - m_fSquareTargetGradualizationFactor)) + (iTargetBlue * m_fSquareTargetGradualizationFactor)
            m_iSquareColor(y * m_iWidthBlks + x) = RGB(iThisRed, iThisGreen, iThisBlue)
        End If
    Next x
    Next y
End Sub

Sub DepositPiece(pp As cPiece, x As Long, y As Long)
    Dim i As Long
    Dim finalx As Long
    Dim finaly As Long
    Dim iColor As Long
    
    ReduceColors
    RedrawBoard
    For i = 0 To pp.GetBlockCount() - 1
        finalx = pp.GetRelX(i) + x
        finaly = pp.GetRelY(i) + y
        iColor = pp.GetBlockColor(i)
        DrawBlock finalx, finaly, iColor
        m_bSquareOcc(finaly * m_iWidthBlks + finalx) = True
        m_iSquareColor(finaly * m_iWidthBlks + finalx) = iColor
    Next i
End Sub

Sub RedrawCurrentPiece()
    If (m_bInPlay) Then
        DrawPiece m_pCurrentPiece, m_iCurrentX, m_iCurrentY
    End If
End Sub

Sub ZapRows()
    Dim x As Long
    Dim y As Long
    
    Dim bRowFull As Boolean
    Dim bBoardChanged As Boolean
    Dim j As Long
    
    Dim desty As Long
    Dim destx As Long
    
    bBoardChanged = False
    If (m_iGravitydx = 0) Then
        For y = 0 To m_iHeightBlks - 1
            bRowFull = True
            For x = 0 To m_iWidthBlks - 1
                If (m_bSquareOcc(y * m_iWidthBlks + x) = False) Then
                    bRowFull = False
                End If
            Next x
            If (bRowFull = True) Then
                bBoardChanged = True
                If (m_iGravitydy > 0) Then
                    desty = 1
                Else
                    desty = m_iHeightBlks - 2
                End If
                For j = y To desty Step -m_iGravitydy
                    For x = 0 To m_iWidthBlks - 1
                        m_bSquareOcc(j * m_iWidthBlks + x) = m_bSquareOcc((j - m_iGravitydy) * m_iWidthBlks + x)
                        m_iSquareColor(j * m_iWidthBlks + x) = m_iSquareColor((j - m_iGravitydy) * m_iWidthBlks + x)
                    Next x
                Next j
                If (m_iGravitydy > 0) Then
                    desty = 0
                Else
                    desty = (m_iHeightBlks - 1) * m_iWidthBlks
                End If
                For x = 0 To m_iWidthBlks - 1
                    m_bSquareOcc(desty + x) = False
                    m_iSquareColor(desty + x) = m_iBoardColor
                Next x
                y = y - 1
            End If
        Next y
    Else
        For x = 0 To m_iWidthBlks - 1
            bRowFull = True
            For y = 0 To m_iHeightBlks - 1
                If (m_bSquareOcc(y * m_iWidthBlks + x) = False) Then
                    bRowFull = False
                End If
            Next y
            If (bRowFull = True) Then
                bBoardChanged = True
                If (m_iGravitydx > 0) Then
                    destx = 1
                Else
                    destx = m_iWidthBlks - 2
                End If
                For j = x To destx Step -m_iGravitydx
                    For y = 0 To m_iHeightBlks - 1
                        m_bSquareOcc(y * m_iWidthBlks + j) = m_bSquareOcc(y * m_iWidthBlks + (j - m_iGravitydx))
                        m_iSquareColor(y * m_iWidthBlks + j) = m_iSquareColor(y * m_iWidthBlks + (j - m_iGravitydx))
                    Next y
                Next j
                If (m_iGravitydx > 0) Then
                    destx = 0
                Else
                    destx = m_iWidthBlks - 1
                End If
                For y = 0 To m_iHeightBlks - 1
                    m_bSquareOcc(y * m_iWidthBlks + destx) = False
                    m_iSquareColor(y * m_iWidthBlks + destx) = m_iBoardColor
                Next y
                x = x - 1
            End If
        Next x
    End If
    If (bBoardChanged) Then
        RedrawBoard
    End If
End Sub

Function FindHighestGettableRow() As Long
    Dim x As Long
    Dim y As Long
    
    Dim minx As Long
    Dim miny As Long
    Dim thasx As Long
    Dim thasy As Long
    
    Dim gx As Long
    Dim gy As Long
    
    If (m_iGravitydx >= 0) Then
        minx = m_iWidthBlks * m_iGravitydx
    Else
        minx = -1
    End If
    If (m_iGravitydy >= 0) Then
        miny = m_iHeightBlks * m_iGravitydy
    Else
        miny = -1
    End If
    For y = 0 To m_iHeightBlks - 1
    For x = 0 To m_iWidthBlks - 1
        If (m_bSquareOcc(y * m_iWidthBlks + x)) Then
            gx = x + m_iGravitydx
            gy = y + m_iGravitydy
            If ((gx >= 0) And (gx < m_iWidthBlks) And (gy >= 0) And (gy < m_iHeightBlks)) Then
                If (m_bSquareOcc(gy * m_iWidthBlks + gx) = False) Then
                    thasx = x * m_iGravitydx
                    thasy = y * m_iGravitydy
                    If (thasx < minx) Then
                        minx = thasx
                    End If
                    If (thasy < miny) Then
                        miny = thasy
                    End If
                End If
            End If
        Else
        End If
    Next x
    Next y
    If (m_iGravitydx = 0) Then
        If (m_iGravitydy > 0) Then
            FindHighestGettableRow = miny
        Else
            FindHighestGettableRow = m_iHeightBlks - 1 + miny
        End If
    Else
        If (m_iGravitydx > 0) Then
            FindHighestGettableRow = minx
        Else
            FindHighestGettableRow = m_iWidthBlks - 1 + minx
        End If
    End If
End Function

Function FindCurrentTopBlock() As Long
    Dim x As Long
    Dim y As Long

    Dim minx As Long
    Dim miny As Long
    Dim thasx As Long
    Dim thasy As Long

    If (m_iGravitydx >= 0) Then
        minx = m_iWidthBlks * m_iGravitydx
    Else
        minx = -1
    End If
    If (m_iGravitydy >= 0) Then
        miny = m_iHeightBlks * m_iGravitydy
    Else
        miny = -1
    End If
    For y = 0 To m_iHeightBlks - 1
    For x = 0 To m_iWidthBlks - 1
        If (m_bSquareOcc(y * m_iWidthBlks + x)) Then
            thasx = x * m_iGravitydx
            thasy = y * m_iGravitydy
            If (thasx < minx) Then
                minx = thasx
            End If
            If (thasy < miny) Then
                miny = thasy
            End If
        End If
    Next x
    Next y
    If (m_iGravitydx = 0) Then
        If (m_iGravitydy > 0) Then
            FindCurrentTopBlock = miny
        Else
            FindCurrentTopBlock = m_iHeightBlks - 1 + miny
        End If
    Else
        If (m_iGravitydx > 0) Then
            FindCurrentTopBlock = minx
        Else
            FindCurrentTopBlock = m_iWidthBlks - 1 + minx
        End If
    End If
End Function

Sub AdjustSpeed()
    Dim iNewTop As Long
    
    iNewTop = FindHighestGettableRow()
    
    If (m_iGravitydx = 0) Then
        m_fPercentToTop = (m_iHeightBlks - iNewTop) / m_iHeightBlks
    Else
        m_fPercentToTop = (m_iWidthBlks - iNewTop) / m_iWidthBlks
    End If
    
    Dim fLBottom As Single
    Dim fLTop As Single
    Dim fLThis As Single
    fLBottom = Log(1# / m_fBottomTickInterval)
    fLTop = Log(1# / m_fTopTickInterval)
    fLThis = fLBottom * (1 - m_fPercentToTop) + (fLTop * m_fPercentToTop)
    m_fTargetTickInterval = 1# / Exp(fLThis)

    If (m_fTickInterval < m_fTargetTickInterval) Then
        m_fTickInterval = ((1 - m_fSpeedGradualizationScalarSlower) * m_fTickInterval) + (m_fSpeedGradualizationScalarSlower * m_fTargetTickInterval)
    Else
        m_fTickInterval = ((1 - m_fSpeedGradualizationScalarFaster) * m_fTickInterval) + (m_fSpeedGradualizationScalarFaster * m_fTargetTickInterval)
    End If

    m_iTickInterval = Int(m_fTickInterval + 0.5)
    m_iLastTopBlock = iNewTop
    
    Dim fAbove As Single
    Dim fMult As Single
    fAbove = (m_fPercentToTop - 0.375) / 0.125
    fMult = Exp(fAbove * Log(1.0005))
    m_fBottomTickInterval = m_fBottomTickInterval * fMult
End Sub

Sub AdvancePiece()
    DepositPiece m_pCurrentPiece, m_iCurrentX, m_iCurrentY
    ZapRows
    CreateNewCurrentPiece
    DrawPiece m_pCurrentPiece, m_iCurrentX, m_iCurrentY
    AdjustSpeed
End Sub

Sub DropCurrentPiece()
    Dim dx As Long
    Dim dy As Long
    dx = m_iGravitydx
    dy = m_iGravitydy
    While (CanMovePiece(m_pCurrentPiece, m_iCurrentX + dx, m_iCurrentY + dy))
        dy = dy + m_iGravitydy
        dx = dx + m_iGravitydx
    Wend
    dy = dy - m_iGravitydy
    dx = dx - m_iGravitydx
    MoveCurrentPiece dx, dy
    AdvancePiece
End Sub

Sub InitBoard()
    ReDim m_bSquareOcc(m_iHeightBlks * m_iWidthBlks - 1) As Boolean
    ReDim m_iSquareColor(m_iHeightBlks * m_iWidthBlks - 1) As Long
    
    Dim x As Long
    Dim y As Long
    
    For y = 0 To m_iHeightBlks - 1
    For x = 0 To m_iWidthBlks - 1
        DrawBlock x, y, GenerateColor(Int(Rnd(1) * 3), -255#)
        m_bSquareOcc(y * m_iWidthBlks + x) = False
        m_iSquareColor(y * m_iWidthBlks + x) = m_iBoardColor
    Next x
    Next y
    m_bInPlay = True
End Sub

Sub StartGame()
    ScrambleBoard
    InitBoard
    DrawBoardBorder
    CreateNewCurrentPiece
    DrawPiece m_pCurrentPiece, m_iCurrentX, m_iCurrentY
    m_iLastTopBlock = 99999
    m_fTickInterval = 500
    m_iTickInterval = 2000
    m_iLastTickMove = 0
    m_bPaused = False
    m_iSpeedAdjustTickInterval = 200
    m_iBlocksPerPieceMax = Int(m_fBigPieceSize + 1)
    m_iTicksLeftToGravityReverse = m_iTicksBetweenGravityReversal
    m_iPieceCount = 0
    m_iZigDnom = 1 ' factors of 20: 1,20,10,5,4,2
    m_iZigOffset = 0
    m_iZagDnom = 1
    m_iZagOffset = 0
End Sub

Sub SetOptionVariables(fBigPieceSize As Single, fBottomTickInterval As Single, fSpeedGradualizationScalarFaster As Single)
    m_fBigPieceSize = fBigPieceSize
    m_fBottomTickInterval = fBottomTickInterval
    m_fSpeedGradualizationScalarSlower = fSpeedGradualizationScalarFaster
End Sub

Sub OptionsClosed()
    Set m_pOptions = Nothing
    m_bOptionsUp = False
End Sub

Sub DrawSuccessBar(fScore As Double)
    Dim y As Long
    Dim height As Long
    Dim maxs As Double
    
    maxs = CDbl(m_iMaxScore)
    
    height = (m_iHeightBlks * m_iBlkSizePx)  '+ m_iBlkSizePx - 1
    If (fScore > maxs) Then
        maxs = fScore
        m_iMaxScore = CLng(maxs + 0.5)
    End If
    y = ((maxs - fScore) / maxs) * height
    Me.Line ((m_iWidthBlks * m_iBlkSizePx) + 1, 0)-((m_iWidthBlks * m_iBlkSizePx) + 7, y - 1), RGB(255, 255, 255), BF
    Me.Line ((m_iWidthBlks * m_iBlkSizePx) + 1, y)-((m_iWidthBlks * m_iBlkSizePx) + 7, height), RGB(128, 128, 255), BF
End Sub

Sub EraseSuccessBar()
    Dim height As Long
    height = (m_iHeightBlks * m_iBlkSizePx) + m_iBlkSizePx - 1
    Me.Line ((m_iWidthBlks * m_iBlkSizePx) + 1, 0)-((m_iWidthBlks * m_iBlkSizePx) + 7, height), RGB(255, 255, 255), BF
End Sub

Sub DoTimerTick()
    Dim iTick As Long
    If (m_bInPlay) Then
        If (m_bInEvent = False) Then
            iTick = GetTickCount()
            If (iTick - m_iLastTickMove > m_iTickInterval) Then
                If (m_bPaused = False) Then
                    If (CanMovePiece(m_pCurrentPiece, m_iCurrentX + m_iGravitydx, m_iCurrentY + m_iGravitydy)) Then
                        MoveCurrentPiece m_iGravitydx, m_iGravitydy
                    Else
                        AdvancePiece
                    End If
                End If
                m_iLastTickMove = iTick
            End If
            If (iTick - m_iLastTickSpeedAdjust > m_iSpeedAdjustTickInterval) Then
                If (m_bPaused = False) Then
                    AdjustSpeed
                    m_iTicksLeftToGravityReverse = m_iTicksLeftToGravityReverse - m_iSpeedAdjustTickInterval
                End If
                m_iLastTickSpeedAdjust = iTick
            End If
        End If
    Else
        StartGame
    End If
End Sub

Sub LoadCustomData(sFile As String)
    Dim hIn As Long
    Dim sLine As String
    
    hIn = FreeFile()
    On Error GoTo russia
    Open sFile For Input As #hIn
        While (Not EOF(hIn))
            Line Input #hIn, sLine
                Select Case GetName(sLine)
                Case "iWidthBlks"
                    m_iWidthBlks = CLng(GetValue(sLine))
                Case "iHeightBlks"
                    m_iHeightBlks = CLng(GetValue(sLine))
                Case "iBlkSizePx"
                    m_iBlkSizePx = CLng(GetValue(sLine))
                Case "iBlocksPerPieceMin"
                    m_iBlocksPerPieceMin = CLng(GetValue(sLine))
                Case "iBlocksPerPieceMax"
                    m_iBlocksPerPieceMax = CLng(GetValue(sLine))
                Case "iBoardColor"
                    m_iBoardColor = CLng(GetValue(sLine))
                Case "bRotateClockWise"
                    m_bRotateClockWise = CBool(GetValue(sLine))
                Case "iGravitydx"
                    m_iGravitydx = CLng(GetValue(sLine))
                Case "iGravitydy"
                    m_iGravitydy = CLng(GetValue(sLine))
                Case "fSpeedGradualizationScalarFaster"
                    m_fSpeedGradualizationScalarSlower = CSng(GetValue(sLine))
                Case "fSpeedGradualizationScalarSlower"
                    m_fSpeedGradualizationScalarFaster = CSng(GetValue(sLine))
                Case "fBottomTickInterval"
                    m_fBottomTickInterval = CSng(GetValue(sLine))
                Case "fTopTickInterval"
                    m_fTopTickInterval = CSng(GetValue(sLine))
                Case "fSmallPieceSize"
                    m_fSmallPieceSize = CSng(GetValue(sLine))
                Case "fBigPieceSize"
                    m_fBigPieceSize = CSng(GetValue(sLine))
                Case "iTicksBetweenGravityReversal"
                    m_iTicksBetweenGravityReversal = CLng(GetValue(sLine))
                Case "iSquareTargetColor"
                    m_iSquareTargetColor = CLng(GetValue(sLine))
                Case "fSquareTargetGradualizationFactor"
                    m_fSquareTargetGradualizationFactor = CSng(GetValue(sLine))
                Case "iMaxScore"
                    m_iMaxScore = CLng(GetValue(sLine))
                Case Else
                    MsgBox "Error in waynetris.ini: unrecognized name/value pair" + Chr$(13) + Chr$(10) + sLine
                End Select
            Wend
    Close #hIn
russia:
End Sub

Sub SaveCustomData(sFile As String)
    Dim hOut As Long
    
    hOut = FreeFile()
    Open sFile For Output As #hOut
        Print #hOut, "fBottomTickInterval=" + CStr(m_fBottomTickInterval)
        Print #hOut, "iMaxScore=" + CStr(m_iMaxScore)
    Close #hOut
End Sub

Private Sub Form_DblClick()
    m_bInEvent = True
    StartGame
    m_bInEvent = False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    m_bInEvent = True
    Dim bMoveKey As Boolean
    Dim dx As Long
    Dim dy As Long
    Dim throwawy As Single
    If (KeyCode = 40) Then
        bMoveKey = True
        dx = 0
        dy = 1
    End If
    If (KeyCode = 37) Then
        bMoveKey = True
        dx = -1
        dy = 0
        throwawy = Rnd(1)
    End If
    If (KeyCode = 39) Then
        bMoveKey = True
        dx = 1
        dy = 0
    End If
    If (KeyCode = 38) Then
        bMoveKey = True
        dx = 0
        dy = -1
        throwawy = Rnd(1)
    End If
    If ((KeyCode = 80) Or (KeyCode = 76) Or (KeyCode = 82)) Then
        m_bPaused = (m_bPaused = False)
    End If
    If (bMoveKey) Then
        If ((m_bInPlay) And (m_bPaused = False)) Then
            If ((dx = m_iGravitydx) And (dy = m_iGravitydy)) Then
                TryToMoveCurrentPiece dx, dy
            Else
                If ((dx = -m_iGravitydx) And (dy = -m_iGravitydy)) Then
                    TryToRotateCurrentPiece
                Else
                    TryToMoveCurrentPiece dx, dy
                End If
            End If
        End If
    End If
    If (False) Then
        If ((KeyCode = 79) Or (KeyCode = 83)) Then
            If (m_bOptionsUp = False) Then
                Set m_pOptions = New frmOptions
                m_bOptionsUp = True
            End If
            m_pOptions.Show
            m_pOptions.SetValues Me, m_fBigPieceSize, m_fBottomTickInterval, m_fSpeedGradualizationScalarSlower
            m_bPaused = True
        End If
    End If
    If ((KeyCode = 79) Or (KeyCode = 83)) Then
        m_bShowScore = Not m_bShowScore
        If (m_bShowScore = False) Then
            EraseSuccessBar
        End If
    End If
    m_bInEvent = False
End Sub

Private Sub Form_Load()
    m_bInEvent = True
    m_iWidthBlks = 13 '11
    m_iHeightBlks = 20 '18
    m_iBlkSizePx = 20
    m_iBlocksPerPieceMin = 2
    m_iBlocksPerPieceMax = 4
    m_iBoardColor = 16777215
    m_bRotateClockWise = True
    m_iGravitydx = 0
    m_iGravitydy = 1
    m_fSpeedGradualizationScalarSlower = 0.02  '0.0175   '0.08
    m_fSpeedGradualizationScalarFaster = 1 '0.1 '0.002
    m_fBottomTickInterval = 2000 '37.64182 '75 '25
    m_fTopTickInterval = 3000
    m_fSmallPieceSize = 2#
    m_fBigPieceSize = 4.4 '4.33607 '5 ' normal 4.65 '4.5
    m_iTicksBetweenGravityReversal = 10800000 '300000 '
    m_iSquareTargetColor = RGB(128, 128, 192)
    m_fSquareTargetGradualizationFactor = 0.025
    m_iGravityReversals = 0
    m_bShowScore = True
    
    LoadCustomData App.Path + "\waynetris.ini"
    
    m_bInEvent = False
End Sub

Private Sub Form_Paint()
    m_bInEvent = True
    DrawBoardBorder
    RedrawBoard
    RedrawCurrentPiece
    m_bInEvent = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveCustomData App.Path + "\waynetris.ini"
End Sub

Private Sub Timer1_Timer()
    DoTimerTick
End Sub
