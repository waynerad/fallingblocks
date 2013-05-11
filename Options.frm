VERSION 5.00
Begin VB.Form frmOptions 
   Caption         =   "Options"
   ClientHeight    =   4065
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   4065
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.HScrollBar hsAdrenaline 
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   2760
      Width           =   4335
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Height          =   495
      Left            =   2040
      TabIndex        =   8
      Top             =   3480
      Width           =   2535
   End
   Begin VB.HScrollBar hsSpeed 
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   1680
      Width           =   4335
   End
   Begin VB.HScrollBar hsCcomplexity 
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   4335
   End
   Begin VB.Label lblSpeedGradFactorFaster 
      Caption         =   "Label12"
      Height          =   255
      Left            =   1680
      TabIndex        =   15
      Top             =   3120
      Width           =   1575
   End
   Begin VB.Label lblBottomTickInterval 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label11"
      Height          =   255
      Left            =   2040
      TabIndex        =   14
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label lblBigPieceSize 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label10"
      Height          =   255
      Left            =   1800
      TabIndex        =   13
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Caption         =   "More"
      Height          =   255
      Left            =   3120
      TabIndex        =   12
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label Label8 
      Caption         =   "Less"
      Height          =   255
      Left            =   480
      TabIndex        =   11
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label Label7 
      Caption         =   "Adrenaline Rush Factor - how much high speed is used"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   2520
      Width           =   4335
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "faster"
      Height          =   255
      Left            =   3480
      TabIndex        =   7
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "slower"
      Height          =   255
      Left            =   480
      TabIndex        =   6
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "Maximum Speed - makes game twitchy"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   1440
      Width           =   4215
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "higher"
      Height          =   255
      Left            =   3360
      TabIndex        =   4
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "lower"
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Complexity - makes game thinky"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   4335
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim m_pCallback As frmTetFE
Dim m_bInSetValues As Boolean

Sub SetValues(pCallback As frmTetFE, fBigPieceSize As Single, fBottomTickInterval As Single, fSpeedGradualizationScalarFaster As Single)
    m_bInSetValues = True
    Set m_pCallback = pCallback
    Me.hsCcomplexity.Value = CLng(((fBigPieceSize - 3#) / 3#) * 32767#)
    Me.hsSpeed.Value = CLng((1 - (Log(fBottomTickInterval) / Log(1000#))) * 32767#)
    Me.hsAdrenaline.Value = CLng(((Log(fSpeedGradualizationScalarFaster) - Log(1)) / (Log(0.0001) - Log(1))) * 32767#)
    
    Me.lblBigPieceSize.Caption = CStr(fBigPieceSize)
    Me.lblBottomTickInterval.Caption = CStr(fBottomTickInterval)
    Me.lblSpeedGradFactorFaster.Caption = CStr(fSpeedGradualizationScalarFaster)
    
    m_bInSetValues = False
End Sub

Sub SendValuesBack()
    Dim fBigPieceSize As Single
    Dim fBottomTickInterval As Single
    Dim fSpeedGradualizationScalarFaster As Single
    
    If (m_bInSetValues) Then
        Exit Sub
    End If
    
    fBigPieceSize = ((CSng(Me.hsCcomplexity.Value) / 32767#) * 3#) + 3#
    fBottomTickInterval = Exp(((-((CSng(Me.hsSpeed.Value) / 32767#) - 1)) * Log(1000)))
    fSpeedGradualizationScalarFaster = Exp(((CSng(Me.hsAdrenaline.Value) / 32767#) * (Log(0.0001) - Log(1))) + Log(1))
    
    m_pCallback.SetOptionVariables fBigPieceSize, fBottomTickInterval, fSpeedGradualizationScalarFaster

    Me.lblBigPieceSize.Caption = CStr(fBigPieceSize)
    Me.lblBottomTickInterval.Caption = CStr(fBottomTickInterval)
    Me.lblSpeedGradFactorFaster.Caption = CStr(fSpeedGradualizationScalarFaster)
End Sub

Private Sub cmdOk_Click()
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SendValuesBack
    m_pCallback.OptionsClosed
    Unload Me
End Sub

Private Sub hsAdrenaline_Change()
    SendValuesBack
End Sub

Private Sub hsCcomplexity_Change()
    SendValuesBack
End Sub

Private Sub hsSpeed_Change()
    SendValuesBack
End Sub


