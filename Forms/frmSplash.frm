VERSION 5.00
Begin VB.Form frmSplash 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4410
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   5910
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   5910
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   810
      Top             =   660
   End
   Begin VB.Label lblTitulo 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Sistema de pedidos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   750
      Index           =   3
      Left            =   2115
      TabIndex        =   3
      Top             =   1545
      Width           =   4560
   End
   Begin VB.Label lblTitulo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DROPS.code"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   45
      TabIndex        =   2
      Top             =   45
      Width           =   1140
   End
   Begin VB.Label lblTitulo 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   1400
      TabIndex        =   1
      Top             =   4120
      Width           =   4440
   End
   Begin VB.Label lblTitulo 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "KDS Software"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   360
      Index           =   0
      Left            =   10
      TabIndex        =   0
      Top             =   3750
      Width           =   5835
   End
   Begin VB.Shape Barra 
      BackColor       =   &H00000000&
      BorderColor     =   &H00000000&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   2
      Left            =   0
      Top             =   0
      Width           =   2500
   End
   Begin VB.Shape Barra 
      BackColor       =   &H00C0C0C0&
      BorderColor     =   &H00C0C0C0&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   0
      Left            =   -120
      Top             =   4110
      Width           =   7095
   End
   Begin VB.Shape Barra 
      BackColor       =   &H00808080&
      BorderColor     =   &H00C0C0C0&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   1
      Left            =   0
      Top             =   0
      Width           =   7095
   End
   Begin VB.Image Image1 
      Height          =   4995
      Left            =   0
      Picture         =   "frmSplash.frx":0000
      Top             =   300
      Width           =   8835
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

On Error GoTo TrataErro

    lblTitulo(0).Caption = App.FileDescription
    lblTitulo(2).Caption = "Versão " & App.Major & "." & App.Minor
    Call VerTela
    Timer1.Enabled = True
    
TrataErro:
    If Err.Number <> 0 Then TrataErros
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmSplash = Nothing
End Sub

Private Sub Timer1_Timer()
On Error GoTo TrataErro
    DoEvents
    Load frmMain
    DoEvents
    frmMain.Show
    Unload frmSplash
TrataErro:
    If Err.Number <> 0 Then TrataErros
End Sub
