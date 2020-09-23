VERSION 5.00
Begin VB.Form Frm_Color_Attitudes 
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "Farbeinstellungen für farbigen Html-Code"
   ClientHeight    =   1785
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5895
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1785
   ScaleWidth      =   5895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin VB.CommandButton Cmd_Cancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4020
      TabIndex        =   5
      Top             =   540
      Width           =   1815
   End
   Begin VB.CommandButton Cmd_Ok 
      Caption         =   "OK"
      Height          =   375
      Left            =   4020
      TabIndex        =   4
      Top             =   120
      Width           =   1815
   End
   Begin VB.Frame Frame_Main 
      Caption         =   "Farbeinstellungen"
      Height          =   1545
      Left            =   150
      TabIndex        =   2
      Top             =   90
      Width           =   3795
      Begin VB.CheckBox Chk_Bold 
         Caption         =   "Html-Klammern in fetter Schriftart"
         Height          =   255
         Left            =   150
         TabIndex        =   1
         Top             =   840
         Width           =   2955
      End
      Begin VB.Label Cmd_Color 
         Height          =   285
         Left            =   2610
         TabIndex        =   0
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Farbe der Html-Klammern (<>):"
         Height          =   210
         Left            =   180
         TabIndex        =   3
         Top             =   510
         Width           =   2205
      End
   End
End
Attribute VB_Name = "Frm_Color_Attitudes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Cmd_Cancel_Click()
On Error GoTo Error
    Unload Me
Exit Sub
Error:
End Sub

Private Sub Cmd_Color_Click()
On Error GoTo Error
    Mdi_Frm_Main.Com_Dialog.ShowColor
    Cmd_Color.BackColor = Mdi_Frm_Main.Com_Dialog.Color
Exit Sub
Error:
End Sub

Private Sub Cmd_Ok_Click()
On Error GoTo Error
    m_Color = Cmd_Color.BackColor
    If Chk_Bold.Value = vbChecked Then
        m_Tags_Bold = True
    Else
        m_Tags_Bold = False
    End If
    Unload Me
Exit Sub
Error:
End Sub

Private Sub Form_Load()
On Error GoTo Error
    Cmd_Color.BackColor = m_Color
    If m_Tags_Bold = True Then
        Chk_Bold.Value = vbChecked
    Else
        Chk_Bold.Value = vbUnchecked
    End If
Exit Sub
Error:
End Sub
