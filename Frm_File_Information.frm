VERSION 5.00
Begin VB.Form Frm_File_Information 
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "Dateieigenschaften"
   ClientHeight    =   2385
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6945
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
   ScaleHeight     =   2385
   ScaleWidth      =   6945
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin VB.CommandButton Cmd_Ok 
      Caption         =   "OK"
      Height          =   375
      Left            =   5100
      TabIndex        =   7
      Top             =   150
      Width           =   1815
   End
   Begin VB.Frame Frame_Main 
      Height          =   2205
      Left            =   120
      TabIndex        =   0
      Top             =   30
      Width           =   4905
      Begin VB.Label Lab_Char_Number_Name 
         Alignment       =   1  'Rechts
         Height          =   210
         Left            =   1770
         TabIndex        =   6
         Top             =   1455
         Width           =   3045
      End
      Begin VB.Label Lab_Char_Number 
         AutoSize        =   -1  'True
         Caption         =   "Anzahl der Zeichen:"
         Height          =   210
         Left            =   120
         TabIndex        =   5
         Top             =   1455
         Width           =   1470
      End
      Begin VB.Label Lab_File_Size_Name 
         Alignment       =   1  'Rechts
         Height          =   210
         Left            =   1770
         TabIndex        =   4
         Top             =   1830
         Width           =   3045
      End
      Begin VB.Label Lab_File_Name_Name 
         Alignment       =   2  'Zentriert
         Height          =   1050
         Left            =   1770
         TabIndex        =   3
         Top             =   300
         Width           =   3045
         WordWrap        =   -1  'True
      End
      Begin VB.Label Lab_File_Size 
         Caption         =   "Größe (Kilo-Byte):"
         Height          =   210
         Left            =   120
         TabIndex        =   2
         Top             =   1830
         Width           =   1395
      End
      Begin VB.Label Lab_File_Name 
         AutoSize        =   -1  'True
         Caption         =   "Dateiname:"
         Height          =   210
         Left            =   120
         TabIndex        =   1
         Top             =   300
         Width           =   795
      End
   End
End
Attribute VB_Name = "Frm_File_Information"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Cmd_Ok_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Lab_File_Name_Name.Caption = m_Mdi.ActiveForm.Caption
    Lab_Char_Number_Name.Caption = Len(m_Mdi.ActiveForm.Rtf_Text.Text)
    Lab_File_Size_Name.Caption = Len(m_Mdi.ActiveForm.Rtf_Text.Text) / 1024
End Sub

