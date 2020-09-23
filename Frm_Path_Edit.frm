VERSION 5.00
Begin VB.Form Frm_Path_Edit 
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "Pfadangaben"
   ClientHeight    =   2415
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6300
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
   ScaleHeight     =   2415
   ScaleWidth      =   6300
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin VB.CommandButton Cmd_Cancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4410
      TabIndex        =   6
      Top             =   660
      Width           =   1845
   End
   Begin VB.CommandButton Cmd_Ok 
      Caption         =   "OK"
      Height          =   375
      Left            =   4410
      TabIndex        =   5
      Top             =   210
      Width           =   1845
   End
   Begin VB.Frame Frame_Main 
      Caption         =   "Pfadangaben"
      Height          =   2205
      Left            =   120
      TabIndex        =   2
      Top             =   90
      Width           =   4245
      Begin VB.TextBox txt_Path_Pictures 
         Height          =   315
         Left            =   480
         TabIndex        =   1
         Top             =   1650
         Width           =   3555
      End
      Begin VB.TextBox txt_Path_Open_Save 
         Height          =   315
         Left            =   480
         TabIndex        =   0
         Top             =   780
         Width           =   3555
      End
      Begin VB.Label Lab_Two 
         AutoSize        =   -1  'True
         Caption         =   "Aktueller Pfad für Bilder (nur bei Html-Dateien)"
         Height          =   210
         Left            =   150
         TabIndex        =   4
         Top             =   1290
         Width           =   3315
      End
      Begin VB.Label Lab_One 
         AutoSize        =   -1  'True
         Caption         =   "Aktueller Pfad für Öffnen/Speichern-Dialog"
         Height          =   210
         Left            =   150
         TabIndex        =   3
         Top             =   420
         Width           =   3090
      End
   End
End
Attribute VB_Name = "Frm_Path_Edit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Cmd_Cancel_Click()
    Unload Me
End Sub

Private Sub Cmd_Ok_Click()
    SaveSetting "Ascii-Editor", "Attitudes", "Active_Dir", txt_Path_Open_Save.Text
    SaveSetting "Ascii-Editor", "Attitudes", "Active_Pic", _
        txt_Path_Pictures.Text
    Unload Me
End Sub

Private Sub Form_Load()
    txt_Path_Open_Save.Text = GetSetting("Ascii-Editor", _
        "Attitudes", "Active_Dir", "C:\")
    txt_Path_Pictures.Text = GetSetting("Ascii-Editor", _
        "Attitudes", "Active_Pic", "C:\")
End Sub
