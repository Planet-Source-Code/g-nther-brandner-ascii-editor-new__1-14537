VERSION 5.00
Begin VB.Form Frm_Date_Time 
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "Datum/Uhrzeit"
   ClientHeight    =   2670
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6435
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
   ScaleHeight     =   2670
   ScaleWidth      =   6435
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin VB.CommandButton Cmd_Ok 
      Caption         =   "OK"
      Height          =   375
      Left            =   4590
      TabIndex        =   3
      Top             =   180
      Width           =   1815
   End
   Begin VB.CommandButton Cmd_Actual_Time 
      Caption         =   "Actual Time"
      Height          =   405
      Left            =   4590
      TabIndex        =   2
      Top             =   630
      Width           =   1815
   End
   Begin VB.Frame Frame_Main 
      Caption         =   "Uhrzeit"
      Height          =   2445
      Left            =   150
      TabIndex        =   0
      Top             =   60
      Width           =   4395
      Begin VB.Label Lab_Time 
         Alignment       =   2  'Zentriert
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   48
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1185
         Left            =   120
         TabIndex        =   1
         Top             =   660
         Width           =   4185
      End
   End
End
Attribute VB_Name = "Frm_Date_Time"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Cmd_Actual_Time_Click()
    Lab_Time.Caption = Time
End Sub

Private Sub Cmd_Ok_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Lab_Time.Caption = Time
End Sub
