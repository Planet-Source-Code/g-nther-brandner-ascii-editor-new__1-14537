VERSION 5.00
Begin VB.Form Frm_Opened_Files 
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "Zuletzt geöffnete Dateien"
   ClientHeight    =   2640
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6495
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
   ScaleHeight     =   2640
   ScaleWidth      =   6495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin VB.CommandButton Cmd_Cancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4650
      TabIndex        =   3
      Top             =   600
      Width           =   1815
   End
   Begin VB.CommandButton Cmd_Ok 
      Caption         =   "OK"
      Height          =   375
      Left            =   4650
      TabIndex        =   2
      Top             =   150
      Width           =   1815
   End
   Begin VB.Frame Frame_Main 
      Height          =   2505
      Left            =   120
      TabIndex        =   1
      Top             =   30
      Width           =   4485
      Begin VB.ListBox List_Opened_Files 
         Height          =   2160
         Left            =   90
         TabIndex        =   0
         Top             =   210
         Width           =   4305
      End
   End
End
Attribute VB_Name = "Frm_Opened_Files"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Cmd_Cancel_Click()
    Unload Me
End Sub

Private Sub Cmd_Ok_Click()
    Unload Me
End Sub

