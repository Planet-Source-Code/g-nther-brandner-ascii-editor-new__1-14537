VERSION 5.00
Begin VB.Form Frm_Info 
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "Info"
   ClientHeight    =   2610
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6570
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
   ScaleHeight     =   2610
   ScaleWidth      =   6570
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin VB.CommandButton Cmd_Ok 
      Caption         =   "OK"
      Height          =   375
      Left            =   4710
      TabIndex        =   2
      Top             =   150
      Width           =   1815
   End
   Begin VB.Frame Frame_Main 
      Caption         =   "Info"
      Height          =   2475
      Left            =   90
      TabIndex        =   0
      Top             =   30
      Width           =   4575
      Begin VB.Label Label1 
         Alignment       =   2  'Zentriert
         Caption         =   "Günther Brandner 2000 - executer@aon.at"
         Height          =   465
         Left            =   90
         TabIndex        =   1
         Top             =   1080
         Width           =   4335
         WordWrap        =   -1  'True
      End
   End
End
Attribute VB_Name = "Frm_Info"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'//

Private Sub Cmd_Ok_Click()
On Error GoTo Error
    Unload Me
Exit Sub
Error:
    MsgBox Err.Description, vbCritical
End Sub

