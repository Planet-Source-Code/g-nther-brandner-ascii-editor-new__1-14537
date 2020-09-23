VERSION 5.00
Begin VB.Form Frm_Change_File 
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "Datei wechseln"
   ClientHeight    =   3045
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
   ScaleHeight     =   3045
   ScaleWidth      =   6435
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin VB.CommandButton Cmd_Cancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4560
      TabIndex        =   3
      Top             =   720
      Width           =   1815
   End
   Begin VB.CommandButton Cmd_Ok 
      Caption         =   "OK"
      Height          =   375
      Left            =   4560
      TabIndex        =   2
      Top             =   240
      Width           =   1815
   End
   Begin VB.Frame Frame_Main 
      Caption         =   "Geöffnete Fenster"
      Height          =   2805
      Left            =   120
      TabIndex        =   1
      Top             =   90
      Width           =   4395
      Begin VB.ListBox List_Opened_Windows 
         Height          =   2370
         Left            =   150
         TabIndex        =   0
         Top             =   300
         Width           =   4125
      End
   End
End
Attribute VB_Name = "Frm_Change_File"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'//

Private Sub Cmd_Cancel_Click()
On Error GoTo Error
    Unload Me
Exit Sub
Error:
End Sub

Private Sub Cmd_Ok_Click()
On Error GoTo Error
    m_Window_Index = List_Opened_Windows.ListIndex + 1
    Call Change_Window
Exit Sub
Error:
    MsgBox Err.Description, vbCritical
End Sub

Private Sub Form_Load()
On Error GoTo Error
    For i = 1 To Forms.Count - 3
        List_Opened_Windows.AddItem i & ": " & Forms(i).Caption
    Next
Exit Sub
Error:
    MsgBox Err.Description, vbCritical
End Sub

Private Sub List_Opened_Windows_DblClick()
On Error GoTo Error
    Call Cmd_Ok_Click
Exit Sub
Error:
    MsgBox Err.Description, vbCritical
End Sub
