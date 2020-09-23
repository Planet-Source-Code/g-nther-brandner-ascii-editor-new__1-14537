VERSION 5.00
Begin VB.Form Frm_Print 
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "Drucken"
   ClientHeight    =   4845
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7980
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
   ScaleHeight     =   4845
   ScaleWidth      =   7980
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin VB.CommandButton Cmd_Cancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6090
      TabIndex        =   15
      Top             =   630
      Width           =   1845
   End
   Begin VB.CommandButton Cmd_Print 
      Caption         =   "Print"
      Height          =   375
      Left            =   6090
      TabIndex        =   14
      Top             =   150
      Width           =   1845
   End
   Begin VB.Frame Frame_Main 
      Caption         =   "Drucken"
      Height          =   4695
      Left            =   120
      TabIndex        =   6
      Top             =   60
      Width           =   5895
      Begin VB.CommandButton Cmd_Change_File 
         Caption         =   "Change file"
         Height          =   375
         Left            =   4140
         TabIndex        =   17
         Top             =   2790
         Width           =   1665
      End
      Begin VB.CommandButton Cmd_Printer_Setup 
         Caption         =   "Printer Setup"
         Height          =   375
         Left            =   4140
         TabIndex        =   16
         Top             =   4230
         Width           =   1665
      End
      Begin VB.Frame Frame_Print_Head 
         Height          =   525
         Left            =   90
         TabIndex        =   2
         Top             =   2190
         Width           =   5715
         Begin VB.TextBox Txt_Print_Head 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'Kein
            Enabled         =   0   'False
            Height          =   210
            Left            =   1530
            TabIndex        =   1
            Top             =   210
            Width           =   4155
         End
         Begin VB.CheckBox Chk_Print_Head 
            Caption         =   "Kopfzeile"
            Height          =   210
            Left            =   90
            TabIndex        =   0
            Top             =   210
            Width           =   1035
         End
      End
      Begin VB.Frame Frame_Information 
         Height          =   525
         Left            =   90
         TabIndex        =   3
         Top             =   3660
         Width           =   5715
         Begin VB.Label Label_Information 
            Alignment       =   2  'Zentriert
            AutoSize        =   -1  'True
            Caption         =   "Um den aktuellen Drucker zu änderen, klicken Sie auf 'Drucker einrichten'!"
            Height          =   210
            Left            =   180
            TabIndex        =   4
            Top             =   210
            Width           =   5280
         End
      End
      Begin VB.Frame Frame_File 
         Height          =   525
         Left            =   90
         TabIndex        =   13
         Top             =   960
         Width           =   1815
         Begin VB.Label Label_File 
            AutoSize        =   -1  'True
            Caption         =   "Datei:"
            Height          =   210
            Left            =   90
            TabIndex        =   5
            Top             =   210
            Width           =   405
         End
      End
      Begin VB.Frame Frame_File_Name 
         Height          =   525
         Left            =   2100
         TabIndex        =   10
         Top             =   960
         Width           =   3705
         Begin VB.Label Label_File_Name 
            Height          =   225
            Left            =   90
            TabIndex        =   12
            Top             =   210
            Width           =   3495
         End
      End
      Begin VB.Frame Frame_Printer_Name 
         Height          =   525
         Left            =   2100
         TabIndex        =   9
         Top             =   210
         Width           =   3705
         Begin VB.Label Label_Printer_Name 
            Height          =   210
            Left            =   90
            TabIndex        =   11
            Top             =   210
            Width           =   3525
         End
      End
      Begin VB.Frame Frame_Printer 
         Height          =   525
         Left            =   90
         TabIndex        =   7
         Top             =   210
         Width           =   1815
         Begin VB.Label Label_Printer 
            AutoSize        =   -1  'True
            Caption         =   "Drucker:"
            Height          =   210
            Left            =   90
            TabIndex        =   8
            Top             =   210
            Width           =   615
         End
      End
   End
End
Attribute VB_Name = "Frm_Print"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'//

Private Sub Chk_Print_Head_Click()
On Error GoTo Error
    If Chk_Print_Head.Value = vbChecked Then
        Txt_Print_Head.BackColor = vbWhite
        Txt_Print_Head.Enabled = True
    Else
        Txt_Print_Head.BackColor = &H8000000F
        Txt_Print_Head.Enabled = False
    End If
Exit Sub
Error:
    MsgBox Err.Description, vbCritical
End Sub

Private Sub Cmd_Change_File_Click()
On Error GoTo Error
    Load Frm_Change_File
    Frm_Change_File.Show (1)
Exit Sub
Error:
    MsgBox Err.Description, vbCritical
End Sub

Private Sub cmd_printer_setup_Click()
On Error GoTo Error
    m_Mdi.Com_Dialog.ShowPrinter
    Unload Me
    Load Frm_Print
    Frm_Print.Show (1)
Exit Sub
Error:
End Sub

Private Sub Cmd_Cancel_Click()
On Error GoTo Error
    Unload Me
Exit Sub
Error:
    MsgBox Err.Description, vbCritical
End Sub

Private Sub Cmd_Print_Click()
On Error GoTo Error
    If Chk_Print_Head.Value = vbChecked Then
        Dim m_Print_String As String
        m_Print_String = Txt_Print_Head.Text
        If Printer.TextWidth(m_Print_String) > Printer.Width Then
            MsgBox "Kopfzeile zu lang!", vbInformation
            Exit Sub
        End If
        If Len(m_Print_String) > 0 Then
            Printer.CurrentY = 200
            Printer.CurrentX = Printer.Width / 2 - _
            Printer.TextWidth(m_Print_String)
            Printer.Print m_Print_String
        End If
        Printer.Line (200, 700)-(Printer.Width - 400, 700)
    End If
    Printer.CurrentY = 1000
    Printer.CurrentX = 200
    Printer.Print m_Mdi.ActiveForm.Rtf_Text.Text
    Printer.NewPage
    Unload Me
Exit Sub
Error:
    MsgBox Err.Description, vbCritical
End Sub

Private Sub Cmd_Pritner_Setup_Click()

End Sub

Private Sub Form_Load()
On Error GoTo Error
    Label_Printer_Name.Caption = Printer.DeviceName
    Label_File_Name.Caption = m_Mdi.ActiveForm.Caption
Exit Sub
Error:
    MsgBox Err.Description, vbCritical
End Sub

