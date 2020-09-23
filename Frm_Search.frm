VERSION 5.00
Begin VB.Form Frm_Search 
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "Suchen"
   ClientHeight    =   2265
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9255
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
   ScaleHeight     =   2265
   ScaleWidth      =   9255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin VB.CommandButton Cmd_Cancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   7350
      TabIndex        =   9
      Top             =   600
      Width           =   1845
   End
   Begin VB.CommandButton Cmd_Ok 
      Caption         =   "Suchen"
      Height          =   375
      Left            =   7350
      TabIndex        =   8
      Top             =   150
      Width           =   1845
   End
   Begin VB.Frame Frame_Main 
      Caption         =   "Suchen"
      Height          =   2025
      Left            =   150
      TabIndex        =   3
      Top             =   60
      Width           =   7125
      Begin VB.Frame Frame_Search 
         Height          =   450
         Left            =   120
         TabIndex        =   5
         Top             =   180
         Width           =   1395
         Begin VB.Label Lab_Search 
            AutoSize        =   -1  'True
            Caption         =   "Suchen nach:"
            Height          =   210
            Left            =   90
            TabIndex        =   6
            Top             =   180
            Width           =   1005
         End
      End
      Begin VB.Frame Frame_Search_Text 
         Height          =   450
         Left            =   1560
         TabIndex        =   4
         Top             =   180
         Width           =   5445
         Begin VB.TextBox Txt_Search 
            Appearance      =   0  '2D
            BorderStyle     =   0  'Kein
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   60
            TabIndex        =   0
            Top             =   150
            Width           =   5325
         End
      End
      Begin VB.Frame Frame_Search_Options 
         Height          =   1275
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   6885
         Begin VB.CheckBox Chk_Match_Case 
            Caption         =   "Groß-/Kleinschreibung beachten"
            Height          =   210
            Left            =   180
            TabIndex        =   2
            Top             =   780
            Width           =   2715
         End
         Begin VB.CheckBox Chk_Match_Word 
            Caption         =   "Nur ganzes Wort"
            Height          =   210
            Left            =   180
            TabIndex        =   1
            Top             =   360
            Width           =   2085
         End
      End
   End
End
Attribute VB_Name = "Frm_Search"
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
    MsgBox Err.Description, vbCritical
End Sub

Private Sub Cmd_Ok_Click()
On Error GoTo Error
    Call s_Search_Text
Exit Sub
Error:
    MsgBox Err.Description, vbCritical
End Sub

Sub s_Search_Text()
On Error GoTo Error
    If Chk_Match_Word.Value = vbChecked And Chk_Match_Case.Value = _
        Unchecked Then
        m_Mode_Option = 0
    ElseIf Chk_Match_Case.Value = vbChecked And Chk_Match_Word.Value = _
        Unchecked Then
        m_Mode_Option = 1
    ElseIf Chk_Match_Case.Value = vbChecked And _
        Chk_Match_Word.Value = vbChecked Then
        m_Mode_Option = 2
    Else
        m_Mode_Option = 3
    End If
    m_Search_String = Txt_Search.Text
    If m_Search_Loop = 0 Then
        m_Start_Pos = 0
    ElseIf m_Search_Loop = 1 Then
        m_Start_Pos = m_Mdi.ActiveForm.Rtf_Text.SelStart + _
            Len(m_Mdi.ActiveForm.Rtf_Text.SelText)
    End If
    m_End_Pos = Len(m_Mdi.ActiveForm.Rtf_Text.Text)
    Select Case m_Mode_Option
        Case 0
            S_End = (m_Mdi.ActiveForm.Rtf_Text.Find( _
                m_Search_String, m_Start_Pos, m_End_Pos, rtfWholeWord))
                m_Search_Loop = 1
        Case 1
            S_End = (m_Mdi.ActiveForm.Rtf_Text.Find( _
                m_Search_String, m_Start_Pos, m_End_Pos, rtfMatchCase))
                m_Search_Loop = 1
        Case 2
            S_End = (m_Mdi.ActiveForm.Rtf_Text.Find( _
                m_Search_String, m_Start_Pos, m_End_Pos, rtfWholeWord + _
                    rtfMatchCase))
                m_Search_Loop = 1
        Case 3
            S_End = (m_Mdi.ActiveForm.Rtf_Text.Find( _
                m_Search_String, m_Start_Pos, m_End_Pos))
                m_Search_Loop = 1
    End Select
    If S_End = -1 Then
        m_Mdi.ActiveForm.Rtf_Text.SelStart = 0
        m_Mdi.ActiveForm.Rtf_Text.SelLength = 0
        m_Search_Loop = 0
        Cmd_Ok.Caption = "Suchen"
        Exit Sub
    End If
    Cmd_Ok.Caption = "Weitersuchen"
Exit Sub
Error:
    MsgBox Err.Description, vbCritical
End Sub

Private Sub Form_Load()
On Error GoTo Error
    Search_Loop = 0
    Dim lR As Long
    lR = SetTopMostWindow(Frm_Search.hwnd, True)
Exit Sub
Error:
    MsgBox Err.Description, vbCritical
End Sub
