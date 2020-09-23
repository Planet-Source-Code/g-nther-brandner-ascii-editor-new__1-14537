VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Frm_Document 
   ClientHeight    =   4530
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5565
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Frm_Document.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4530
   ScaleWidth      =   5565
   Begin VB.Frame Frame_Main 
      Height          =   3915
      Left            =   180
      TabIndex        =   1
      Top             =   240
      Width           =   5025
      Begin RichTextLib.RichTextBox Rtf_Text 
         Height          =   3525
         Left            =   180
         TabIndex        =   0
         Top             =   240
         Width           =   4725
         _ExtentX        =   8334
         _ExtentY        =   6218
         _Version        =   393217
         BorderStyle     =   0
         Enabled         =   -1  'True
         HideSelection   =   0   'False
         ScrollBars      =   2
         Appearance      =   0
         TextRTF         =   $"Frm_Document.frx":0442
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
End
Attribute VB_Name = "Frm_Document"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'//

Private Sub Form_Load()
On Error GoTo Error
    Call Form_Resize
Exit Sub
Error:
    MsgBox Err.Description, vbCritical
End Sub

Private Sub Form_Resize()
On Error Resume Next
    Frame_Main.Left = 100
    Frame_Main.Top = 150
    Frame_Main.Width = Me.ScaleWidth - 2 * Frame_Main.Left
    Frame_Main.Height = Me.ScaleHeight - 2 * Frame_Main.Top + 50
    Rtf_Text.Left = 50
    Rtf_Text.Top = 150
    Rtf_Text.Width = Frame_Main.Width - 2 * Rtf_Text.Left
    Rtf_Text.Height = Frame_Main.Height - Rtf_Text.Top - _
        Rtf_Text.Left
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo Error
    Dim m_Reply
    If m_W_Save = 1 Then
        Exit Sub
    End If
    m_Reply = MsgBox("Wollen Sie " & Me.Caption & " speichern?", vbQuestion _
        + vbYesNoCancel, "Speichern?")
    If m_Reply = vbYes Then
            If m_Mdi.ActiveForm.Rtf_Text.FileName <> "" Then
                m_Mdi.ActiveForm.Rtf_Text.SaveFile _
                    m_Mdi.ActiveForm.Rtf_Text.FileName, rtfText
            Else
                m_Mdi.Com_Dialog.Filter = "*.*|*.*"
                m_Mdi.Com_Dialog.ShowSave
                m_Save_Path = m_Mdi.Com_Dialog.FileName
                m_Mdi.ActiveForm.Rtf_Text.SaveFile m_Save_Path, rtfText
                m_Mdi.ActiveForm.Caption = m_Save_Path
                m_Mdi.ActiveForm.Rtf_Text.FileName = m_Save_Path
            End If
        Unload Me
    ElseIf m_Reply = vbNo Then
        Unload Me
    ElseIf m_Reply = vbCancel Then
        Cancel = 1
    End If
Exit Sub
Error:
    Cancel = 1
End Sub

Private Sub Rtf_Text_KeyPress(KeyAscii As Integer)
    Rtf_Text.SelColor = vbBlack
    Rtf_Text.SelBold = False
    
End Sub

