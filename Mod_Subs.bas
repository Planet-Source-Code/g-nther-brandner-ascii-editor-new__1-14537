Attribute VB_Name = "Mod_Subs"
Option Explicit

Sub Main()
On Error GoTo Error
    Set m_Mdi = Mdi_Frm_Main
    Load m_Mdi
    m_Mdi.Show
    m_W_Save = 0
Exit Sub
Error:
    MsgBox Err.Description, vbCritical
End Sub

Sub Change_Window()
On Error GoTo Error
    Unload Frm_Change_File
    Unload Frm_Print
    Forms(m_Window_Index).SetFocus
    Load Frm_Print
    Frm_Print.Show (1)
Exit Sub
Error:
    MsgBox Err.Description, vbCritical
End Sub

Sub Sub_Path_Control()
    Dim i, x
    For i = 1 To Len(m_Save_Path)
        x = Mid(m_Save_Path, i, 1)
        If x = "." Then
            m_Char_Index = 0
        End If
        m_Char_Index = m_Char_Index + 1
    Next
End Sub

Sub Sub_Color_Html()
On Error GoTo Error
    If m_Mdi.ActiveForm Is Nothing Then Exit Sub
    Dim r, x, i
    i = 0
    Set r = m_Mdi.ActiveForm.Rtf_Text
    Dim Str_String
    Str_String = String(Len(r.Text), 0)
    Str_String = r.Text
    r.Visible = False
    Do Until i = Len(Str_String)
    DoEvents
        x = Mid(Str_String, i + 1, 1)
        If x = "<" Then
            r.SelStart = i
            r.SelLength = 1
            r.SelColor = m_Color
            If m_Tags_Bold = True Then
                r.SelBold = True
            End If
        ElseIf x = ">" Then
            r.SelStart = i
            r.SelLength = 1
            r.SelColor = m_Color
            If m_Tags_Bold = True Then
                r.SelBold = True
            End If
        End If
        i = i + 1
    Loop
    r.SelStart = 0
    r.SelLength = 0
    r.Visible = True
Exit Sub
Error:
    MsgBox Err.Description, vbCritical
End Sub

Sub Sub_Optimize_Html()
    If m_Mdi.ActiveForm Is Nothing Then Exit Sub
    Dim r, x, i, z As String, a As Long
    i = 0
    Set r = m_Mdi.ActiveForm.Rtf_Text
    Dim Str_String
    Str_String = String(Len(r.Text), 0)
    Str_String = r.Text
    r.Visible = False
    Do Until i = Len(Str_String)
        DoEvents
        x = Mid(Str_String, i + 1, 1)
        If Asc(x) = 13 Or Asc(x) = 10 Then
            '-
        Else
            z = z + x
        End If
        i = i + 1
    Loop
    Dim Str_String2 As String, b As String
    Str_String2 = String(Len(z), 0)
    Str_String2 = z
    r.Visible = False
    a = 0
    Do Until a = Len(Str_String2)
        DoEvents
        x = Mid(Str_String2, a + 1, 1)
        If x = "<" Then
            b = b + Chr(13)
            b = b + Chr(10)
            b = b + x
        Else
            b = b + x
        End If
        a = a + 1
    Loop
    r.Text = b
    r.SelStart = 0
    r.SelLength = Len(r.Text)
    r.SelColor = vbBlack
    r.SelBold = False
    r.SelStart = 0
    r.SelLength = 0
    r.Visible = True
End Sub

Sub Sub_Proof_Html()
On Error GoTo Error
    If m_Mdi.ActiveForm Is Nothing Then Exit Sub
    Dim r, x, i, m_Open As Long, m_Close As Long
    i = 0
    Set r = m_Mdi.ActiveForm.Rtf_Text
    Dim Str_String
    Str_String = String(Len(r.Text), 0)
    Str_String = r.Text
    r.Visible = False
    Do Until i = Len(Str_String)
    DoEvents
        x = Mid(Str_String, i + 1, 1)
        If x = "<" Then
            m_Open = m_Open + 1
        ElseIf x = ">" Then
            m_Close = m_Close + 1
        End If
        i = i + 1
    Loop
    If m_Close = m_Open Then
        MsgBox "Der Code ist fehlerfrei!", vbInformation
    Else
        MsgBox "Code ist falsch! " & m_Open & " Klammern auf - " _
            & m_Close & " Klammern zu!", _
            vbCritical
    End If
    r.Visible = True
Exit Sub
Error:
    MsgBox Err.Description, vbCritical
End Sub

Sub Write_Last_Doc()
    '//
End Sub
