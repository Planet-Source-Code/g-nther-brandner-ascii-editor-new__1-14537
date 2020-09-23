VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm Mdi_Frm_Main 
   AutoShowChildren=   0   'False
   BackColor       =   &H8000000C&
   Caption         =   "Ascii-Edit - The Editor for the Expert"
   ClientHeight    =   9030
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   11745
   Icon            =   "Mdi_Frm_Main.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'Bildschirmmitte
   Begin MSComDlg.CommonDialog Com_Dialog 
      Left            =   30
      Top             =   420
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Unten ausrichten
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   8715
      Width           =   11745
      _ExtentX        =   20717
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            Alignment       =   1
            AutoSize        =   1
            Bevel           =   2
            Enabled         =   0   'False
            Object.Width           =   4022
            TextSave        =   "EINFG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   4
            Alignment       =   1
            AutoSize        =   1
            Bevel           =   2
            Enabled         =   0   'False
            Object.Width           =   4022
            TextSave        =   "ROLL"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   1
            Bevel           =   2
            Object.Width           =   4022
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Bevel           =   2
            Object.Width           =   4022
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   4022
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   30
      Top             =   390
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Mdi_Frm_Main.frx":0442
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Mdi_Frm_Main.frx":0554
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Mdi_Frm_Main.frx":0666
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Mdi_Frm_Main.frx":0778
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Mdi_Frm_Main.frx":088A
            Key             =   "Cut"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Mdi_Frm_Main.frx":099C
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Mdi_Frm_Main.frx":0AAE
            Key             =   "Paste"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbToolBar 
      Align           =   1  'Oben ausrichten
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   11745
      _ExtentX        =   20717
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "imlToolbarIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   10
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Neu"
            Object.ToolTipText     =   "Neu"
            ImageKey        =   "New"
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Öffnen"
            Object.ToolTipText     =   "Öffnen"
            ImageKey        =   "Open"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Speichern"
            Object.ToolTipText     =   "Speichern"
            ImageKey        =   "Save"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Drucken"
            Object.ToolTipText     =   "Drucken"
            ImageKey        =   "Print"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Ausschneiden"
            Object.ToolTipText     =   "Ausschneiden"
            ImageKey        =   "Cut"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Kopieren"
            Object.ToolTipText     =   "Kopieren"
            ImageKey        =   "Copy"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Einfügen"
            Object.ToolTipText     =   "Einfügen"
            ImageKey        =   "Paste"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnu_File 
      Caption         =   "&Datei"
      Begin VB.Menu mnu_New 
         Caption         =   "&Neu"
      End
      Begin VB.Menu mnu_Open 
         Caption         =   "&Öffnen..."
      End
      Begin VB.Menu mnu_Sep0 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_Close 
         Caption         =   "&Schließen"
      End
      Begin VB.Menu mnu_Close_All 
         Caption         =   "&Alle schließen"
      End
      Begin VB.Menu mnu_Close_All_Without_Save 
         Caption         =   "A&lle schließen ohne speichern"
      End
      Begin VB.Menu mnu_Sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_Save 
         Caption         =   "S&peichern"
      End
      Begin VB.Menu mnu_Save_As 
         Caption         =   "Speichern &unter..."
      End
      Begin VB.Menu mnu_Sep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_Printer_Setup 
         Caption         =   "Drucker &einrichten..."
      End
      Begin VB.Menu mnu_Print 
         Caption         =   "&Drucken..."
      End
      Begin VB.Menu mnu_Sep9 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_Path_Edit 
         Caption         =   "&Pfadangaben..."
      End
      Begin VB.Menu mnu_Sep10 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_File_Information 
         Caption         =   "&Dateieigenschaften..."
      End
      Begin VB.Menu mnu_Sep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_Files_Opened_Last 
         Caption         =   "&Zuletzt geöffnete Dateien..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnu_Sep7 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_Exit 
         Caption         =   "&Beenden"
      End
   End
   Begin VB.Menu mnu_Edit 
      Caption         =   "&Bearbeiten"
      Begin VB.Menu mnu_Copy 
         Caption         =   "&Kopieren"
      End
      Begin VB.Menu mnu_Cut 
         Caption         =   "&Ausschneiden"
      End
      Begin VB.Menu mnu_Paste 
         Caption         =   "&Einfügen"
      End
      Begin VB.Menu mnu_Sep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_Search 
         Caption         =   "&Suchen..."
      End
   End
   Begin VB.Menu mnu_View 
      Caption         =   "&Ansicht"
      Begin VB.Menu mnu_Toolbar 
         Caption         =   "&Symbolleiste"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnu_Status_Bar 
         Caption         =   "S&tatusleiste"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnu_Html 
      Caption         =   "&Html"
      Begin VB.Menu mnu_Proof_Html 
         Caption         =   "&Html-Code auf Richtigkeit überprüfen"
      End
      Begin VB.Menu mnu_Color_Code 
         Caption         =   "&Farbiger Html-Code"
      End
      Begin VB.Menu mnu_Color_Attitudes 
         Caption         =   "Farb&einstellungen für farbigen Html-Code..."
      End
      Begin VB.Menu mnu_Optimize_Html 
         Caption         =   "Html-Code &optimieren"
      End
   End
   Begin VB.Menu mnu_System 
      Caption         =   "&System"
      Begin VB.Menu mnu_Explorer 
         Caption         =   "&Explorer starten"
      End
   End
   Begin VB.Menu mnu_Window 
      Caption         =   "&Fenster"
      Begin VB.Menu mnu_View_New 
         Caption         =   "&Neues Fenster"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnu_Sep5 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_Cascade 
         Caption         =   "&Überlappend"
      End
      Begin VB.Menu mnu_Horizontal 
         Caption         =   "&Horizontal anordnen"
      End
      Begin VB.Menu mnu_Vertical 
         Caption         =   "&Vertikal anordnen"
      End
      Begin VB.Menu mnu_Symbols 
         Caption         =   "&Symbole anordnen"
      End
   End
   Begin VB.Menu mnu_Help_Help 
      Caption         =   "&?"
      Begin VB.Menu mnu_Help 
         Caption         =   "&Hilfe..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnu_Sep6 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_Info 
         Caption         =   "&Info..."
      End
   End
End
Attribute VB_Name = "Mdi_Frm_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'//


Private Sub MDIForm_Load()
On Error GoTo Error
    Dim t_Time  As SYSTEMTIME
    Dim m_Time As String
    GetSystemTime t_Time
    m_Time = t_Time.wDay & ". " & MonthName(t_Time.wMonth) & " " & _
        t_Time.wYear
    StatusBar.Panels(4).Text = m_Time
    m_Color = vbRed
    m_Active_Dir = GetSetting("Ascii-Editor", "Attitudes", "Active_Dir", "C:\")
Exit Sub
Error:
End Sub

Private Sub mnu_Cascade_Click()
On Error GoTo Error
    m_Mdi.Arrange vbCascade
Exit Sub
Error:
    MsgBox Err.Description, vbCritical
End Sub

Private Sub mnu_Close_All_Click()
On Error GoTo Error
    For i = 1 To Forms.Count - 1
        m_W_Save = 0
        Unload Forms(1)
    Next
Exit Sub
Error:
    MsgBox Err.Description, vbCritical
End Sub

Private Sub mnu_Close_All_Without_Save_Click()
On Error GoTo Error
If ActiveForm Is Nothing Then Exit Sub
Reply = MsgBox("Wollen Sie wirklich alle Dateien schließen, ohne zu speichern?" _
    , vbYesNo + vbQuestion)
If Reply = vbNo Then Exit Sub
    For i = 1 To Forms.Count - 1
        m_W_Save = 1
        Unload Forms(1)
    Next
Exit Sub
Error:
    MsgBox Err.Description, vbCritical
End Sub

Private Sub mnu_Close_Click()
On Error GoTo Error
    If ActiveForm Is Nothing Then Exit Sub
    Unload ActiveForm
Exit Sub
Error:
    MsgBox Err.Description, vbCritical
End Sub

Private Sub mnu_Color_Attitudes_Click()
    Load Frm_Color_Attitudes
    Frm_Color_Attitudes.Show (1)
End Sub

Private Sub mnu_Color_Code_Click()
    Sub_Color_Html
End Sub

Private Sub mnu_Copy_Click()
On Error GoTo Error
    If ActiveForm Is Nothing Then Exit Sub
    If ActiveForm.Rtf_Text.SelLength = 0 Then Exit Sub
    Clipboard.Clear
    Clipboard.SetText (ActiveForm.Rtf_Text.SelText)
Exit Sub
Error:
    MsgBox Err.Description, vbCritical
End Sub

Private Sub mnu_Cut_Click()
On Error GoTo Error
    If ActiveForm Is Nothing Then Exit Sub
    If ActiveForm.Rtf_Text.SelLength = 0 Then Exit Sub
    Call mnu_Copy_Click
    ActiveForm.Rtf_Text.SelText = ""
Exit Sub
Error:
    MsgBox Err.Description, vbCritical
End Sub
Private Sub mnu_Edit_Click()
    If ActiveForm Is Nothing Then
        mnu_Copy.Enabled = False
        mnu_Cut.Enabled = False
        mnu_Paste.Enabled = False
        mnu_Search.Enabled = False
    Else
        mnu_Copy.Enabled = True
        mnu_Cut.Enabled = True
        mnu_Paste.Enabled = True
        mnu_Search.Enabled = True
    End If
End Sub

Private Sub mnu_Exit_Click()
On Error GoTo Error
    End
Exit Sub
Error:
    MsgBox Err.Description, vbCritical
End Sub

Sub Load_New_Doc()
On Error GoTo Error
    Set m_Doc = New Frm_Document
    Load (m_Doc)
Exit Sub
Error:
    MsgBox Err.Description, vbCritical
End Sub

Private Sub mnu_Explorer_Click()
On Error GoTo Error
    Dim str_Dir As String
    Dim str_Len As Long
    str_Dir = String(255, 0)
    str_Len = GetWindowsDirectory(str_Dir, Len(str_Dir))
    str_Dir = Left(str_Dir, str_Len)
    str_Dir = str_Dir & "\" & "explorer.exe"
    Call Shell(str_Dir, vbNormalFocus)
Exit Sub
Error:
    MsgBox Err.Description, vbCritical
End Sub

Private Sub mnu_File_Click()
    If ActiveForm Is Nothing Then
        mnu_Close.Enabled = False
        mnu_Close_All.Enabled = False
        mnu_Close_All_Without_Save.Enabled = False
        mnu_Save.Enabled = False
        mnu_Save_As.Enabled = False
        mnu_Print.Enabled = False
        mnu_File_Information.Enabled = False
    Else
        mnu_Close.Enabled = True
        mnu_Close_All.Enabled = True
        mnu_Close_All_Without_Save.Enabled = True
        mnu_Save.Enabled = True
        mnu_Save_As.Enabled = True
        mnu_Print.Enabled = True
        mnu_File_Information.Enabled = True
    End If
End Sub

Private Sub mnu_File_Information_Click()
    If ActiveForm Is Nothing Then Exit Sub
    Load Frm_File_Information
    Frm_File_Information.Show (1)
End Sub

Private Sub mnu_Files_Opened_Last_Click()
    Load Frm_Opened_Files
    Frm_Opened_Files.Show (1)
End Sub

Private Sub mnu_Horizontal_Click()
On Error GoTo Error
    m_Mdi.Arrange vbHorizontal
Exit Sub
Error:
    MsgBox Err.Description, vbCritical
End Sub

Private Sub mnu_Html_Click()
    If ActiveForm Is Nothing Then
        mnu_Optimize_Html.Enabled = False
        mnu_Color_Attitudes.Enabled = False
        mnu_Color_Code.Enabled = False
        mnu_Proof_Html.Enabled = False
        Exit Sub
    End If
    Dim i, x
    m_A_Index = 0
    m_Caption = m_Mdi.ActiveForm.Caption
    For i = 1 To Len(m_Caption)
        x = Mid(m_Caption, i, 1)
        If x = "." Then
            m_A_Index = 1
        End If
        If m_A_Index = 1 Then
            m_Extension = m_Extension & x
        End If
    Next
    If m_Extension = ".htm" Or m_Extension = ".html" Then
        mnu_Optimize_Html.Enabled = True
        mnu_Color_Attitudes.Enabled = True
        mnu_Color_Code.Enabled = True
        mnu_Proof_Html.Enabled = True
    Else
        mnu_Optimize_Html.Enabled = False
        mnu_Color_Attitudes.Enabled = False
        mnu_Color_Code.Enabled = False
        mnu_Proof_Html.Enabled = False
    End If
End Sub

Private Sub mnu_Info_Click()
On Error GoTo Error
    Load Frm_Info
    Frm_Info.Show (1)
Exit Sub
Error:
    MsgBox Err.Description, vbCritical
End Sub

Private Sub mnu_New_Click()
On Error GoTo Error
    Call Load_New_Doc
    m_Doc_Number = m_Doc_Number + 1
    m_Doc.Caption = "Dokument" & " " & m_Doc_Number
    m_Doc.Show
Exit Sub
Error:
    MsgBox Err.Description, vbCritical
End Sub

Private Sub mnu_Open_Click()
On Error GoTo Error
    m_Active_Dir = GetSetting("Ascii-Editor", "Attitudes", "Active_Dir", "C:\")
    Com_Dialog.InitDir = m_Active_Dir
    Com_Dialog.Filter = "Alle Dateitypen (*.*)|*.*|Textdatei (*.txt)|*.txt|Html-Dateien (*.htm)|*.htm|Inf-Dateien (*.inf)|*.inf|"
    Com_Dialog.ShowOpen
    m_File_Path = Com_Dialog.FileName
    For i = 1 To Forms.Count - 1
        m_Window_Caption = Forms(i).Caption
        If m_Window_Caption = m_File_Path Then
            MsgBox "Die Datei ist bereits geöffnet!", vbInformation
            Exit Sub
        End If
    Next
    If Len(m_File_Path) < 3 Then Exit Sub
    Call Load_New_Doc
    m_Doc.Show
    ActiveForm.Caption = m_File_Path
    ActiveForm.Rtf_Text.LoadFile m_File_Path, rtfText
Exit Sub
Error:
End Sub

Private Sub mnu_Optimize_Html_Click()
    Call Sub_Optimize_Html
End Sub

Private Sub mnu_Page_Preview_Click()
    If ActiveForm Is Nothing Then Exit Sub
    Load Frm_Page_Preview
    Frm_Page_Preview.Show (1)
End Sub

Private Sub mnu_Paste_Click()
On Error GoTo Error
    If Clipboard Is Nothing Then Exit Sub
    If ActiveForm Is Nothing Then Exit Sub
    Dim m_String As String, m_String_Left As String, _
    m_String_Right As String, m_String_Rtf As String
    m_String = Clipboard.GetText
    m_String_Length = Len(m_String)
    m_String_Rtf = ActiveForm.Rtf_Text.Text
    m_sel_start = ActiveForm.Rtf_Text.SelStart
    With ActiveForm
        m_Start = .Rtf_Text.SelStart
        m_String_Left = Mid(m_String_Rtf, 1, m_Start)
        m_String_Right = Mid(m_String_Rtf, m_Start + 1, Len(m_String_Rtf))
        .Rtf_Text.Text = m_String_Left & m_String & m_String_Right
        .Rtf_Text.SelLength = 0
        .Rtf_Text.SelStart = m_sel_start + m_String_Length
    End With
Exit Sub
Error:
    MsgBox Err.Description, vbCritical
End Sub

Private Sub mnu_Path_Edit_Click()
    Load Frm_Path_Edit
    Frm_Path_Edit.Show (1)
End Sub

Private Sub mnu_Print_Click()
On Error GoTo Error
    If ActiveForm Is Nothing Then Exit Sub
    Load Frm_Print
    Frm_Print.Show (1)
Exit Sub
Error:
End Sub

Private Sub mnu_Printer_Setup_Click()
    Com_Dialog.ShowPrinter
End Sub

Private Sub mnu_Proof_Html_Click()
    Call Sub_Proof_Html
End Sub

Private Sub mnu_Save_As_Click()
On Error GoTo Error
    If ActiveForm Is Nothing Then Exit Sub
    m_Active_Dir = GetSetting("Ascii-Editor", "Attitudes", "Active_Dir", "C:\")
    Com_Dialog.InitDir = m_Active_Dir
    Com_Dialog.Filter = "Alle Dateitypen (*.*)|*.*|Textdatei (*.txt)|*.txt|Html-Dateien (*.htm)|*.htm|Inf-Dateien (*.inf)|*.inf|"
    Com_Dialog.ShowSave
    m_Save_Path = Com_Dialog.FileName
    Sub_Path_Control
    If m_Char_Index >= 4 Then
        ActiveForm.Rtf_Text.SaveFile m_Save_Path, rtfText
        ActiveForm.Caption = m_Save_Path
        ActiveForm.Rtf_Text.FileName = m_Save_Path
    Else
        MsgBox "Der Dateiname ist ungültig!", vbInformation
    End If
    Dim m_Caption As String, m_A_Index As Long, m_Extension As String
    If m_Mdi.ActiveForm Is Nothing Then Exit Sub
Exit Sub
Error:
End Sub

Private Sub mnu_Save_Click()
On Error GoTo Error
    If ActiveForm Is Nothing Then Exit Sub
    If ActiveForm.Rtf_Text.FileName <> "" Then
        ActiveForm.Rtf_Text.SaveFile ActiveForm.Rtf_Text.FileName, rtfText
    Else
        Call mnu_Save_As_Click
    End If
Exit Sub
Error:
    MsgBox Err.Description, vbCritical
End Sub

Private Sub mnu_Search_Click()
On Error GoTo Error
    If ActiveForm Is Nothing Then Exit Sub
    Load Frm_Search
    Frm_Search.Show (0)
Exit Sub
Error:
    MsgBox Err.Description, vbCritical
End Sub

Private Sub mnu_Status_Bar_Click()
On Error GoTo Error
    If StatusBar.Visible = True Then
        StatusBar.Visible = False
        mnu_Status_Bar.Checked = False
    Else
        StatusBar.Visible = True
        mnu_Status_Bar.Checked = True
    End If
Exit Sub
Error:
    MsgBox Err.Description, vbCritical
End Sub

Private Sub mnu_Symbols_Click()
On Error GoTo Error
    m_Mdi.Arrange vbArrangeIcons
Exit Sub
Error:
    MsgBox Err.Description, vbCritical
End Sub

Private Sub mnu_Toolbar_Click()
On Error GoTo Error
    If tbToolBar.Visible = True Then
        tbToolBar.Visible = False
        mnu_Toolbar.Checked = False
    Else
        tbToolBar.Visible = True
        mnu_Toolbar.Checked = True
    End If
Exit Sub
Error:
    MsgBox Err.Description, vbCritical
End Sub

Private Sub mnu_Vertical_Click()
On Error GoTo Error
    m_Mdi.Arrange vbVertical
Exit Sub
Error:
    MsgBox Err.Description, vbCritical
End Sub

Private Sub mnu_Window_Click()
    If ActiveForm Is Nothing Then
        mnu_Cascade.Enabled = False
        mnu_Horizontal.Enabled = False
        mnu_Symbols.Enabled = False
        mnu_Vertical.Enabled = False
    Else
        mnu_Cascade.Enabled = True
        mnu_Horizontal.Enabled = True
        mnu_Symbols.Enabled = True
        mnu_Vertical.Enabled = True
    End If
End Sub

Private Sub StatusBar_PanelClick(ByVal Panel As MSComctlLib.Panel)
    Select Case Panel.Index
        Case 4
            Load Frm_Date_Time
            Frm_Date_Time.Show (1)
    End Select
End Sub

Private Sub tbToolBar_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo Error
    Select Case Button.Index
        Case 1
            Call mnu_New_Click
        Case 2
            Call mnu_Open_Click
        Case 3
            Call mnu_Save_Click
        Case 5
            Call mnu_Print_Click
        Case 7
            Call mnu_Cut_Click
        Case 8
            Call mnu_Copy_Click
        Case 9
            Call mnu_Paste_Click
    End Select
Exit Sub
Error:
    MsgBox Err.Description, vbCritical
End Sub
