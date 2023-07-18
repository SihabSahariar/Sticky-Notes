VERSION 5.00
Begin VB.Form frm_Sticky 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   3810
   ClientLeft      =   6555
   ClientTop       =   4110
   ClientWidth     =   2655
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frm_Sticky.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   3810
   ScaleWidth      =   2655
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtLeft 
      Height          =   495
      Left            =   4920
      TabIndex        =   4
      Top             =   960
      Width           =   735
   End
   Begin VB.TextBox txtTop 
      Height          =   495
      Left            =   4920
      TabIndex        =   3
      Top             =   240
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   3255
      Left            =   120
      MaxLength       =   200
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   360
      Width           =   2415
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "x"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.5
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   2400
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Menu mnM 
      Caption         =   "mEnu"
      Visible         =   0   'False
      Begin VB.Menu abt 
         Caption         =   "About"
      End
      Begin VB.Menu mnCut 
         Caption         =   "Cut"
      End
      Begin VB.Menu mnCopy 
         Caption         =   "Copy"
      End
      Begin VB.Menu mnPast 
         Caption         =   "Paste"
      End
      Begin VB.Menu mnDelete 
         Caption         =   "Delete"
      End
      Begin VB.Menu mnSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnSlct 
         Caption         =   "Select All"
      End
      Begin VB.Menu mnSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnb 
         Caption         =   "Blue"
      End
      Begin VB.Menu mng 
         Caption         =   "Green"
      End
      Begin VB.Menu mnPk 
         Caption         =   "Pink"
      End
      Begin VB.Menu mnPb 
         Caption         =   "Purble"
      End
      Begin VB.Menu mnw 
         Caption         =   "White"
      End
      Begin VB.Menu mny 
         Caption         =   "Yellow"
      End
   End
End
Attribute VB_Name = "frm_Sticky"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2

Private Sub abt_Click()
MsgBox "Programmer: Sihab Sahariar Sizan", vbInformation, "About"

End Sub

Private Sub Form_Load()
On Error Resume Next
    AddToStartUp App.EXEName, App.Path & "\" & App.EXEName & ".exe", True

frm_Sticky.BackColor = GetSetting(App.Title, "Settings", "backCol")
Text1.BackColor = GetSetting(App.Title, "Settings", "txtbackCol")
Text1.ForeColor = GetSetting(App.Title, "Settings", "txtforCol")
Label1.ForeColor = GetSetting(App.Title, "Settings", "lbl1forCol")
Label2.ForeColor = GetSetting(App.Title, "Settings", "lbl2forCol")
Text1.Text = GetSetting(App.Title, "settings", "txtnote")
txtTop.Text = GetSetting(App.Title, "settings", "frmTop")
txtLeft.Text = GetSetting(App.Title, "settings", "frmleft")

Me.Top = txtTop.Text
Me.Left = txtLeft.Text
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then PopupMenu mnM
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.Visible = True
Label2.Visible = True
    If Button = vbLeftButton Then
        ReleaseCapture
        SendMessageLong Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
    End If
txtTop.Text = Me.Top
txtLeft.Text = Me.Left
SaveSetting App.Title, "Settings", "frmTop", Trim(txtTop.Text)
SaveSetting App.Title, "Settings", "frmleft", Trim(txtLeft.Text)

End Sub
Private Sub Label1_Click()
   Dim Form As New frm_Sticky
    Load Form
    Form.Visible = True
   Form.Left = frm_Sticky.Left + frm_Sticky.Width + 20
   Form.Top = frm_Sticky.Top
Form.BackColor = GetSetting(App.Title, "Settings", "backCol")
Form.Text1.BackColor = GetSetting(App.Title, "Settings", "txtbackCol")
Form.Text1.ForeColor = GetSetting(App.Title, "Settings", "txtforCol")
Form.Label1.ForeColor = GetSetting(App.Title, "Settings", "lbl1forCol")
Form.Label2.ForeColor = GetSetting(App.Title, "Settings", "lbl2forCol")
SaveSetting App.Title, "Settings", "txtNote1", Trim(Form.Text1.Text)
End Sub
Private Sub Label2_Click()
Unload Me
End Sub
Private Sub mnb_Click()
blue
End Sub
Private Sub mng_Click()
green
End Sub
Private Sub mnPb_Click()
Purble
End Sub
Private Sub mnPk_Click()
pink
End Sub
Private Sub mnw_Click()
white
End Sub
Private Sub mny_Click()
yellow
End Sub
Private Sub Text1_Change()
SaveSetting App.Title, "Settings", "txtNote", Trim(Text1.Text)
End Sub
Private Sub Text1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.Visible = False
Label2.Visible = False
End Sub
Sub blue()
frm_Sticky.BackColor = &H504000
Text1.BackColor = &H80000003
Text1.ForeColor = &H0&
Label1.ForeColor = &HFFFFFF
Label2.ForeColor = &HFFFFFF
SaveSetting App.Title, "Settings", "backCol", Trim(frm_Sticky.BackColor)
SaveSetting App.Title, "Settings", "txtbackCol", Trim(Text1.BackColor)
SaveSetting App.Title, "Settings", "txtforCol", Trim(Text1.ForeColor)
SaveSetting App.Title, "Settings", "lbl1forCol", Trim(Label1.ForeColor)
SaveSetting App.Title, "Settings", "lbl2forCol", Trim(Label2.ForeColor)
End Sub
Sub green()
frm_Sticky.BackColor = &H8000&
Text1.BackColor = &HC0FFC0
Text1.ForeColor = &H0&
Label1.ForeColor = &HFFFFFF
Label2.ForeColor = &HFFFFFF
SaveSetting App.Title, "Settings", "backCol", Trim(frm_Sticky.BackColor)
SaveSetting App.Title, "Settings", "txtbackCol", Trim(Text1.BackColor)
SaveSetting App.Title, "Settings", "txtforCol", Trim(Text1.ForeColor)
SaveSetting App.Title, "Settings", "lbl1forCol", Trim(Label1.ForeColor)
SaveSetting App.Title, "Settings", "lbl2forCol", Trim(Label2.ForeColor)
End Sub
Sub pink()
frm_Sticky.BackColor = &HFF00FF
Text1.BackColor = &HFFC0FF
Text1.ForeColor = &H0&
Label1.ForeColor = &HFFFFFF
Label2.ForeColor = &HFFFFFF
SaveSetting App.Title, "Settings", "backCol", Trim(frm_Sticky.BackColor)
SaveSetting App.Title, "Settings", "txtbackCol", Trim(Text1.BackColor)
SaveSetting App.Title, "Settings", "txtforCol", Trim(Text1.ForeColor)
SaveSetting App.Title, "Settings", "lbl1forCol", Trim(Label1.ForeColor)
SaveSetting App.Title, "Settings", "lbl2forCol", Trim(Label2.ForeColor)
End Sub
Sub Purble()
frm_Sticky.BackColor = &HFF6080
Text1.BackColor = &HFFC0C0
Text1.ForeColor = &H0&
Label1.ForeColor = &HFFFFFF
Label2.ForeColor = &HFFFFFF
SaveSetting App.Title, "Settings", "backCol", Trim(frm_Sticky.BackColor)
SaveSetting App.Title, "Settings", "txtbackCol", Trim(Text1.BackColor)
SaveSetting App.Title, "Settings", "txtforCol", Trim(Text1.ForeColor)
SaveSetting App.Title, "Settings", "lbl1forCol", Trim(Label1.ForeColor)
SaveSetting App.Title, "Settings", "lbl2forCol", Trim(Label2.ForeColor)
End Sub
Sub white()
frm_Sticky.BackColor = &HFFFFFF
Text1.BackColor = &HFFFFFF
Text1.ForeColor = &H0&
Label1.ForeColor = &H0&
Label2.ForeColor = &H0&
SaveSetting App.Title, "Settings", "backCol", Trim(frm_Sticky.BackColor)
SaveSetting App.Title, "Settings", "txtbackCol", Trim(Text1.BackColor)
SaveSetting App.Title, "Settings", "txtforCol", Trim(Text1.ForeColor)
SaveSetting App.Title, "Settings", "lbl1forCol", Trim(Label1.ForeColor)
SaveSetting App.Title, "Settings", "lbl2forCol", Trim(Label2.ForeColor)
End Sub
Sub yellow()
frm_Sticky.BackColor = &HFFFF&
Text1.BackColor = &H80FFFF
Text1.ForeColor = &H0&
Label1.ForeColor = &H0&
Label2.ForeColor = &H0&
SaveSetting App.Title, "Settings", "backCol", Trim(frm_Sticky.BackColor)
SaveSetting App.Title, "Settings", "txtbackCol", Trim(Text1.BackColor)
SaveSetting App.Title, "Settings", "txtforCol", Trim(Text1.ForeColor)
SaveSetting App.Title, "Settings", "lbl1forCol", Trim(Label1.ForeColor)
SaveSetting App.Title, "Settings", "lbl2forCol", Trim(Label2.ForeColor)
End Sub
Private Sub DoEditThing(whatThing As String, onWhat As Object)
Dim Send$
   Select Case whatThing
        Case "Copy"
            Send = "^C"
        Case "Cut"
            Send = "^X"
        Case "Paste"
            Send = "^V"
        Case "Undo"
            Send = "^Z"

    End Select
    If Len(Send) Then
        onWhat.SetFocus
        SendKeys Send
    End If
End Sub
Private Sub mnCopy_Click()
DoEditThing "Copy", Text1
End Sub
Private Sub mnCut_Click()
DoEditThing "Cut", Text1
End Sub
Private Sub mnPast_Click()
DoEditThing "Paste", Text1
End Sub
Private Sub mnDelete_Click()
Text1.Text = ""
End Sub
Private Sub mnSlct_Click()
           ' Text1.SelText.Select

End Sub

