VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form TestForm 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Flat VBControls Extension"
   ClientHeight    =   3120
   ClientLeft      =   150
   ClientTop       =   390
   ClientWidth     =   4245
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   4245
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      Caption         =   "Enable Form"
      Height          =   255
      Left            =   1522
      TabIndex        =   11
      Top             =   2040
      Value           =   1  'Checked
      Width           =   1200
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3600
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   26
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TestForm.frx":0000
            Key             =   "AS"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TestForm.frx":045C
            Key             =   "AU"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TestForm.frx":08B8
            Key             =   "BE"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TestForm.frx":0D14
            Key             =   "BR"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TestForm.frx":1170
            Key             =   "CA"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TestForm.frx":15CC
            Key             =   "DE"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TestForm.frx":1A28
            Key             =   "FI"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TestForm.frx":1E84
            Key             =   "FR"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TestForm.frx":22E0
            Key             =   "GE"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TestForm.frx":273C
            Key             =   "IR"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TestForm.frx":2B98
            Key             =   "IT"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TestForm.frx":2FF4
            Key             =   "JA"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TestForm.frx":3450
            Key             =   "ME"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TestForm.frx":38AC
            Key             =   "NE"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TestForm.frx":3D08
            Key             =   "NO"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TestForm.frx":4164
            Key             =   "NZ"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TestForm.frx":45C0
            Key             =   "PO"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TestForm.frx":4A1C
            Key             =   "RS"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TestForm.frx":4E78
            Key             =   "RU"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TestForm.frx":52D4
            Key             =   "KO"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TestForm.frx":5730
            Key             =   "SP"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TestForm.frx":5B8C
            Key             =   "SW"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TestForm.frx":5FE8
            Key             =   "SZ"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TestForm.frx":6444
            Key             =   "TU"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TestForm.frx":68A0
            Key             =   "UK"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TestForm.frx":6CFC
            Key             =   "US"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageCombo ImageCombo1 
      Height          =   330
      Left            =   1185
      TabIndex        =   10
      Top             =   1005
      Width           =   2625
      _ExtentX        =   4630
      _ExtentY        =   582
      _Version        =   393216
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      ImageList       =   "ImageList1"
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   315
      Left            =   2115
      Max             =   120
      Min             =   5
      TabIndex        =   5
      Top             =   1500
      Value           =   5
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   1185
      TabIndex        =   4
      Top             =   540
      Width           =   2625
   End
   Begin VB.ComboBox cboAge 
      Height          =   315
      ItemData        =   "TestForm.frx":7158
      Left            =   1185
      List            =   "TestForm.frx":715A
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1500
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   1185
      TabIndex        =   0
      Top             =   60
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   615
      Left            =   2175
      TabIndex        =   2
      Top             =   2400
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Height          =   615
      Left            =   375
      TabIndex        =   1
      Top             =   2400
      Width           =   1695
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Country:"
      Height          =   195
      Left            =   465
      TabIndex        =   9
      Top             =   1080
      Width           =   585
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Age:"
      Height          =   195
      Left            =   600
      TabIndex        =   8
      Top             =   1560
      Width           =   330
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Address:"
      Height          =   195
      Left            =   435
      TabIndex        =   7
      Top             =   600
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Name:"
      Height          =   195
      Left            =   585
      TabIndex        =   6
      Top             =   120
      Width           =   465
   End
   Begin VB.Menu File 
      Caption         =   "&File"
      Begin VB.Menu mnuFile 
         Caption         =   "&Exit"
         Index           =   10
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu Help 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelp 
         Caption         =   "&On-Line"
         Index           =   10
         Begin VB.Menu mnuOnLine 
            Caption         =   "Official &WebPage"
            Index           =   10
         End
         Begin VB.Menu mnuOnLine 
            Caption         =   "E-Mail"
            Index           =   20
         End
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "&View ReadMe File"
         Index           =   20
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "-"
         Index           =   29
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "&About"
         Index           =   30
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "TestForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private K() As cControlFlater, I As Integer

Private Sub Form_Load()
    Dim CTL As Control, J As Integer
    
    With ImageCombo1.ComboItems
        .Add , , "Australia", "AS"
        .Add , , "Austria", "AU"
        .Add , , "Belgium", "BE"
        .Add , , "Brazil", "BR"
        .Add , , "Canada", "CA"
        .Add , , "Denmark", "DE"
        .Add , , "Finland", "FI"
        .Add , , "France", "FR"
        .Add , , "Germany", "GE"
        .Add , , "Ireland", "IR"
        .Add , , "Italy", "IT"
        .Add , , "Japan", "JA"
        .Add , , "Mexico", "ME"
        .Add , , "Netherlands", "NE"
        .Add , , "Norway", "NO"
        .Add , , "NewZeland", "NZ"
        .Add , , "Portugal", "PO"
        .Add , , "Republic of South Africa", "RS"
        .Add , , "Russia", "RU"
        .Add , , "Korea", "KO"
        .Add , , "Spain", "SP"
        .Add , , "Sweden", "SW"
        .Add , , "Switzerland", "SZ"
        .Add , , "Turkey", "TU"
        .Add , , "United Kingdom", "UK"
        .Add , , "United States of America", "US"
        .Add , , "Other"
    End With
    For J = 120 To 5 Step -1
        cboAge.AddItem J
    Next J
    
    For Each CTL In Me.Controls
        Select Case TypeName(CTL)
        Case "CommandButton", "TextBox", "ComboBox", "ImageCombo", "HScrollBar"
            ReDim Preserve K(I)
            Set K(I) = New cControlFlater
            K(I).Attach CTL
            I = I + 1
        End Select
    Next CTL
'    Set C = New cControlFlater
'    C.Attach Text1
End Sub

Private Sub Check1_Click()
    On Error Resume Next
    Dim CTL As Control
    For Each CTL In Me.Controls
        If CTL.Name <> "Check1" Then CTL.Enabled = CBool(Check1.Value)
    Next CTL
End Sub

Private Sub cboAge_Click()
    HScroll1.Value = Val(cboAge.Text)
End Sub

Private Sub HScroll1_Change()
    cboAge.Text = HScroll1.Value
End Sub

Private Sub HScroll1_Scroll()
    HScroll1_Change
End Sub

Private Sub mnuFile_Click(Index As Integer)
    Select Case Index
    Case 10
        Unload Me
    End Select
End Sub

Private Sub mnuHelp_Click(Index As Integer)
    Dim Buffer As String
    Select Case Index
    Case 20
        ShellExecute Me.hwnd, "open", "ReadMe.htm", "", "", 1
    Case 30
        Buffer = Buffer & "This code is not completely mine! You can find the original code at VBAccelerator (www.VBAccelerator.com)." & vbCrLf
        Buffer = Buffer & "I only added suport for Flat CommandButtons and Flat Horizontal ScrollBars (Vertical ScrollBars can be made with some changes to the code)." & vbCrLf
        Buffer = Buffer & "On a future release, I intend to add suport for Vertical Scrollbars, CheckBoxes, OptionButtons and other controls!" & vbCrLf & vbCrLf
        Buffer = Buffer & "There is still a problem with the scroller on the ScrollBars! If you know how to find it and paint over it, please do tell me!" & vbCrLf & vbCrLf
        Buffer = Buffer & "Programmed by Pedro Lamas" & vbCrLf
        Buffer = Buffer & "Copyright ©1997-2000 Underground Software" & vbCrLf
        MsgBox Buffer, vbApplicationModal + vbInformation, "About"
    End Select
End Sub

Private Sub mnuOnLine_Click(Index As Integer)
    Select Case Index
    Case 10
        ShellExecute Me.hwnd, "open", "http://vbhelp.cjb.net", "", "", 1
    Case 20
        ShellExecute Me.hwnd, "open", "mailto:support@vbhelp.cjb.net", "", "", 1
    End Select
End Sub
