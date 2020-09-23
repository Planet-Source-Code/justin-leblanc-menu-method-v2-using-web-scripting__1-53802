VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Begin VB.Form FrmMain 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Menu Method V2"
   ClientHeight    =   3360
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   7170
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
   ScaleHeight     =   3360
   ScaleWidth      =   7170
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdOk 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ok"
      Height          =   270
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Set Select Menu Theme"
      Top             =   3000
      Width           =   375
   End
   Begin VB.ComboBox CBBThemes 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox TxtPage 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   2040
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   480
      Width           =   4815
   End
   Begin VB.PictureBox PFrame 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2175
      Left            =   210
      ScaleHeight     =   2175
      ScaleWidth      =   1350
      TabIndex        =   0
      Top             =   240
      Width           =   1345
      Begin SHDocVwCtl.WebBrowser WBBMenu 
         Height          =   2655
         Left            =   -30
         TabIndex        =   1
         Top             =   -30
         Width           =   2055
         ExtentX         =   3625
         ExtentY         =   4683
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   ""
      End
   End
   Begin VB.Label lblMenuTheme 
      BackStyle       =   0  'Transparent
      Caption         =   " &Menu Theme"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label lblclickit 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Right-Click the Menu"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Image ImgFrame 
      Height          =   3150
      Left            =   1800
      Picture         =   "FrmMain.frx":0000
      Top             =   120
      Width           =   5280
   End
   Begin VB.Label lblWebMenu 
      BackStyle       =   0  'Transparent
      Caption         =   " &Web Menu"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   1335
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Visible         =   0   'False
      Begin VB.Menu mnuWebmenu 
         Caption         =   "Web Menu"
      End
      Begin VB.Menu mnuWebScript 
         Caption         =   "Web Script"
      End
      Begin VB.Menu mnuCustomize 
         Caption         =   "Customize"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
      Begin VB.Menu mnuspace1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMinimize 
         Caption         =   "Minimize"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "Close"
      End
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





'Web Menu method V2
'By: Justion LeBlanc
'v_b_h_e_l_p@yahoo.com
'vote if you like it

'This is built off the first example i made only I now am showing you how you can
'use an external document like in this case a .INI document for the menu layout
'
'Note: I included 2 menu themes for examples (themes must be in same location as the .exe to use because I set the path as to app.path)
'
Private Sub Form_Load()
'Load Menu Themes
Call LoadThemes(App.Path, FrmMain.CBBThemes)
'set default theme variable
Call LoadMenuTheme(App.Path & "\" & CBBThemes.Text & ".ini", CurrentMenuTheme)
'to be able to use the document.write function with the web browser you MUST set a blank page to the web browser
'to do this we will navigate "About:Blank"
Call WBBMenu.Navigate("About:Blank")
End Sub

Private Sub mnuAbout_Click()
Call WBBMenu.Navigate("http://127.0.0.1/ABOUT")
End Sub

Private Sub mnuClose_Click()
End
End Sub

Private Sub mnuCustomize_Click()
Call WBBMenu.Navigate("http://127.0.0.1/CUSTOMIZE")
End Sub

Private Sub mnuMinimize_Click()
Me.WindowState = 1
End Sub

Private Sub mnuWebmenu_Click()
Call WBBMenu.Navigate("http://127.0.0.1/WEBMENU")
End Sub

Private Sub mnuWebScript_Click()
Call WBBMenu.Navigate("http://127.0.0.1/WEBSCRIPT")
End Sub

Private Sub WBBMenu_BeforeNavigate2(ByVal pDisp As Object, URL As Variant, Flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)
'ok this sub is key in following out the menu operations
'lets set up a case senerio
Dim CUrl As String
Dim Pos As Integer
'first we HAVE to "cancel" the link we are navigating so it wont actually goto that page if its not our blank page
If LCase(URL) = "about:blank" Then
  Call WBBMenu.Navigate("http://127.0.0.1/WEBMENU")
  Exit Sub
ElseIf LCase(URL) = "mailto:v_b_h_e_l_p@yahoo.com?subject=hello%20sweet%20project!" Then
  Exit Sub
End If
  Cancel = True
'second pull out case "name"
  Pos = InStrRev(URL, "/") 'I used the instrrev() function which allows look for a string in a string starting from the end of a string (opposite of instr())
  CUrl = Mid(URL, Pos + 1, Len(URL) - Pos) 'mid() function that allows you to pull a specified string from a string by setting what character to start at and what character to end at
'last our cases
Select Case CUrl
  Case "WEBMENU"
    TxtPage.Text = "Web Menu -" & vbCrLf & vbCrLf & "A creative Menu Method."
  Case "WEBSCRIPT"
    TxtPage.Text = "Web Script -" & vbCrLf & vbCrLf & "Using some basic web scripting (HTML, CSS, etc..) you can make your own custom menu. Go to this nice website for some web scripting help - http://www.w3schools.com"
  Case "CUSTOMIZE"
    TxtPage.Text = "Customize -" & vbCrLf & vbCrLf & "By editing the web script you can customize your menu however you want it to look. You can also make it scrollable, just remove the picturebox control I used like a frame 'PFrame'"
  Case "ABOUT"
    TxtPage.Text = "About -" & vbCrLf & vbCrLf & "Simple Menu Method V2" & vbCrLf & "By: Justin LeBlanc" & vbCrLf & "Vote if you like this"
  Case "POPMENU"
    Call Me.PopupMenu(Me.mnuFile)
End Select
End Sub

Private Sub WBBMenu_NavigateComplete2(ByVal pDisp As Object, URL As Variant)
'to make sure the page is done navigating the blank page before we write to it
'we'll have it write to it once navigation is complete
Call RightClickScript("http://127.0.0.1/POPMENU", FrmMain.WBBMenu)
DoEvents:
'Call MenuScript(FrmMain.WBBMenu)
Call Write2Web(FrmMain.WBBMenu, CurrentMenuTheme)
End Sub

Private Sub CmdOk_Click()
'sets new menu theme
Call LoadMenuTheme(App.Path & "\" & CBBThemes.Text & ".ini", CurrentMenuTheme)
Call WBBMenu.Navigate("About:Blank")
End Sub
