VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tree Builder Version 3"
   ClientHeight    =   8055
   ClientLeft      =   2565
   ClientTop       =   1500
   ClientWidth     =   9990
   Icon            =   "frmOptions1.frx":0000
   KeyPreview      =   -1  'True
   ScaleHeight     =   8055
   ScaleWidth      =   9990
   Begin VB.CommandButton ReadConfig 
      Caption         =   "Read Config"
      Height          =   375
      Left            =   8640
      TabIndex        =   182
      ToolTipText     =   "Use this to read all values from a saved configuration file."
      Top             =   7440
      Width           =   1095
   End
   Begin VB.CommandButton ResetAll 
      Caption         =   "Reset All"
      Height          =   375
      Left            =   6000
      TabIndex        =   9
      ToolTipText     =   "Use this to reset all default values."
      Top             =   7440
      Width           =   1095
   End
   Begin VB.CommandButton SaveConfig 
      Caption         =   "Save Config"
      Height          =   375
      Left            =   7320
      TabIndex        =   12
      ToolTipText     =   "Use this to save values that you have modified."
      Top             =   7440
      Width           =   1095
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   240
      TabIndex        =   6
      ToolTipText     =   "Applies only when files and directories are being created."
      Top             =   7080
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Min             =   1
   End
   Begin VB.CommandButton CmdWrite 
      Caption         =   "Write Files"
      Height          =   375
      Left            =   3360
      TabIndex        =   10
      ToolTipText     =   $"frmOptions1.frx":0CCA
      Top             =   7440
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog cmnDialog 
      Left            =   9000
      Top             =   7440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   3
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample4 
         Caption         =   "Sample 4"
         Height          =   1785
         Left            =   2100
         TabIndex        =   5
         Top             =   840
         Width           =   2055
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   2
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample3 
         Caption         =   "Sample 3"
         Height          =   1785
         Left            =   1545
         TabIndex        =   4
         Top             =   675
         Width           =   2055
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   1
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample2 
         Caption         =   "Sample 2"
         Height          =   1785
         Left            =   645
         TabIndex        =   3
         Top             =   300
         Width           =   2055
      End
   End
   Begin VB.CommandButton HelpRead 
      Caption         =   "Help"
      Height          =   375
      Left            =   4680
      TabIndex        =   11
      ToolTipText     =   "Click here for On-Line help and readme information."
      Top             =   7440
      Width           =   1095
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      TabIndex        =   8
      ToolTipText     =   "Click here to close Tree Builder"
      Top             =   7440
      Width           =   1335
   End
   Begin VB.CommandButton CmdUpdate 
      Caption         =   "Run Update"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   7
      ToolTipText     =   "Once you have supplied the above BOLDED items, hit this button to create users."
      Top             =   7440
      Width           =   1335
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8055
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   14208
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      OLEDropMode     =   1
      TabCaption(0)   =   "ICE Settings"
      TabPicture(0)   =   "frmOptions1.frx":0D89
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label33"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label13"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label12"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label14"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label9"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label11"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label30"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label28"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label48"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label49"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label50"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label51"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label53"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label54"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "CmdViewICE"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "QS"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Text31"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "CustomCheck"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Text30"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "CmdBrowse10"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "cmdViewCust"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "cmdClean"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "cmdViewRice"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "CmdBrowse5"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Text19"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Combo1"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "CheckDel"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "cmdViewImp"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "cmdViewExp"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "AnonCheck"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "Text14"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "UserHomeCheck"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "CmdBrowse4"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "Text7"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "CheckAdd"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "RetrieveTreeCheck"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "CmdBrowse3"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "Text6"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "Combo6"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "Combo2"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "Text1"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "Text2"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "Text3"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "Text4"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "Text5"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "CmdBrowse1"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).Control(46)=   "CmdBrowse2"
      Tab(0).Control(46).Enabled=   0   'False
      Tab(0).Control(47)=   "StopOnErrCheck"
      Tab(0).Control(47).Enabled=   0   'False
      Tab(0).Control(48)=   "InstallCleanup"
      Tab(0).Control(48).Enabled=   0   'False
      Tab(0).ControlCount=   49
      TabCaption(1)   =   "User Configuration"
      TabPicture(1)   =   "frmOptions1.frx":0DA5
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label16"
      Tab(1).Control(1)=   "Label2"
      Tab(1).Control(2)=   "Label1"
      Tab(1).Control(3)=   "Label3"
      Tab(1).Control(4)=   "Label101"
      Tab(1).Control(5)=   "Label100"
      Tab(1).Control(6)=   "Label5"
      Tab(1).Control(7)=   "Label6"
      Tab(1).Control(8)=   "Label7"
      Tab(1).Control(9)=   "Label4"
      Tab(1).Control(10)=   "Label8"
      Tab(1).Control(11)=   "Label15"
      Tab(1).Control(12)=   "Label20"
      Tab(1).Control(13)=   "Label25"
      Tab(1).Control(14)=   "Label27"
      Tab(1).Control(15)=   "Label29"
      Tab(1).Control(16)=   "Label31"
      Tab(1).Control(17)=   "Label47"
      Tab(1).Control(18)=   "Label43"
      Tab(1).Control(19)=   "Label44"
      Tab(1).Control(20)=   "CustContCheck"
      Tab(1).Control(21)=   "Text29"
      Tab(1).Control(22)=   "SelectAllUserSetCheck"
      Tab(1).Control(23)=   "Text101"
      Tab(1).Control(24)=   "Text105"
      Tab(1).Control(25)=   "Text109"
      Tab(1).Control(26)=   "Text113"
      Tab(1).Control(27)=   "UserSetFourCheck"
      Tab(1).Control(28)=   "UserSetThreeCheck"
      Tab(1).Control(29)=   "UserSetTwoCheck"
      Tab(1).Control(30)=   "UserSetOneCheck"
      Tab(1).Control(31)=   "Text10"
      Tab(1).Control(32)=   "Text11"
      Tab(1).Control(33)=   "Text12"
      Tab(1).Control(34)=   "Text13"
      Tab(1).Control(35)=   "Text17"
      Tab(1).Control(36)=   "Text16"
      Tab(1).Control(37)=   "Text15"
      Tab(1).Control(38)=   "Text18"
      Tab(1).Control(39)=   "Text22"
      Tab(1).Control(40)=   "Text21"
      Tab(1).Control(41)=   "Text20"
      Tab(1).Control(42)=   "Text23"
      Tab(1).Control(43)=   "Text27"
      Tab(1).Control(44)=   "Text26"
      Tab(1).Control(45)=   "Text25"
      Tab(1).Control(46)=   "Text28"
      Tab(1).Control(47)=   "Combo14"
      Tab(1).Control(48)=   "Text8"
      Tab(1).Control(49)=   "Text9"
      Tab(1).Control(50)=   "Text112"
      Tab(1).Control(51)=   "Text108"
      Tab(1).Control(52)=   "Text104"
      Tab(1).Control(53)=   "Text100"
      Tab(1).Control(54)=   "Text114"
      Tab(1).Control(55)=   "Text110"
      Tab(1).Control(56)=   "Text106"
      Tab(1).Control(57)=   "Text102"
      Tab(1).Control(58)=   "Text115"
      Tab(1).Control(59)=   "Text111"
      Tab(1).Control(60)=   "Text107"
      Tab(1).Control(61)=   "Text103"
      Tab(1).Control(62)=   "SkipTreeCreateCheck"
      Tab(1).ControlCount=   63
      TabCaption(2)   =   "Tree Configuration"
      TabPicture(2)   =   "frmOptions1.frx":0DC1
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label23"
      Tab(2).Control(1)=   "Label17"
      Tab(2).Control(2)=   "Label66"
      Tab(2).Control(3)=   "Label65"
      Tab(2).Control(4)=   "Label64"
      Tab(2).Control(5)=   "Text24"
      Tab(2).Control(6)=   "Text1603"
      Tab(2).Control(7)=   "Text1503"
      Tab(2).Control(8)=   "Text1403"
      Tab(2).Control(9)=   "Text1303"
      Tab(2).Control(10)=   "Text1203"
      Tab(2).Control(11)=   "Text1103"
      Tab(2).Control(12)=   "Text1602"
      Tab(2).Control(13)=   "Text1502"
      Tab(2).Control(14)=   "Text1402"
      Tab(2).Control(15)=   "Text1302"
      Tab(2).Control(16)=   "Text1202"
      Tab(2).Control(17)=   "Text1102"
      Tab(2).Control(18)=   "Text1601"
      Tab(2).Control(19)=   "Text1501"
      Tab(2).Control(20)=   "Text1401"
      Tab(2).Control(21)=   "Text1301"
      Tab(2).Control(22)=   "Text1201"
      Tab(2).Control(23)=   "Text1101"
      Tab(2).Control(24)=   "Text200"
      Tab(2).Control(25)=   "Combo606"
      Tab(2).Control(26)=   "Combo605"
      Tab(2).Control(27)=   "Combo604"
      Tab(2).Control(28)=   "Combo506"
      Tab(2).Control(29)=   "Combo505"
      Tab(2).Control(30)=   "Combo504"
      Tab(2).Control(31)=   "Combo406"
      Tab(2).Control(32)=   "Combo405"
      Tab(2).Control(33)=   "Combo404"
      Tab(2).Control(34)=   "Combo306"
      Tab(2).Control(35)=   "Combo305"
      Tab(2).Control(36)=   "Combo304"
      Tab(2).Control(37)=   "Combo206"
      Tab(2).Control(38)=   "Combo205"
      Tab(2).Control(39)=   "Combo204"
      Tab(2).Control(40)=   "Combo106"
      Tab(2).Control(41)=   "Combo105"
      Tab(2).Control(42)=   "Combo104"
      Tab(2).Control(43)=   "Combo199"
      Tab(2).ControlCount=   44
      TabCaption(3)   =   "Home Dir Configuration"
      TabPicture(3)   =   "frmOptions1.frx":0DDD
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "CmdBrowse6"
      Tab(3).Control(1)=   "Text403"
      Tab(3).Control(2)=   "Text401"
      Tab(3).Control(3)=   "Text404"
      Tab(3).Control(4)=   "Text405"
      Tab(3).Control(5)=   "UHDReset"
      Tab(3).Control(6)=   "Label45"
      Tab(3).Control(7)=   "Label41"
      Tab(3).Control(8)=   "Label40"
      Tab(3).Control(9)=   "Label19"
      Tab(3).Control(10)=   "Label21"
      Tab(3).Control(11)=   "Label22"
      Tab(3).ControlCount=   12
      Begin VB.CheckBox SkipTreeCreateCheck 
         Caption         =   "Create Tree/Container Information"
         Height          =   255
         Left            =   -74760
         TabIndex        =   181
         Top             =   3840
         Value           =   1  'Checked
         Width           =   2775
      End
      Begin VB.CommandButton InstallCleanup 
         Caption         =   "Cleanup Installation Files?"
         Height          =   375
         Left            =   7200
         TabIndex        =   180
         ToolTipText     =   "Cleanup files created by Tree Builder."
         Top             =   480
         Width           =   2295
      End
      Begin VB.CheckBox StopOnErrCheck 
         Caption         =   "&Stop on Error(s)"
         Height          =   255
         Left            =   2040
         TabIndex        =   134
         ToolTipText     =   "If this is checked the Update will stop on errors, otherwise it will continue to attempt all changes in the LDIF file."
         Top             =   3000
         Width           =   1935
      End
      Begin VB.CommandButton CmdBrowse2 
         Caption         =   "Browse"
         Height          =   375
         Left            =   6600
         TabIndex        =   133
         ToolTipText     =   "Browse to a location of your choice."
         Top             =   3600
         Width           =   1335
      End
      Begin VB.CommandButton CmdBrowse1 
         Caption         =   "Browse"
         Height          =   375
         Left            =   4440
         TabIndex        =   132
         ToolTipText     =   $"frmOptions1.frx":0DF9
         Top             =   2040
         Width           =   1335
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   3240
         TabIndex        =   131
         Text            =   "C:\temp\rice.bat"
         ToolTipText     =   "This is the file name used to execute the ICE.EXE."
         Top             =   3600
         Width           =   3255
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   240
         TabIndex        =   130
         Text            =   "F:\Public\Rootcert.der"
         ToolTipText     =   "Typically located at ""F:\Public\Rootcert.der"""
         Top             =   2040
         Width           =   3975
      End
      Begin VB.TextBox Text3 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   2640
         PasswordChar    =   "*"
         TabIndex        =   129
         Text            =   "test"
         ToolTipText     =   "Default Value=Test"
         Top             =   1440
         Width           =   975
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   240
         TabIndex        =   128
         Text            =   "cn=admin,o=novell"
         ToolTipText     =   "Must be a user with rights to add at the Tree level."
         Top             =   1440
         Width           =   2295
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   240
         TabIndex        =   127
         Text            =   "255.255.255.255"
         ToolTipText     =   "IP Address of the LDAP server."
         Top             =   840
         Width           =   2655
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "frmOptions1.frx":0EA3
         Left            =   8520
         List            =   "frmOptions1.frx":0EBC
         Style           =   1  'Simple Combo
         TabIndex        =   126
         Text            =   "1"
         ToolTipText     =   "Value should be numeric and start at 1. Upper limit is only relevant to time and disk space."
         Top             =   1800
         Width           =   975
      End
      Begin VB.TextBox Text103 
         Height          =   285
         Left            =   -67080
         TabIndex        =   125
         Text            =   "Novell"
         Top             =   4680
         Width           =   1575
      End
      Begin VB.TextBox Text107 
         Height          =   285
         Left            =   -67080
         TabIndex        =   124
         Text            =   "Novell"
         Top             =   5040
         Width           =   1575
      End
      Begin VB.TextBox Text111 
         Height          =   285
         Left            =   -67080
         TabIndex        =   123
         Text            =   "Novell"
         Top             =   5400
         Width           =   1575
      End
      Begin VB.TextBox Text115 
         Height          =   285
         Left            =   -67080
         TabIndex        =   122
         Text            =   "Novell"
         Top             =   5760
         Width           =   1575
      End
      Begin VB.TextBox Text102 
         Height          =   285
         Left            =   -69000
         TabIndex        =   121
         Text            =   "Provo"
         Top             =   4680
         Width           =   1575
      End
      Begin VB.TextBox Text106 
         Height          =   285
         Left            =   -69000
         TabIndex        =   120
         Text            =   "Boston"
         Top             =   5040
         Width           =   1575
      End
      Begin VB.TextBox Text110 
         Height          =   285
         Left            =   -69000
         TabIndex        =   119
         Text            =   "Bangalore"
         Top             =   5400
         Width           =   1575
      End
      Begin VB.TextBox Text114 
         Height          =   285
         Left            =   -69000
         TabIndex        =   118
         Text            =   "Dublin"
         Top             =   5760
         Width           =   1575
      End
      Begin VB.TextBox Text100 
         Height          =   285
         Left            =   -72840
         TabIndex        =   117
         Text            =   "Users"
         Top             =   4680
         Width           =   1575
      End
      Begin VB.TextBox Text104 
         Height          =   285
         Left            =   -72840
         TabIndex        =   116
         Text            =   "Users"
         Top             =   5040
         Width           =   1575
      End
      Begin VB.TextBox Text108 
         Height          =   285
         Left            =   -72840
         TabIndex        =   115
         Text            =   "Users"
         Top             =   5400
         Width           =   1575
      End
      Begin VB.TextBox Text112 
         Height          =   285
         Left            =   -72840
         TabIndex        =   114
         Text            =   "Users"
         Top             =   5760
         Width           =   1575
      End
      Begin VB.TextBox Text9 
         Height          =   285
         Left            =   -72840
         TabIndex        =   113
         Text            =   "novell.com"
         Top             =   6600
         Width           =   1935
      End
      Begin VB.TextBox Text8 
         Height          =   285
         Left            =   -72840
         TabIndex        =   112
         Text            =   "mh.novell.com"
         Top             =   6240
         Width           =   1935
      End
      Begin VB.ComboBox Combo14 
         Height          =   315
         ItemData        =   "frmOptions1.frx":0EF1
         Left            =   -71520
         List            =   "frmOptions1.frx":0EFB
         TabIndex        =   111
         Text            =   "No"
         ToolTipText     =   "Select ""Yes"" to override the above password and make the password the same as the username."
         Top             =   2880
         Width           =   855
      End
      Begin VB.TextBox Text28 
         Height          =   285
         Left            =   -67200
         TabIndex        =   110
         Text            =   "test"
         ToolTipText     =   "Select a password to be used for ALL users, otherwise select ""Make All Passwords Match User Name""."
         Top             =   2400
         Width           =   1935
      End
      Begin VB.TextBox Text25 
         Height          =   285
         Left            =   -67200
         TabIndex        =   109
         Text            =   "DublinTestUser"
         ToolTipText     =   "UserName"
         Top             =   1320
         Width           =   1935
      End
      Begin VB.TextBox Text26 
         Height          =   285
         Left            =   -67200
         TabIndex        =   108
         Text            =   "Patrick"
         ToolTipText     =   "First Name"
         Top             =   1680
         Width           =   1935
      End
      Begin VB.TextBox Text27 
         Height          =   285
         Left            =   -67200
         TabIndex        =   107
         Text            =   "O'Hare"
         ToolTipText     =   "Last Name"
         Top             =   2040
         Width           =   1935
      End
      Begin VB.TextBox Text23 
         Height          =   285
         Left            =   -69360
         TabIndex        =   106
         Text            =   "test"
         ToolTipText     =   "Select a password to be used for ALL users, otherwise select ""Make All Passwords Match User Name""."
         Top             =   2400
         Width           =   1935
      End
      Begin VB.TextBox Text20 
         Height          =   285
         Left            =   -69360
         TabIndex        =   105
         Text            =   "BangaloreTestUser"
         ToolTipText     =   "UserName"
         Top             =   1320
         Width           =   1935
      End
      Begin VB.TextBox Text21 
         Height          =   285
         Left            =   -69360
         TabIndex        =   104
         Text            =   "Sudarshan"
         ToolTipText     =   "First Name"
         Top             =   1680
         Width           =   1935
      End
      Begin VB.TextBox Text22 
         Height          =   285
         Left            =   -69360
         TabIndex        =   103
         Text            =   "Sarkar"
         ToolTipText     =   "Last Name"
         Top             =   2040
         Width           =   1935
      End
      Begin VB.TextBox Text18 
         Height          =   285
         Left            =   -71520
         TabIndex        =   102
         Text            =   "test"
         ToolTipText     =   "Select a password to be used for ALL users, otherwise select ""Make All Passwords Match User Name""."
         Top             =   2400
         Width           =   1935
      End
      Begin VB.TextBox Text15 
         Height          =   285
         Left            =   -71520
         TabIndex        =   101
         Text            =   "BostonTestUser"
         ToolTipText     =   "UserName"
         Top             =   1320
         Width           =   1935
      End
      Begin VB.TextBox Text16 
         Height          =   285
         Left            =   -71520
         TabIndex        =   100
         Text            =   "Jack"
         ToolTipText     =   "First Name"
         Top             =   1680
         Width           =   1935
      End
      Begin VB.TextBox Text17 
         Height          =   285
         Left            =   -71520
         TabIndex        =   99
         Text            =   "Malone"
         ToolTipText     =   "Last Name"
         Top             =   2040
         Width           =   1935
      End
      Begin VB.TextBox Text13 
         Height          =   285
         Left            =   -73680
         TabIndex        =   98
         Text            =   "test"
         ToolTipText     =   "Select a password to be used for ALL users, otherwise select ""Make All Passwords Match User Name""."
         Top             =   2400
         Width           =   1935
      End
      Begin VB.TextBox Text12 
         Height          =   285
         Left            =   -73680
         TabIndex        =   97
         Text            =   "Doe"
         ToolTipText     =   "Last Name"
         Top             =   2040
         Width           =   1935
      End
      Begin VB.TextBox Text11 
         Height          =   285
         Left            =   -73680
         TabIndex        =   96
         Text            =   "John"
         ToolTipText     =   "First Name"
         Top             =   1680
         Width           =   1935
      End
      Begin VB.TextBox Text10 
         Height          =   285
         Left            =   -73680
         TabIndex        =   95
         Text            =   "ProvoTestUser"
         ToolTipText     =   "UserName"
         Top             =   1320
         Width           =   1935
      End
      Begin VB.CommandButton CmdBrowse6 
         Caption         =   "Browse"
         Height          =   375
         Left            =   -67920
         TabIndex        =   94
         ToolTipText     =   "Browse to the Drive and Directory you want to create user home directories in. The sub-directory."
         Top             =   1560
         Width           =   1335
      End
      Begin VB.TextBox Text403 
         Height          =   285
         Left            =   -70920
         TabIndex        =   93
         Text            =   "C:\Users\"
         ToolTipText     =   $"frmOptions1.frx":0F08
         Top             =   1560
         Width           =   2895
      End
      Begin VB.TextBox Text401 
         Height          =   285
         Left            =   -74760
         TabIndex        =   92
         Text            =   "cn=server1_Vol1,ou=OrganizationalUnit,o=Container#0#\Users_directory"
         ToolTipText     =   $"frmOptions1.frx":0FC4
         Top             =   1080
         Width           =   6735
      End
      Begin VB.ComboBox Combo199 
         Height          =   315
         ItemData        =   "frmOptions1.frx":108A
         Left            =   -71760
         List            =   "frmOptions1.frx":1367
         TabIndex        =   91
         Top             =   960
         Width           =   735
      End
      Begin VB.ComboBox Combo6 
         Height          =   315
         ItemData        =   "frmOptions1.frx":1737
         Left            =   8520
         List            =   "frmOptions1.frx":1741
         TabIndex        =   90
         Text            =   "NDS 8"
         Top             =   2160
         Width           =   975
      End
      Begin VB.CheckBox UserSetOneCheck 
         Caption         =   "Create User(s)"
         Height          =   255
         Left            =   -73680
         TabIndex        =   89
         ToolTipText     =   "Select the user set below."
         Top             =   720
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.CheckBox UserSetTwoCheck 
         Caption         =   "Create User(s)"
         Height          =   255
         Left            =   -71520
         TabIndex        =   88
         ToolTipText     =   "Select the user set below."
         Top             =   720
         Width           =   1575
      End
      Begin VB.CheckBox UserSetThreeCheck 
         Caption         =   "Create User(s)"
         Height          =   255
         Left            =   -69360
         TabIndex        =   87
         ToolTipText     =   "Select the user set below."
         Top             =   720
         Width           =   1575
      End
      Begin VB.CheckBox UserSetFourCheck 
         Caption         =   "Create User(s)"
         Height          =   255
         Left            =   -67200
         TabIndex        =   86
         ToolTipText     =   "Select the user set below."
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   3240
         TabIndex        =   85
         Text            =   "C:\Temp\ldif_exp.txt"
         Top             =   4080
         Width           =   3255
      End
      Begin VB.CommandButton CmdBrowse3 
         Caption         =   "Browse"
         Height          =   375
         Left            =   6600
         TabIndex        =   84
         Top             =   4080
         Width           =   1335
      End
      Begin VB.CheckBox RetrieveTreeCheck 
         Caption         =   "Retrieve Tree Information"
         Height          =   255
         Left            =   7200
         TabIndex        =   83
         ToolTipText     =   "Selecting this option will gather LDAP information and write it to an LDIF file."
         Top             =   960
         Width           =   2295
      End
      Begin VB.CheckBox CheckAdd 
         Caption         =   "Add"
         Height          =   255
         Left            =   240
         TabIndex        =   82
         ToolTipText     =   "Add Tree and User Information."
         Top             =   3000
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.TextBox Text7 
         Height          =   285
         Left            =   3240
         TabIndex        =   81
         Text            =   "C:\Temp\ldif_imp.txt"
         Top             =   4560
         Width           =   3255
      End
      Begin VB.CommandButton CmdBrowse4 
         Caption         =   "Browse"
         Height          =   375
         Left            =   6600
         TabIndex        =   80
         Top             =   4560
         Width           =   1335
      End
      Begin VB.CheckBox UserHomeCheck 
         Caption         =   "Create User Home Directories?   Select the ""Home Dir Configuration"" Tab and modify the NDS Home Directory Path."
         Height          =   255
         Left            =   240
         TabIndex        =   79
         ToolTipText     =   "Create User Home Directories, select the ""Home Dir Configuration"" Tab and modify the NDS Home Directory Path."
         Top             =   6480
         Width           =   9015
      End
      Begin VB.TextBox Text113 
         Height          =   285
         Left            =   -70920
         TabIndex        =   78
         Text            =   "Internationalization"
         Top             =   5760
         Width           =   1575
      End
      Begin VB.TextBox Text109 
         Height          =   285
         Left            =   -70920
         TabIndex        =   77
         Text            =   "ProtocolEngineering"
         Top             =   5400
         Width           =   1575
      End
      Begin VB.TextBox Text105 
         Height          =   285
         Left            =   -70920
         TabIndex        =   76
         Text            =   "Accounting"
         Top             =   5040
         Width           =   1575
      End
      Begin VB.TextBox Text101 
         Height          =   285
         Left            =   -70920
         TabIndex        =   75
         Text            =   "Engineering"
         Top             =   4680
         Width           =   1575
      End
      Begin VB.ComboBox Combo104 
         Height          =   315
         ItemData        =   "frmOptions1.frx":1753
         Left            =   -74640
         List            =   "frmOptions1.frx":17AB
         Style           =   1  'Simple Combo
         TabIndex        =   74
         Text            =   "Level4Users"
         Top             =   3840
         Width           =   1455
      End
      Begin VB.ComboBox Combo105 
         Height          =   315
         ItemData        =   "frmOptions1.frx":18E9
         Left            =   -74640
         List            =   "frmOptions1.frx":1941
         Style           =   1  'Simple Combo
         TabIndex        =   73
         Text            =   "Level5Users"
         Top             =   4320
         Width           =   1455
      End
      Begin VB.ComboBox Combo106 
         Height          =   315
         ItemData        =   "frmOptions1.frx":1A7F
         Left            =   -74640
         List            =   "frmOptions1.frx":1AD7
         Style           =   1  'Simple Combo
         TabIndex        =   72
         Text            =   "Level6Users"
         Top             =   4800
         Width           =   1455
      End
      Begin VB.ComboBox Combo204 
         Height          =   315
         ItemData        =   "frmOptions1.frx":1C15
         Left            =   -73080
         List            =   "frmOptions1.frx":1C6D
         Style           =   1  'Simple Combo
         TabIndex        =   71
         Text            =   "Level4Users"
         Top             =   3840
         Width           =   1455
      End
      Begin VB.ComboBox Combo205 
         Height          =   315
         ItemData        =   "frmOptions1.frx":1DAB
         Left            =   -73080
         List            =   "frmOptions1.frx":1E03
         Style           =   1  'Simple Combo
         TabIndex        =   70
         Text            =   "Level5Users"
         Top             =   4320
         Width           =   1455
      End
      Begin VB.ComboBox Combo206 
         Height          =   315
         ItemData        =   "frmOptions1.frx":1F41
         Left            =   -73080
         List            =   "frmOptions1.frx":1F99
         Style           =   1  'Simple Combo
         TabIndex        =   69
         Text            =   "Level6Users"
         Top             =   4800
         Width           =   1455
      End
      Begin VB.ComboBox Combo304 
         Height          =   315
         ItemData        =   "frmOptions1.frx":20D7
         Left            =   -71520
         List            =   "frmOptions1.frx":212F
         Style           =   1  'Simple Combo
         TabIndex        =   68
         Text            =   "Level4Users"
         Top             =   3840
         Width           =   1455
      End
      Begin VB.ComboBox Combo305 
         Height          =   315
         ItemData        =   "frmOptions1.frx":226D
         Left            =   -71520
         List            =   "frmOptions1.frx":22C5
         Style           =   1  'Simple Combo
         TabIndex        =   67
         Text            =   "Level5Users"
         Top             =   4320
         Width           =   1455
      End
      Begin VB.ComboBox Combo306 
         Height          =   315
         ItemData        =   "frmOptions1.frx":2403
         Left            =   -71520
         List            =   "frmOptions1.frx":245B
         Style           =   1  'Simple Combo
         TabIndex        =   66
         Text            =   "Level6Users"
         Top             =   4800
         Width           =   1455
      End
      Begin VB.ComboBox Combo404 
         Height          =   315
         ItemData        =   "frmOptions1.frx":2599
         Left            =   -69960
         List            =   "frmOptions1.frx":25F1
         Style           =   1  'Simple Combo
         TabIndex        =   65
         Text            =   "Level4Users"
         Top             =   3840
         Width           =   1455
      End
      Begin VB.ComboBox Combo405 
         Height          =   315
         ItemData        =   "frmOptions1.frx":272F
         Left            =   -69960
         List            =   "frmOptions1.frx":2787
         Style           =   1  'Simple Combo
         TabIndex        =   64
         Text            =   "Level5Users"
         Top             =   4320
         Width           =   1455
      End
      Begin VB.ComboBox Combo406 
         Height          =   315
         ItemData        =   "frmOptions1.frx":28C5
         Left            =   -69960
         List            =   "frmOptions1.frx":291D
         Style           =   1  'Simple Combo
         TabIndex        =   63
         Text            =   "Level6Users"
         Top             =   4800
         Width           =   1455
      End
      Begin VB.ComboBox Combo504 
         Height          =   315
         ItemData        =   "frmOptions1.frx":2A5B
         Left            =   -68400
         List            =   "frmOptions1.frx":2AB3
         Style           =   1  'Simple Combo
         TabIndex        =   62
         Text            =   "Level4Users"
         Top             =   3840
         Width           =   1455
      End
      Begin VB.ComboBox Combo505 
         Height          =   315
         ItemData        =   "frmOptions1.frx":2BF1
         Left            =   -68400
         List            =   "frmOptions1.frx":2C49
         Style           =   1  'Simple Combo
         TabIndex        =   61
         Text            =   "Level5Users"
         Top             =   4320
         Width           =   1455
      End
      Begin VB.ComboBox Combo506 
         Height          =   315
         ItemData        =   "frmOptions1.frx":2D87
         Left            =   -68400
         List            =   "frmOptions1.frx":2DDF
         Style           =   1  'Simple Combo
         TabIndex        =   60
         Text            =   "Level6Users"
         Top             =   4800
         Width           =   1455
      End
      Begin VB.ComboBox Combo604 
         Height          =   315
         ItemData        =   "frmOptions1.frx":2F1D
         Left            =   -66840
         List            =   "frmOptions1.frx":2F75
         Style           =   1  'Simple Combo
         TabIndex        =   59
         Text            =   "Level4FullDepthContainer"
         Top             =   3840
         Width           =   1455
      End
      Begin VB.ComboBox Combo605 
         Height          =   315
         ItemData        =   "frmOptions1.frx":30B3
         Left            =   -66840
         List            =   "frmOptions1.frx":310B
         Style           =   1  'Simple Combo
         TabIndex        =   58
         Text            =   "Level5FullDepthContainer"
         Top             =   4320
         Width           =   1455
      End
      Begin VB.ComboBox Combo606 
         Height          =   315
         ItemData        =   "frmOptions1.frx":3249
         Left            =   -66840
         List            =   "frmOptions1.frx":32A1
         Style           =   1  'Simple Combo
         TabIndex        =   57
         Text            =   "Level6FullDepthContainerFullDepthContainerFullDepth"
         Top             =   4800
         Width           =   1455
      End
      Begin VB.TextBox Text14 
         Height          =   285
         Left            =   3720
         TabIndex        =   56
         Text            =   "o=Novell"
         ToolTipText     =   "Typically the Organizational container. Default ""o=Novell"""
         Top             =   1440
         Width           =   1695
      End
      Begin VB.CheckBox AnonCheck 
         Caption         =   "Anonymous Bind, Non-SSL"
         Height          =   255
         Left            =   4440
         TabIndex        =   55
         ToolTipText     =   "LDAP access without using a username or password, usually fine for searching a container location."
         Top             =   960
         Width           =   2415
      End
      Begin VB.CommandButton cmdViewExp 
         Caption         =   "View Export File"
         Height          =   375
         Left            =   8160
         TabIndex        =   54
         ToolTipText     =   "View the file after creation."
         Top             =   4080
         Width           =   1335
      End
      Begin VB.CommandButton cmdViewImp 
         Caption         =   "View Import File"
         Height          =   375
         Left            =   8160
         TabIndex        =   53
         ToolTipText     =   "View the file after creation. This file will ONLY be created if the Retrieve box has been checked and the Update completed."
         Top             =   4560
         Width           =   1335
      End
      Begin VB.CheckBox CheckDel 
         Caption         =   "Delete"
         Height          =   255
         Left            =   1080
         TabIndex        =   52
         ToolTipText     =   "Remove Tree and User Information."
         Top             =   3000
         Width           =   975
      End
      Begin VB.CheckBox SelectAllUserSetCheck 
         Caption         =   "Select All"
         Height          =   255
         Left            =   -74760
         TabIndex        =   51
         ToolTipText     =   "Select ALL user sets."
         Top             =   720
         Width           =   1095
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "frmOptions1.frx":33DF
         Left            =   3240
         List            =   "frmOptions1.frx":33E1
         Style           =   1  'Simple Combo
         TabIndex        =   50
         Text            =   "636"
         ToolTipText     =   "Port 389 non-SSL, Port 636 SSL."
         Top             =   840
         Width           =   975
      End
      Begin VB.TextBox Text200 
         Height          =   285
         Left            =   -71760
         TabIndex        =   49
         Text            =   "Novell"
         Top             =   1560
         Width           =   1695
      End
      Begin VB.TextBox Text1101 
         Height          =   285
         Left            =   -74640
         TabIndex        =   48
         Text            =   "Boston"
         Top             =   2400
         Width           =   1455
      End
      Begin VB.TextBox Text1201 
         Height          =   285
         Left            =   -73080
         TabIndex        =   47
         Text            =   "Dublin"
         Top             =   2400
         Width           =   1455
      End
      Begin VB.TextBox Text1301 
         Height          =   285
         Left            =   -71520
         TabIndex        =   46
         Text            =   "Provo"
         Top             =   2400
         Width           =   1455
      End
      Begin VB.TextBox Text1401 
         Height          =   285
         Left            =   -69960
         TabIndex        =   45
         Text            =   "Bangalore"
         Top             =   2400
         Width           =   1455
      End
      Begin VB.TextBox Text1501 
         Height          =   285
         Left            =   -68400
         TabIndex        =   44
         Text            =   "Cambridge"
         Top             =   2400
         Width           =   1455
      End
      Begin VB.TextBox Text1601 
         Height          =   285
         Left            =   -66840
         TabIndex        =   43
         Text            =   "Level1DuesseldorfFullDepthContainer"
         Top             =   2400
         Width           =   1455
      End
      Begin VB.TextBox Text1102 
         Height          =   285
         Left            =   -74640
         TabIndex        =   42
         Text            =   "Accounting"
         Top             =   2880
         Width           =   1455
      End
      Begin VB.TextBox Text1202 
         Height          =   285
         Left            =   -73080
         TabIndex        =   41
         Text            =   "Internationalization"
         Top             =   2880
         Width           =   1455
      End
      Begin VB.TextBox Text1302 
         Height          =   285
         Left            =   -71520
         TabIndex        =   40
         Text            =   "Engineering"
         Top             =   2880
         Width           =   1455
      End
      Begin VB.TextBox Text1402 
         Height          =   285
         Left            =   -69960
         TabIndex        =   39
         Text            =   "ProtocolEngineering"
         Top             =   2880
         Width           =   1455
      End
      Begin VB.TextBox Text1502 
         Height          =   285
         Left            =   -68400
         TabIndex        =   38
         Text            =   "Marketing"
         Top             =   2880
         Width           =   1455
      End
      Begin VB.TextBox Text1602 
         Height          =   285
         Left            =   -66840
         TabIndex        =   37
         Text            =   "Level2FullDepthContainer"
         Top             =   2880
         Width           =   1455
      End
      Begin VB.TextBox Text1103 
         Height          =   285
         Left            =   -74640
         TabIndex        =   36
         Text            =   "Users"
         Top             =   3360
         Width           =   1455
      End
      Begin VB.TextBox Text1203 
         Height          =   285
         Left            =   -73080
         TabIndex        =   35
         Text            =   "Users"
         Top             =   3360
         Width           =   1455
      End
      Begin VB.TextBox Text1303 
         Height          =   285
         Left            =   -71520
         TabIndex        =   34
         Text            =   "Users"
         Top             =   3360
         Width           =   1455
      End
      Begin VB.TextBox Text1403 
         Height          =   285
         Left            =   -69960
         TabIndex        =   33
         Text            =   "Users"
         Top             =   3360
         Width           =   1455
      End
      Begin VB.TextBox Text1503 
         Height          =   285
         Left            =   -68400
         TabIndex        =   32
         Text            =   "Users"
         Top             =   3360
         Width           =   1455
      End
      Begin VB.TextBox Text1603 
         Height          =   285
         Left            =   -66840
         TabIndex        =   31
         Text            =   "Level3FullDepthContainer"
         Top             =   3360
         Width           =   1455
      End
      Begin VB.TextBox Text404 
         Height          =   285
         Left            =   -70920
         TabIndex        =   30
         Text            =   "public_html"
         ToolTipText     =   "This is a sub dir under the users dir that will be created. The name used here is for web server home dir testing."
         Top             =   2040
         Width           =   2895
      End
      Begin VB.TextBox Text405 
         Height          =   285
         Left            =   -70920
         TabIndex        =   29
         Text            =   "index.html"
         ToolTipText     =   $"frmOptions1.frx":33E3
         Top             =   2520
         Width           =   2895
      End
      Begin VB.TextBox Text19 
         Height          =   285
         Left            =   3240
         TabIndex        =   28
         Text            =   "C:\Temp\"
         ToolTipText     =   $"frmOptions1.frx":34A0
         Top             =   5040
         Width           =   3255
      End
      Begin VB.CommandButton CmdBrowse5 
         Caption         =   "Browse"
         Height          =   375
         Left            =   6600
         TabIndex        =   27
         ToolTipText     =   "Browse to a location of your choosing to create files using Tree Builder."
         Top             =   5040
         Width           =   1335
      End
      Begin VB.CommandButton cmdViewRice 
         Caption         =   "View ICE Batch"
         Height          =   375
         Left            =   8160
         TabIndex        =   26
         ToolTipText     =   "View the file after creation."
         Top             =   3600
         Width           =   1335
      End
      Begin VB.CommandButton cmdClean 
         Caption         =   "Delete Files"
         Height          =   375
         Left            =   8160
         TabIndex        =   25
         ToolTipText     =   "Cleanup files created by Tree Builder."
         Top             =   5040
         Width           =   1335
      End
      Begin VB.TextBox Text24 
         Height          =   285
         Left            =   -72960
         TabIndex        =   24
         Text            =   "123456789012"
         Top             =   5880
         Width           =   1215
      End
      Begin VB.TextBox Text29 
         Height          =   285
         Left            =   -71520
         TabIndex        =   23
         Text            =   ",ou=Provo,o=novell"
         ToolTipText     =   $"frmOptions1.frx":3543
         Top             =   3360
         Width           =   4095
      End
      Begin VB.CheckBox CustContCheck 
         Caption         =   "Add User Set 1 to this Context"
         Height          =   255
         Left            =   -74760
         TabIndex        =   22
         Top             =   3360
         Width           =   3135
      End
      Begin VB.CommandButton cmdViewCust 
         Caption         =   "View Custom File"
         Height          =   375
         Left            =   8160
         TabIndex        =   21
         ToolTipText     =   "View the contents of the custom file."
         Top             =   5520
         Width           =   1335
      End
      Begin VB.CommandButton CmdBrowse10 
         Caption         =   "Browse"
         Height          =   375
         Left            =   6600
         TabIndex        =   20
         ToolTipText     =   "Browse to the custom LDIF file."
         Top             =   5520
         Width           =   1335
      End
      Begin VB.TextBox Text30 
         Height          =   285
         Left            =   3240
         TabIndex        =   19
         Text            =   "C:\Temp\ldif_cust.txt"
         ToolTipText     =   "Identify an existing LDIF file for LDAP user creation."
         Top             =   5520
         Width           =   3255
      End
      Begin VB.CheckBox CustomCheck 
         Caption         =   "Use This Custom LDIF File"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         ToolTipText     =   $"frmOptions1.frx":35CE
         Top             =   5520
         Width           =   2655
      End
      Begin VB.TextBox Text31 
         Height          =   285
         Left            =   5520
         TabIndex        =   17
         Text            =   "base"
         ToolTipText     =   "Use ""base"" for one level, ""one"" to search one level deep and ""sub"" to search all sub containers."
         Top             =   1440
         Width           =   1695
      End
      Begin VB.CommandButton UHDReset 
         Caption         =   "Reset"
         Height          =   375
         Left            =   -67920
         TabIndex        =   16
         ToolTipText     =   "Reset to the default string. Do NOT change the #0#\."
         Top             =   1080
         Width           =   1335
      End
      Begin VB.CommandButton QS 
         Caption         =   "Quick Start"
         Height          =   375
         Left            =   8520
         TabIndex        =   15
         ToolTipText     =   "Browse to the custom LDIF file."
         Top             =   1320
         Width           =   975
      End
      Begin VB.CommandButton CmdViewICE 
         Caption         =   "View ICE Log"
         Height          =   375
         Left            =   8160
         TabIndex        =   14
         ToolTipText     =   "View the file in the event you closed the dialog display box."
         Top             =   3120
         Width           =   1335
      End
      Begin VB.Label Label54 
         Caption         =   "Rapid ICE Batch File"
         Height          =   255
         Left            =   240
         TabIndex        =   179
         Top             =   3600
         Width           =   1935
      End
      Begin VB.Label Label53 
         Caption         =   "RootCert.der Location"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   178
         Top             =   1800
         Width           =   2055
      End
      Begin VB.Label Label51 
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2640
         TabIndex        =   177
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label50 
         Caption         =   "User Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   176
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label49 
         Caption         =   "LDAP Port"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3240
         TabIndex        =   175
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label48 
         Caption         =   "IP Address of LDAP Server"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   174
         Top             =   600
         Width           =   2415
      End
      Begin VB.Label Label28 
         Caption         =   "Number of Users to Create"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6000
         TabIndex        =   173
         Top             =   1800
         Width           =   2295
      End
      Begin VB.Label Label44 
         Caption         =   "Organization"
         Height          =   255
         Left            =   -67080
         TabIndex        =   172
         Top             =   4320
         Width           =   1695
      End
      Begin VB.Label Label43 
         Caption         =   "Organizational Unit"
         Height          =   255
         Left            =   -69000
         TabIndex        =   171
         Top             =   4320
         Width           =   1455
      End
      Begin VB.Label Label47 
         Caption         =   "Organizational Unit"
         Height          =   255
         Left            =   -72840
         TabIndex        =   170
         Top             =   4320
         Width           =   1695
      End
      Begin VB.Label Label31 
         Caption         =   "Domain Name"
         Height          =   255
         Left            =   -74760
         TabIndex        =   169
         Top             =   6600
         Width           =   1215
      End
      Begin VB.Label Label29 
         Caption         =   "Mail Server Name"
         Height          =   255
         Left            =   -74760
         TabIndex        =   168
         Top             =   6240
         Width           =   1455
      End
      Begin VB.Label Label27 
         Caption         =   "Make All Password Match User Name"
         Height          =   255
         Left            =   -74760
         TabIndex        =   167
         Top             =   2880
         Width           =   3255
      End
      Begin VB.Label Label25 
         Caption         =   "User Set 4"
         Height          =   255
         Left            =   -67200
         TabIndex        =   166
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label Label20 
         Caption         =   "User Set 3"
         Height          =   255
         Left            =   -69360
         TabIndex        =   165
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label Label15 
         Caption         =   "User Set 2"
         Height          =   255
         Left            =   -71520
         TabIndex        =   164
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label Label8 
         Caption         =   "User Set 4, Context 4"
         Height          =   255
         Left            =   -74760
         TabIndex        =   163
         Top             =   5760
         Width           =   2055
      End
      Begin VB.Label Label4 
         Caption         =   "Password"
         Height          =   255
         Left            =   -74760
         TabIndex        =   162
         Top             =   2400
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "User Set 3, Context 3"
         Height          =   255
         Left            =   -74760
         TabIndex        =   161
         Top             =   5400
         Width           =   2055
      End
      Begin VB.Label Label6 
         Caption         =   "User Set 2, Context 2"
         Height          =   255
         Left            =   -74760
         TabIndex        =   160
         Top             =   5040
         Width           =   2055
      End
      Begin VB.Label Label5 
         Caption         =   "User Set 1, Context 1"
         Height          =   255
         Left            =   -74760
         TabIndex        =   159
         Top             =   4680
         Width           =   2055
      End
      Begin VB.Label Label100 
         Caption         =   "User Set 1"
         Height          =   255
         Left            =   -73680
         TabIndex        =   158
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label Label101 
         Caption         =   "Uniquely Definable Data by User Set"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -71520
         TabIndex        =   157
         Top             =   480
         Width           =   3255
      End
      Begin VB.Label Label3 
         Caption         =   "Surname"
         Height          =   255
         Left            =   -74760
         TabIndex        =   156
         Top             =   2040
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "User ID"
         Height          =   255
         Left            =   -74760
         TabIndex        =   155
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Given Name"
         Height          =   255
         Left            =   -74760
         TabIndex        =   154
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label Label45 
         Caption         =   "Select Location for User Home Directories"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74760
         TabIndex        =   153
         Top             =   1560
         Width           =   4335
      End
      Begin VB.Label Label41 
         Caption         =   "NDS Home Directory Path - ONLY modify this line if you are creating user home dirs"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74760
         TabIndex        =   152
         Top             =   840
         Width           =   7215
      End
      Begin VB.Label Label40 
         Caption         =   $"frmOptions1.frx":369E
         Height          =   735
         Left            =   -74760
         TabIndex        =   151
         Top             =   6120
         Width           =   9015
      End
      Begin VB.Label Label64 
         Caption         =   "Organizational Units (OU's)"
         Height          =   255
         Left            =   -71520
         TabIndex        =   150
         Top             =   2040
         Width           =   2295
      End
      Begin VB.Label Label65 
         Caption         =   "Organizational Container"
         Height          =   255
         Left            =   -74400
         TabIndex        =   149
         Top             =   1560
         Width           =   1935
      End
      Begin VB.Label Label66 
         Caption         =   "Country Container"
         Height          =   255
         Left            =   -74400
         TabIndex        =   148
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label Label30 
         Caption         =   "Select Version of NDS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6000
         TabIndex        =   147
         Top             =   2160
         Width           =   2295
      End
      Begin VB.Label Label11 
         Caption         =   "389=non-SSL- 636=SSL"
         Height          =   255
         Left            =   4320
         TabIndex        =   146
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label Label9 
         Caption         =   "LDIF Path and File Name for Export"
         Height          =   255
         Left            =   240
         TabIndex        =   145
         Top             =   4080
         Width           =   2895
      End
      Begin VB.Label Label14 
         Caption         =   "LDIF Path and File Name for Import"
         Height          =   255
         Left            =   240
         TabIndex        =   144
         Top             =   4560
         Width           =   2895
      End
      Begin VB.Label Label16 
         Caption         =   "Organizational Unit"
         Height          =   255
         Left            =   -70920
         TabIndex        =   143
         Top             =   4320
         Width           =   1455
      End
      Begin VB.Label Label17 
         Caption         =   $"frmOptions1.frx":373B
         Height          =   375
         Left            =   -74640
         TabIndex        =   142
         Top             =   5400
         Width           =   8655
      End
      Begin VB.Label Label12 
         Caption         =   "Base DN"
         Height          =   255
         Left            =   3720
         TabIndex        =   141
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label19 
         Caption         =   "User Home Sub-Dir Name"
         Height          =   255
         Left            =   -74760
         TabIndex        =   140
         Top             =   2040
         Width           =   3015
      End
      Begin VB.Label Label21 
         Caption         =   $"frmOptions1.frx":37DF
         Height          =   495
         Left            =   -74760
         TabIndex        =   139
         Top             =   5520
         Width           =   8895
      End
      Begin VB.Label Label22 
         Caption         =   "User Home File Name"
         Height          =   255
         Left            =   -74760
         TabIndex        =   138
         Top             =   2520
         Width           =   3015
      End
      Begin VB.Label Label13 
         Caption         =   "Working Directory"
         Height          =   255
         Left            =   240
         TabIndex        =   137
         Top             =   5040
         Width           =   2895
      End
      Begin VB.Label Label23 
         Caption         =   $"frmOptions1.frx":38BB
         Height          =   495
         Left            =   -74640
         TabIndex        =   136
         Top             =   6240
         Width           =   8655
      End
      Begin VB.Label Label33 
         Caption         =   "Search Scope"
         Height          =   255
         Left            =   5520
         TabIndex        =   135
         Top             =   1200
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Const STILL_ACTIVE = &H103
Const PROCESS_QUERY_INFORMATION = &H400

Private Sub InstallCleanup_Click()
InstCleaner.Show
End Sub

Private Sub HelpRead_Click()
RetVal = ShellExecute(0, "open", "iexplore", "http://st12.provo.novell.com/vol1/TreeBuilder/TreeBuilderReadme.html", "", SW_SHOW)
End Sub

Private Sub QS_Click()
QuickSTart.Show
End Sub

Private Sub CustomCheck_Click()
If CustomCheck = 1 Then RetrieveTreeCheck = 0
If CustomCheck = 1 Then CheckAdd = 0
If CustomCheck = 1 Then CheckDel = 0
End Sub

Private Sub ReadConfig_Click()
MyChar = MsgBox(MyChar, 0, "Here is the character I got.")
End Sub

Private Sub ResetAll_Click()

   Text1.Text = "255.255.255.255"
   Combo1.Text = "636"
   AnonCheck = 0
   RetrieveTreeCheck = 0
   Text2.Text = "cn=admin,o=novell"
   Text3.Text = "test"
   Text14.Text = "o=Novell"
   Text31.Text = "base"
   Combo2.Text = "1"
   Combo6.Text = "NDS 8"
   CheckAdd = 1
   CheckDel = 0
   StopOnErrCheck = 0
   Text4.Text = "F:\Public\Rootcert.der"
   Text5.Text = "C:\temp\rice.bat"
   Text6.Text = "C:\Temp\ldif_exp.txt"
   Text7.Text = "C:\Temp\ldif_imp.txt"
   Text19.Text = "C:\Temp\"
   CustomCheck = 0
   Text30.Text = "C:\Temp\ldif_cust.txt"
   UserHomeCheck = 0
   SelectAllUserSetCheck = 0
   UserSetOneCheck = 1
   UserSetTwoCheck = 0
   UserSetThreeCheck = 0
   UserSetFourCheck = 0
   Text10.Text = "ProvoTestUser"
   Text11.Text = "John"
   Text12.Text = "Doe"
   Text13.Text = "test"
   Text15.Text = "BostonTestUser"
   Text16.Text = "Jack"
   Text17.Text = "Malone"
   Text18.Text = "test"
   Text20.Text = "BangaloreTestUser"
   Text21.Text = "Sudarshan"
   Text22.Text = "Sarkar"
   Text23.Text = "test"
   Text25.Text = "DublinTestUser"
   Text26.Text = "Patrick"
   Text27.Text = "O'Hare"
   Text28.Text = "test"
   Combo14.Text = "No"
   CustContCheck = 0
   Text29.Text = ",ou=Provo,o=novell"
   SkipTreeCreateCheck = 1
   Text100.Text = "Users"
   Text101.Text = "Engineering"
   Text102.Text = "Provo"
   Text103.Text = "Novell"
   Text104.Text = "Users"
   Text105.Text = "Accounting"
   Text106.Text = "Boston"
   Text107.Text = "Novell"
   Text108.Text = "Users"
   Text109.Text = "ProtocolEngineering"
   Text110.Text = "Bangalore"
   Text111.Text = "Novell"
   Text112.Text = "Users"
   Text113.Text = "Internationalization"
   Text114.Text = "Dublin"
   Text115.Text = "Novell"
   Text8.Text = "mh.novell.com"
   Text9.Text = "novell.com"
   Combo199.Text = ""
   Text200.Text = "Novell"
   Text1101.Text = "Boston"
   Text1102.Text = "Accounting"
   Text1103.Text = "Users"
   Combo104.Text = "Level4Users"
   Combo105.Text = "Level5Users"
   Combo106.Text = "Level6Users"
   Text1201.Text = "Dublin"
   Text1202.Text = "Internationalization"
   Text1203.Text = "Users"
   Combo204.Text = "Level4Users"
   Combo205.Text = "Level5Users"
   Combo206.Text = "Level6Users"
   Text1301.Text = "Provo"
   Text1302.Text = "Engineering"
   Text1303.Text = "Users"
   Combo304.Text = "Level4Users"
   Combo305.Text = "Level5Users"
   Combo306.Text = "Level6Users"
   Text1401.Text = "Bangalore"
   Text1402.Text = "ProtocolEngineering"
   Text1403.Text = "Users"
   Combo404.Text = "Level4Users"
   Combo405.Text = "Level5Users"
   Combo406.Text = "Level6Users"
   Text1501.Text = "Cambridge"
   Text1502.Text = "Marketing"
   Text1503.Text = "Users"
   Combo504.Text = "Level4Users"
   Combo505.Text = "Level5Users"
   Combo506.Text = "Level6Users"
   Text1601.Text = "Level1DuesseldorfFullDepthContainer"
   Text1602.Text = "Level2FullDepthContainer"
   Text1603.Text = "Level3FullDepthContainer"
   Combo604.Text = "Level4FullDepthContainer"
   Combo605.Text = "Level5FullDepthContainer"
   Combo606.Text = "Level6FullDepthContainerFullDepthContainerFullDepth"
   Text401.Text = "cn=server1_Vol1,ou=OrganizationalUnit,o=Container#0#\Users_directory"
   Text403.Text = "C:\Users\"
   Text404.Text = "public_html"
   Text405.Text = "index.html"
End Sub

Private Sub SaveConfig_Click()
Dim TBConfigFile As String
Dim TBConf As String
Dim Val As Integer

Val = "3"

TBConfigFile = App.Path + "\" + "TreeBuilder.conf"
   Open TBConfigFile For Output As #1 '***File Open***

            Print #1, "[Tree Builder Version 3.3 Configuration file]"
            Print #1, "[Config]"
            Print #1, "[Config Tab 0]"
            Print #1, "[IP=]'" + Text1.Text + "'"
            Print #1, "[Prt=]'" + Combo1.Text + "'"
            Print #1, "[Uname=]'" + Text2.Text + "'"
            Print #1, "[Pwd=]'" + Text3.Text + "'"
            Print #1, "[BDN=]'" + Text14.Text + "'"
            Print #1, "[Srch=]'" + Text31.Text + "'"
            Print #1, "[RtCrt=]'" + Text4.Text + "'"
            Print #1, "[FOut=]'" + Text6.Text + "'"
            Print #1, "[FIn=]'" + Text7.Text + "'"
            Print #1, "[StopOnErrCheck=]'"; (StopOnErrCheck); "'"
            Print #1, "[CFile=]'" + Text30.Text + "'"
            Print #1, "[AnonCheck=]'"; (AnonCheck); "'"
            Print #1, "[RetrieveTreeCheck=]'"; (RetrieveTreeCheck); "'"
            Print #1, "[endval=]'" + Combo2.Text + "'"
            Print #1, "[NDSVer=]'" + Combo6.Text + "'"
            Print #1, "[CheckAdd=]'"; (CheckAdd); "'"
            Print #1, "[CheckDel=]'"; (CheckDel); "'"
            Print #1, "[IceBatch=]'" + Text5.Text + "'"
            Print #1, "[LDIFImportFile=]'" + Text6.Text + "'"
            Print #1, "[LDIFExportFile=]'" + Text7.Text + "'"
            Print #1, "[WorkDir=]'" + Text19.Text + "'"
            Print #1, "[CustomCheck=]'"; (CustomCheck); "'"
            Print #1, "[LDIFCustomFile=]'" + Text30.Text + "'"
            Print #1, "[CreateUserHomeDir=]'"; (UserHomeCheck); "'"
            Print #1,
              
            Print #1, "[Config Tab 1]"
            Print #1, "[SelectAllUserSetCheck=]'"; (SelectAllUserSetCheck); "'"
            Print #1, "[UserSetOneCheck=]'"; (UserSetOneCheck); "'"
            Print #1, "[UserSetTwoCheck=]'"; (UserSetTwoCheck); "'"
            Print #1, "[UserSetThreeCheck=]'"; (UserSetThreeCheck); "'"
            Print #1, "[UserSetFourCheck=]'"; (UserSetFourCheck); "'"
            Print #1, "[UserSetOne=]'" + Text10.Text + "'"
            Print #1, "[UserSetTwo=]'" + Text15.Text + "'"
            Print #1, "[UserSetThree=]'" + Text20.Text + "'"
            Print #1, "[UserSetFour=]'" + Text25.Text + "'"
            Print #1, "[UserSetOneGN=]'" + Text11.Text + "'"
            Print #1, "[UserSetTwoGN=]'" + Text16.Text + "'"
            Print #1, "[UserSetThreeGN=]'" + Text21.Text + "'"
            Print #1, "[UserSetFourGN=]'" + Text26.Text + "'"
            Print #1, "[UserSetOneSN=]'" + Text12.Text + "'"
            Print #1, "[UserSetTwoSN=]'" + Text17.Text + "'"
            Print #1, "[UserSetThreeSN=]'" + Text22.Text + "'"
            Print #1, "[UserSetFourSN=]'" + Text27.Text + "'"
            Print #1, "[UserSetOnePW=]'" + Text13.Text + "'"
            Print #1, "[UserSetTwoPW=]'" + Text18.Text + "'"
            Print #1, "[UserSetThreePW=]'" + Text23.Text + "'"
            Print #1, "[UserSetFourPW=]'" + Text28.Text + "'"
            Print #1, "[PwdMatch=]'" + Combo14.Text + "'"
            Print #1, "[CustomContainer=]'"; (CustContCheck); "'"
            Print #1, "[CustomContainerName=]'"; Text29.Text + "'"
            Print #1, "[SkipTreeCreate=]'"; (SkipTreeCreateCheck); "'"
            Print #1, "[US111=]'" + Text100.Text + "'"
            Print #1, "[US112=]'" + Text101.Text + "'"
            Print #1, "[US113=]'" + Text102.Text + "'"
            Print #1, "[US114=]'" + Text103.Text + "'"
            Print #1, "[US221=]'" + Text104.Text + "'"
            Print #1, "[US222=]'" + Text105.Text + "'"
            Print #1, "[US223=]'" + Text106.Text + "'"
            Print #1, "[US224=]'" + Text107.Text + "'"
            Print #1, "[US331=]'" + Text108.Text + "'"
            Print #1, "[US332=]'" + Text109.Text + "'"
            Print #1, "[US333=]'" + Text110.Text + "'"
            Print #1, "[US334=]'" + Text111.Text + "'"
            Print #1, "[US441=]'" + Text112.Text + "'"
            Print #1, "[US442=]'" + Text113.Text + "'"
            Print #1, "[US443=]'" + Text114.Text + "'"
            Print #1, "[US444=]'" + Text115.Text + "'"
            Print #1, "[MailServerName=]'" + Text8.Text + "'"
            Print #1, "[DomainName=]'" + Text9.Text + "'"
            Print #1,
            
            Print #1, "[Config Tab 2]"
            Print #1, "[CCont=]'" + Combo199.Text + "'"
            Print #1, "[Org=]'" + Text200.Text + "'"
            Print #1, "[OrgU1=]'" + Text1101.Text + "'"
            Print #1, "[OrgU2=]'" + Text1201.Text + "'"
            Print #1, "[OrgU3=]'" + Text1301.Text + "'"
            Print #1, "[OrgU4=]'" + Text1401.Text + "'"
            Print #1, "[OrgU5=]'" + Text1501.Text + "'"
            Print #1, "[OrgU6=]'" + Text1601.Text + "'"
            Print #1, "[OrgU7=]'" + Text1102.Text + "'"
            Print #1, "[OrgU8=]'" + Text1202.Text + "'"
            Print #1, "[OrgU9=]'" + Text1302.Text + "'"
            Print #1, "[OrgU10=]'" + Text1402.Text + "'"
            Print #1, "[OrgU11=]'" + Text1502.Text + "'"
            Print #1, "[OrgU12=]'" + Text1602.Text + "'"
            Print #1, "[OrgU13=]'" + Text1103.Text + "'"
            Print #1, "[OrgU14=]'" + Text1203.Text + "'"
            Print #1, "[OrgU15=]'" + Text1303.Text + "'"
            Print #1, "[OrgU16=]'" + Text1403.Text + "'"
            Print #1, "[OrgU17=]'" + Text1503.Text + "'"
            Print #1, "[OrgU18=''" + Text1603.Text + "'"
            Print #1, "[OrgU19=]'" + Combo104.Text + "'"
            Print #1, "[OrgU20=]'" + Combo204.Text + "'"
            Print #1, "[OrgU21=]'" + Combo304.Text + "'"
            Print #1, "[OrgU22=]'" + Combo404.Text + "'"
            Print #1, "[OrgU23=]'" + Combo504.Text + "'"
            Print #1, "[OrgU24=]'" + Combo604.Text + "'"
            Print #1, "[OrgU25=]'" + Combo105.Text + "'"
            Print #1, "[OrgU26=]'" + Combo205.Text + "'"
            Print #1, "[OrgU27=]'" + Combo305.Text + "'"
            Print #1, "[OrgU28=]'" + Combo405.Text + "'"
            Print #1, "[OrgU29=]'" + Combo505.Text + "'"
            Print #1, "[OrgU30=]'" + Combo605.Text + "'"
            Print #1, "[OrgU31=]'" + Combo106.Text + "'"
            Print #1, "[OrgU32=]'" + Combo206.Text + "'"
            Print #1, "[OrgU33=]'" + Combo306.Text + "'"
            Print #1, "[OrgU34=]'" + Combo406.Text + "'"
            Print #1, "[OrgU35=]'" + Combo506.Text + "'"
            Print #1, "[OrgU36=]'" + Combo606.Text + "'"
            Print #1,
                  
            Print #1, "[Config Tab 3]"
            Print #1, "[HomeDirNDSPath=]'" + Text401.Text + "'"
            Print #1, "[HomeDirPath=]'" + Text401.Text + "'"
            Print #1, "[UserHomeSubDir=]'" + Text404.Text + "'"
            Print #1, "[UserHomeFileName=]'" + Text405.Text + "'"
            Print #1,
   Close #1
End Sub

Private Sub UHDReset_Click()
Text401.Text = "cn=server1_Vol1,ou=OrganizationalUnit,o=Container#0#\Users_directory"
End Sub

Private Sub UserHomeCheck_Click()
If UserHomeCheck = 1 Then frmOptions.SSTab1.Tab = 3
End Sub

Private Sub StopOnErrCheck_Click()
If StopOnErrCheck = 1 Then StopOnErr = ""
If StopOnErrCheck = 0 Then StopOnErr = "-c"
End Sub

Private Sub RetrieveTreeCheck_Click()
If RetrieveTreeCheck = 1 Then CustomCheck = 0
If RetrieveTreeCheck = 1 Then CheckAdd = 0
If RetrieveTreeCheck = 1 Then CheckDel = 0
End Sub

Private Sub CheckAdd_Click()
If CheckAdd = 1 Then CheckDel = 0
If CheckAdd = 1 Then RetrieveTreeCheck = 0
If CheckAdd = 1 Then CType = "add"
If CheckAdd = 0 Then CType = "delete"
If CheckAdd = 1 Then CustomCheck = 0
End Sub

Private Sub CheckDel_Click()
If CheckDel = 1 Then CheckAdd = 0
If CheckDel = 1 Then RetrieveTreeCheck = 0
If CheckDel = 1 Then CustomCheck = 0
End Sub

Private Sub AnonCheck_Click()
If AnonCheck = 1 Then Combo1.Text = "389"
If AnonCheck = 1 Then Text2.Text = ""
If AnonCheck = 1 Then Text3.Text = ""
If AnonCheck = 1 Then Text4.Text = ""
If AnonCheck = 1 Then Text14.Text = ""
If AnonCheck = 1 Then Text31.Text = ""

'Reset Defaults
If AnonCheck = 0 Then Combo1.Text = "636"
If AnonCheck = 0 Then Text2.Text = "cn=admin,o=novell"
If AnonCheck = 0 Then Text3.Text = "test"
If AnonCheck = 0 Then Text4.Text = "F:\Public\Rootcert.der"
If AnonCheck = 0 Then Text14.Text = "o=Novell"
If AnonCheck = 0 Then Text31.Text = "base"
End Sub

Private Sub SelectAllUserSetCheck_Click()
If SelectAllUserSetCheck = 1 Then UserSetOneCheck = 1
If SelectAllUserSetCheck = 1 Then UserSetTwoCheck = 1
If SelectAllUserSetCheck = 1 Then UserSetThreeCheck = 1
If SelectAllUserSetCheck = 1 Then UserSetFourCheck = 1
If SelectAllUserSetCheck = 0 Then UserSetOneCheck = 0
If SelectAllUserSetCheck = 0 Then UserSetTwoCheck = 0
If SelectAllUserSetCheck = 0 Then UserSetThreeCheck = 0
If SelectAllUserSetCheck = 0 Then UserSetFourCheck = 0
End Sub

Private Sub CustContCheck_Click()
If CustContCheck = 1 Then SelectAllUserSetCheck = 0
If CustContCheck = 1 Then UserSetOneCheck = 1
If CustContCheck = 1 Then UserSetTwoCheck = 0
If CustContCheck = 1 Then UserSetThreeCheck = 0
If CustContCheck = 1 Then UserSetFourCheck = 0
If CustContCheck = 0 Then SelectAllUserSetCheck = 0
If CustContCheck = 0 Then UserSetOneCheck = 1
If CustContCheck = 0 Then UserSetTwoCheck = 0
If CustContCheck = 0 Then UserSetThreeCheck = 0
If CustContCheck = 0 Then UserSetFourCheck = 0
End Sub

Private Sub CmdBrowse1_Click()
cmnDialog.FileName = Text4.Text
cmnDialog.ShowOpen
Text4.Text = cmnDialog.FileName
If Text4.Text = "" Then Text4.Text = "F:\Public\Rootcert.der"
End Sub

Private Sub CmdBrowse10_Click()
cmnDialog.FileName = Text30.Text
cmnDialog.ShowOpen
Text30.Text = cmnDialog.FileName
If Text30.Text = "" Then Text30.Text = "C:\Temp\ldif_cust.txt"
End Sub

Private Sub CmdBrowse2_Click()
cmnDialog.FileName = Text5.Text
cmnDialog.ShowOpen
Text5.Text = cmnDialog.FileName
If Text5.Text = "" Then Text5.Text = "C:\temp\rice.bat"
End Sub

Private Sub CmdBrowse3_Click()
cmnDialog.FileName = Text6.Text
cmnDialog.ShowOpen
Text6.Text = cmnDialog.FileName
If Text6.Text = "" Then Text6.Text = "C:\Temp\ldif_exp.txt"
End Sub

Private Sub CmdBrowse4_Click()
cmnDialog.FileName = Text7.Text
cmnDialog.ShowOpen
Text7.Text = cmnDialog.FileName
If Text7.Text = "" Then Text7.Text = "C:\Temp\ldif_imp.txt"
End Sub

Private Sub CmdBrowse6_Click()
Dim getdir As String
    getdir = Text403.Text
        getdir = BrowseForFolder(Me, "Select A Directory to Create User Home Directories in", getdir)
    If Len(getdir) = 0 Then Exit Sub  'user selected cancel
    Text403.Text = getdir
End Sub

Private Sub cmdClean_Click()
   Dim RICEBatch As String
   Dim fileout As String
   Dim filein As String
   Dim RetVal As Long
   
   RICEBatch = Text5.Text
   fileout = Text6.Text
   filein = Text7.Text
   
   RetVal = DeleteFile(RICEBatch)
   RetVal = DeleteFile(fileout)
   RetVal = DeleteFile(filein)
   RetVal = DeleteFile("ice.log")
   
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdViewCust_Click()
      Dim RetVal As Long
      Dim filein As String
           
      filein = Text30.Text
      RetVal = ShellExecute(0, "open", "notepad", filein, "", SW_SHOW)
End Sub

Private Sub cmdViewExp_Click()
      Dim RetVal As Long
      Dim fileout As String
      
      fileout = Text6.Text
      
      RetVal = ShellExecute(0, "open", "notepad", fileout, "", SW_SHOW)
End Sub

Private Sub CmdViewICE_Click()
      Dim RetVal As Long
      Dim ifilename As String
      Dim WPath
      WPath = Text19.Text
      ifilename = WPath + "ice.log"

      RetVal = ShellExecute(0, "open", "notepad", ifilename, "", SW_SHOW)
End Sub

Private Sub cmdViewImp_Click()
      Dim RetVal As Long
      Dim filein As String
           
      filein = Text7.Text
      RetVal = ShellExecute(0, "open", "notepad", filein, "", SW_SHOW)
End Sub

Private Sub CmdBrowse5_Click()
    Dim getdir As String
    getdir = Text19.Text
        getdir = BrowseForFolder(Me, "Select A Directory to write working files to.", getdir)
    If Len(getdir) = 0 Then Exit Sub  'user selected cancel
    Text19.Text = getdir
End Sub

Private Sub cmdViewRice_Click()
      Dim RetVal As Long
      Dim RICEBatch As String
            
      RICEBatch = Text5.Text
      RetVal = ShellExecute(0, "open", "notepad", RICEBatch, "", SW_SHOW)
End Sub

Private Sub Combo1_Change()
If Combo1.Text = "389" Then Text4.Text = ""
End Sub

Private Sub Text24_Change()
Text24.Text = "123456789012"
End Sub

Private Sub Text100_Change()
Text1303.Text = Text100.Text
End Sub

Private Sub Text101_Change()
Text1302.Text = Text101.Text
End Sub

Private Sub Text102_Change()
Text1301.Text = Text102.Text
End Sub

Private Sub Text104_Change()
Text1103.Text = Text104.Text
End Sub

Private Sub Text105_Change()
Text1102.Text = Text105.Text
End Sub

Private Sub Text106_Change()
Text1101.Text = Text106.Text
End Sub

Private Sub Text108_Change()
Text1403.Text = Text108.Text
End Sub

Private Sub Text109_Change()
Text1402.Text = Text109.Text
End Sub

Private Sub Text110_Change()
Text1401.Text = Text110.Text
End Sub

Private Sub Text112_Change()
Text1203.Text = Text112.Text
End Sub

Private Sub Text113_Change()
Text1202.Text = Text113.Text
End Sub

Private Sub Text114_Change()
Text1201.Text = Text114.Text
End Sub

Private Sub Text1101_Change()
Text106.Text = Text1101.Text
End Sub

Private Sub Text1102_Change()
Text105.Text = Text1102.Text
End Sub

Private Sub Text1103_Change()
Text104.Text = Text1103.Text
End Sub

Private Sub Text1201_Change()
Text114.Text = Text1201.Text
End Sub

Private Sub Text1202_Change()
Text113.Text = Text1202.Text
End Sub

Private Sub Text1203_Change()
Text112.Text = Text1203.Text
End Sub

Private Sub Text1301_Change()
Text102.Text = Text1301.Text
End Sub

Private Sub Text1302_Change()
Text101.Text = Text1302.Text
End Sub

Private Sub Text1303_Change()
Text100.Text = Text1303.Text
End Sub

Private Sub Text1401_Change()
Text110.Text = Text1401.Text
End Sub

Private Sub Text1402_Change()
Text109.Text = Text1402.Text
End Sub

Private Sub Text1403_Change()
Text108.Text = Text1403.Text
End Sub

Private Sub Text200_Change()
Text103.Text = Text200.Text
Text107.Text = Text200.Text
Text111.Text = Text200.Text
Text115.Text = Text200.Text
End Sub

Private Sub Text103_Change()
Text200.Text = Text103.Text
Text107.Text = Text103.Text
Text111.Text = Text103.Text
Text115.Text = Text103.Text
End Sub

Private Sub Text107_Change()
Text200.Text = Text107.Text
Text103.Text = Text107.Text
Text111.Text = Text107.Text
Text115.Text = Text107.Text
End Sub

Private Sub Text111_Change()
Text200.Text = Text111.Text
Text103.Text = Text111.Text
Text107.Text = Text111.Text
Text115.Text = Text111.Text
End Sub

Private Sub Text115_Change()
Text200.Text = Text115.Text
Text103.Text = Text115.Text
Text107.Text = Text115.Text
Text111.Text = Text115.Text
End Sub

Private Sub CmdUpdate_Click()
   
   Dim RICEBatch As String
   Dim fileout As String
   Dim filein As String
   Dim beginval As Long
   Dim endval As Long
   Dim ver As String
   Dim servername As String
   Dim domainname As String
   Dim passwordm As String
   Dim GetLDAP As String
   Dim BaseDN As String
   Dim homedir As String
   Dim userid1 As String
   Dim userid2 As String
   Dim userid3 As String
   Dim userid4 As String
   Dim givenname1 As String
   Dim givenname2 As String
   Dim givenname3 As String
   Dim givenname4 As String
   Dim surname1 As String
   Dim surname2 As String
   Dim surname3 As String
   Dim surname4 As String
   Dim password1 As String
   Dim password2 As String
   Dim password3 As String
   Dim password4 As String
   Dim CType As String
   Dim Org1 As String
   Dim Org2 As String
   Dim Org3 As String
   Dim Org4 As String
   Dim OrgUnit1 As String
   Dim OrgUnit2 As String
   Dim OrgUnit3 As String
   Dim OrgUnit4 As String
   Dim OrgUnit11 As String
   Dim OrgUnit22 As String
   Dim OrgUnit33 As String
   Dim OrgUnit44 As String
   Dim OrgUnit111 As String
   Dim OrgUnit222 As String
   Dim OrgUnit333 As String
   Dim OrgUnit444 As String
   Dim CCont As String
   Dim Org As String
   Dim OrgU As String
   Dim OrgU1 As String
   Dim OrgU2 As String
   Dim OrgU3 As String
   Dim OrgU4 As String
   Dim OrgU5 As String
   Dim OrgU6 As String
   Dim OrgU7 As String
   Dim OrgU8 As String
   Dim OrgU9 As String
   Dim OrgU10 As String
   Dim OrgU11 As String
   Dim OrgU12 As String
   Dim OrgU13 As String
   Dim OrgU14 As String
   Dim OrgU15 As String
   Dim OrgU16 As String
   Dim OrgU17 As String
   Dim OrgU18 As String
   Dim OrgU19 As String
   Dim OrgU20 As String
   Dim OrgU21 As String
   Dim OrgU22 As String
   Dim OrgU23 As String
   Dim OrgU24 As String
   Dim OrgU25 As String
   Dim OrgU26 As String
   Dim OrgU27 As String
   Dim OrgU28 As String
   Dim OrgU29 As String
   Dim OrgU30 As String
   Dim OrgU31 As String
   Dim OrgU32 As String
   Dim OrgU33 As String
   Dim OrgU34 As String
   Dim OrgU35 As String
   Dim OrgU36 As String
   Dim RetVal As Long
   Dim UHomeDir As String
   Dim UHomeFile As String
   Dim CustCont As String
   Dim CustLDIF As String
   Dim MyPath
   Dim WPath
   Dim hConsole As Long
      
   RICEBatch = Text5.Text
   fileout = Text6.Text
   filein = Text7.Text
   servername = Text8.Text
   domainname = Text9.Text
   beginval = 1
   endval = Val(Combo2.Text)
   Port = Combo1.Text
   passwordm = Combo14.Text
   ver = Combo6.Text
   WPath = Text19.Text
   homedir = Text401.Text
   MyPath = Text403.Text
   UHomeDir = Text404.Text
   UHomeFile = Text405.Text
   userid1 = Text10.Text
   userid2 = Text15.Text
   userid3 = Text20.Text
   userid4 = Text25.Text
   givenname1 = Text11.Text
   givenname2 = Text16.Text
   givenname3 = Text21.Text
   givenname4 = Text26.Text
   surname1 = Text12.Text
   surname2 = Text17.Text
   surname3 = Text22.Text
   surname4 = Text27.Text
   password1 = Text13.Text
   password2 = Text18.Text
   password3 = Text23.Text
   password4 = Text28.Text
   Org1 = Text103.Text
   Org2 = Text107.Text
   Org3 = Text111.Text
   Org4 = Text115.Text
   OrgUnit1 = Text102.Text
   OrgUnit2 = Text106.Text
   OrgUnit3 = Text110.Text
   OrgUnit4 = Text114.Text
   OrgUnit11 = Text101.Text
   OrgUnit22 = Text105.Text
   OrgUnit33 = Text109.Text
   OrgUnit44 = Text113.Text
   OrgUnit111 = Text100.Text
   OrgUnit222 = Text104.Text
   OrgUnit333 = Text108.Text
   OrgUnit444 = Text112.Text
   CCont = Combo199.Text
   Org = Text200.Text
   OrgU1 = Text1101.Text
   OrgU2 = Text1201.Text
   OrgU3 = Text1301.Text
   OrgU4 = Text1401.Text
   OrgU5 = Text1501.Text
   OrgU6 = Text1601.Text
   OrgU7 = Text1102.Text
   OrgU8 = Text1202.Text
   OrgU9 = Text1302.Text
   OrgU10 = Text1402.Text
   OrgU11 = Text1502.Text
   OrgU12 = Text1602.Text
   OrgU13 = Text1103.Text
   OrgU14 = Text1203.Text
   OrgU15 = Text1303.Text
   OrgU16 = Text1403.Text
   OrgU17 = Text1503.Text
   OrgU18 = Text1603.Text
   OrgU19 = Combo104.Text
   OrgU20 = Combo204.Text
   OrgU21 = Combo304.Text
   OrgU22 = Combo404.Text
   OrgU23 = Combo504.Text
   OrgU24 = Combo604.Text
   OrgU25 = Combo105.Text
   OrgU26 = Combo205.Text
   OrgU27 = Combo305.Text
   OrgU28 = Combo405.Text
   OrgU29 = Combo505.Text
   OrgU30 = Combo605.Text
   OrgU31 = Combo106.Text
   OrgU32 = Combo206.Text
   OrgU33 = Combo306.Text
   OrgU34 = Combo406.Text
   OrgU35 = Combo506.Text
   OrgU36 = Combo606.Text
   CustCont = Text29.Text
   CustLDIF = Text30.Text
   
   On Error GoTo ErrorHandler '*** Start up the error handler
   
   If CheckAdd = 1 Then CType = "add"
   If CheckDel = 1 Then CType = "delete"
                       
           Dim fso, Msg
           Set fso = CreateObject("Scripting.FileSystemObject")
           If (fso.FolderExists(WPath)) Then GoTo CStep180 Else MkDir WPath  'Make new directory or folder.
CStep180:
          ChDrive WPath 'Changes the current drive.
          ChDir WPath 'Changes the current directory or folder.
                              
        If RetrieveTreeCheck = 1 Then GoTo Step2

If CustomCheck = 1 Then GoTo CustLDIFStep1 'If Custom LDIF file is selected bypass the file creation.
                 
   Open fileout For Output As #1 '***File Open***
   
      Print #1, "#This file generated by Novell's Tree Builder Version 3. (Written by: Robert Foster)" 'File Header
      Print #1, "version: 1"
      If CheckDel = 0 Then Print #1,
   
If SkipTreeCreateCheck = 0 Then GoTo BypassTree

'***Need to add code to allow numbers to start at something other than zero at some point***

   If endval = "1" Then GoTo Step4 'If endval is 1 this goes around the Progress Bar Error
   ProgressBar1.Min = beginval
   ProgressBar1.Max = endval
   ProgressBar1.Value = ProgressBar1.Min
   ProgressBar1.Visible = False
         
Step4:
   
If CType = "Delete" Then GoTo Step41

'Start Add Tree Information

      If CCont = "" Then GoTo Step22
      If CType = "delete" Then GoTo Step22

      Print #1, "dn: c=" + CCont
      Print #1, "changetype: " + CType
      Print #1, "objectClass: top"
      Print #1, "objectClass: country"
      If ver = "NDS 8" Then Print #1, "c: " + CCont
      If ver = "NDS 8" Then Print #1, "objectClass: top"
      Print #1,

Step22:

If CheckDel = 1 Then GoTo BypassTreeAdd

      If Not CCont = "" Then Print #1, "dn: o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: o=" + Org
      Print #1, "changetype: " + CType
      If ver = "NDS 7" Then Print #1, "objectClass: top"
      If ver = "NDS 8" Then Print #1, "o: " + Org
      Print #1, "objectClass: organization"
      If ver = "NDS 8" Then Print #1, "objectClass: ndsLoginProperties"
      If ver = "NDS 8" Then Print #1, "objectClass: top"
      Print #1,
      
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU1 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU1 + ",o=" + Org
      Print #1, "changetype: " + CType
      If ver = "NDS 7" Then Print #1, "objectClass: top"
      If ver = "NDS 8" Then Print #1, "ou: " + OrgU1
      Print #1, "objectClass: organizationalUnit"
      If ver = "NDS 8" Then Print #1, "objectClass: ndsLoginProperties"
      If ver = "NDS 8" Then Print #1, "objectClass: top"
      Print #1,
      
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU2 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU2 + ",o=" + Org
      Print #1, "changetype: " + CType
      If ver = "NDS 7" Then Print #1, "objectClass: top"
      If ver = "NDS 8" Then Print #1, "ou: " + OrgU2
      Print #1, "objectClass: organizationalUnit"
      If ver = "NDS 8" Then Print #1, "objectClass: ndsLoginProperties"
      If ver = "NDS 8" Then Print #1, "objectClass: top"
      Print #1,
      
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU3 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU3 + ",o=" + Org
      Print #1, "changetype: " + CType
      If ver = "NDS 7" Then Print #1, "objectClass: top"
      If ver = "NDS 8" Then Print #1, "ou: " + OrgU3
      Print #1, "objectClass: organizationalUnit"
      If ver = "NDS 8" Then Print #1, "objectClass: ndsLoginProperties"
      If ver = "NDS 8" Then Print #1, "objectClass: top"
      Print #1,
      
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU4 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU4 + ",o=" + Org
      Print #1, "changetype: " + CType
      If ver = "NDS 7" Then Print #1, "objectClass: top"
      If ver = "NDS 8" Then Print #1, "ou: " + OrgU4
      Print #1, "objectClass: organizationalUnit"
      If ver = "NDS 8" Then Print #1, "objectClass: ndsLoginProperties"
      If ver = "NDS 8" Then Print #1, "objectClass: top"
      Print #1,
   
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU5 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU5 + ",o=" + Org
      Print #1, "changetype: " + CType
      If ver = "NDS 7" Then Print #1, "objectClass: top"
      If ver = "NDS 8" Then Print #1, "ou: " + OrgU5
      Print #1, "objectClass: organizationalUnit"
      If ver = "NDS 8" Then Print #1, "objectClass: ndsLoginProperties"
      If ver = "NDS 8" Then Print #1, "objectClass: top"
      Print #1,
   
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU6 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU6 + ",o=" + Org
      Print #1, "changetype: " + CType
      If ver = "NDS 7" Then Print #1, "objectClass: top"
      If ver = "NDS 8" Then Print #1, "ou: " + OrgU6
      Print #1, "objectClass: organizationalUnit"
      If ver = "NDS 8" Then Print #1, "objectClass: ndsLoginProperties"
      If ver = "NDS 8" Then Print #1, "objectClass: top"
      Print #1,
            
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU7 + ",ou=" + OrgU1 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU7 + ",ou=" + OrgU1 + ",o=" + Org
      Print #1, "changetype: " + CType
      If ver = "NDS 7" Then Print #1, "objectClass: top"
      If ver = "NDS 8" Then Print #1, "ou: " + OrgU7
      Print #1, "objectClass: organizationalUnit"
      If ver = "NDS 8" Then Print #1, "objectClass: ndsLoginProperties"
      If ver = "NDS 8" Then Print #1, "objectClass: top"
      Print #1,
      
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU8 + ",ou=" + OrgU2 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU8 + ",ou=" + OrgU2 + ",o=" + Org
      Print #1, "changetype: " + CType
      If ver = "NDS 7" Then Print #1, "objectClass: top"
      If ver = "NDS 8" Then Print #1, "ou: " + OrgU8
      Print #1, "objectClass: organizationalUnit"
      If ver = "NDS 8" Then Print #1, "objectClass: ndsLoginProperties"
      If ver = "NDS 8" Then Print #1, "objectClass: top"
      Print #1,
      
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU9 + ",ou=" + OrgU3 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU9 + ",ou=" + OrgU3 + ",o=" + Org
      Print #1, "changetype: " + CType
      If ver = "NDS 7" Then Print #1, "objectClass: top"
      If ver = "NDS 8" Then Print #1, "ou: " + OrgU9
      Print #1, "objectClass: organizationalUnit"
      If ver = "NDS 8" Then Print #1, "objectClass: ndsLoginProperties"
      If ver = "NDS 8" Then Print #1, "objectClass: top"
      Print #1,
      
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU10 + ",ou=" + OrgU4 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU10 + ",ou=" + OrgU4 + ",o=" + Org
      Print #1, "changetype: " + CType
      If ver = "NDS 7" Then Print #1, "objectClass: top"
      If ver = "NDS 8" Then Print #1, "ou: " + OrgU10
      Print #1, "objectClass: organizationalUnit"
      If ver = "NDS 8" Then Print #1, "objectClass: ndsLoginProperties"
      If ver = "NDS 8" Then Print #1, "objectClass: top"
      Print #1,
            
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU11 + ",ou=" + OrgU5 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU11 + ",ou=" + OrgU5 + ",o=" + Org
      Print #1, "changetype: " + CType
      If ver = "NDS 7" Then Print #1, "objectClass: top"
      If ver = "NDS 8" Then Print #1, "ou: " + OrgU11
      Print #1, "objectClass: organizationalUnit"
      If ver = "NDS 8" Then Print #1, "objectClass: ndsLoginProperties"
      If ver = "NDS 8" Then Print #1, "objectClass: top"
      Print #1,
            
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU12 + ",ou=" + OrgU6 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU12 + ",ou=" + OrgU6 + ",o=" + Org
      Print #1, "changetype: " + CType
      If ver = "NDS 7" Then Print #1, "objectClass: top"
      If ver = "NDS 8" Then Print #1, "ou: " + OrgU12
      Print #1, "objectClass: organizationalUnit"
      If ver = "NDS 8" Then Print #1, "objectClass: ndsLoginProperties"
      If ver = "NDS 8" Then Print #1, "objectClass: top"
      Print #1,
      
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU13 + ",ou=" + OrgU7 + ",ou=" + OrgU1 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU13 + ",ou=" + OrgU7 + ",ou=" + OrgU1 + ",o=" + Org
      Print #1, "changetype: " + CType
      If ver = "NDS 7" Then Print #1, "objectClass: top"
      If ver = "NDS 8" Then Print #1, "ou: " + OrgU13
      Print #1, "objectClass: organizationalUnit"
      If ver = "NDS 8" Then Print #1, "objectClass: ndsLoginProperties"
      If ver = "NDS 8" Then Print #1, "objectClass: top"
      Print #1,
      
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU14 + ",ou=" + OrgU8 + ",ou=" + OrgU2 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU14 + ",ou=" + OrgU8 + ",ou=" + OrgU2 + ",o=" + Org
      Print #1, "changetype: " + CType
      If ver = "NDS 7" Then Print #1, "objectClass: top"
      If ver = "NDS 8" Then Print #1, "ou: " + OrgU14
      Print #1, "objectClass: organizationalUnit"
      If ver = "NDS 8" Then Print #1, "objectClass: ndsLoginProperties"
      If ver = "NDS 8" Then Print #1, "objectClass: top"
      Print #1,
      
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU15 + ",ou=" + OrgU9 + ",ou=" + OrgU3 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU15 + ",ou=" + OrgU9 + ",ou=" + OrgU3 + ",o=" + Org
      Print #1, "changetype: " + CType
      If ver = "NDS 7" Then Print #1, "objectClass: top"
      If ver = "NDS 8" Then Print #1, "ou: " + OrgU15
      Print #1, "objectClass: organizationalUnit"
      If ver = "NDS 8" Then Print #1, "objectClass: ndsLoginProperties"
      If ver = "NDS 8" Then Print #1, "objectClass: top"
      Print #1,
      
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU16 + ",ou=" + OrgU10 + ",ou=" + OrgU4 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU16 + ",ou=" + OrgU10 + ",ou=" + OrgU4 + ",o=" + Org
      Print #1, "changetype: " + CType
      If ver = "NDS 7" Then Print #1, "objectClass: top"
      If ver = "NDS 8" Then Print #1, "ou: " + OrgU16
      Print #1, "objectClass: organizationalUnit"
      If ver = "NDS 8" Then Print #1, "objectClass: ndsLoginProperties"
      If ver = "NDS 8" Then Print #1, "objectClass: top"
      Print #1,
      
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU17 + ",ou=" + OrgU11 + ",ou=" + OrgU5 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU17 + ",ou=" + OrgU11 + ",ou=" + OrgU5 + ",o=" + Org
      Print #1, "changetype: " + CType
      If ver = "NDS 7" Then Print #1, "objectClass: top"
      If ver = "NDS 8" Then Print #1, "ou: " + OrgU17
      Print #1, "objectClass: organizationalUnit"
      If ver = "NDS 8" Then Print #1, "objectClass: ndsLoginProperties"
      If ver = "NDS 8" Then Print #1, "objectClass: top"
      Print #1,
      
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU18 + ",ou=" + OrgU12 + ",ou=" + OrgU6 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU18 + ",ou=" + OrgU12 + ",ou=" + OrgU6 + ",o=" + Org
      Print #1, "changetype: " + CType
      If ver = "NDS 7" Then Print #1, "objectClass: top"
      If ver = "NDS 8" Then Print #1, "ou: " + OrgU18
      Print #1, "objectClass: organizationalUnit"
      If ver = "NDS 8" Then Print #1, "objectClass: ndsLoginProperties"
      If ver = "NDS 8" Then Print #1, "objectClass: top"
      Print #1,
      
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU19 + ",ou=" + OrgU13 + ",ou=" + OrgU7 + ",ou=" + OrgU1 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU19 + ",ou=" + OrgU13 + ",ou=" + OrgU7 + ",ou=" + OrgU1 + ",o=" + Org
      Print #1, "changetype: " + CType
      If ver = "NDS 7" Then Print #1, "objectClass: top"
      If ver = "NDS 8" Then Print #1, "ou: " + OrgU19
      Print #1, "objectClass: organizationalUnit"
      If ver = "NDS 8" Then Print #1, "objectClass: ndsLoginProperties"
      If ver = "NDS 8" Then Print #1, "objectClass: top"
      Print #1,
      
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU20 + ",ou=" + OrgU14 + ",ou=" + OrgU8 + ",ou=" + OrgU2 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU20 + ",ou=" + OrgU14 + ",ou=" + OrgU8 + ",ou=" + OrgU2 + ",o=" + Org
      Print #1, "changetype: " + CType
      If ver = "NDS 7" Then Print #1, "objectClass: top"
      If ver = "NDS 8" Then Print #1, "ou: " + OrgU20
      Print #1, "objectClass: organizationalUnit"
      If ver = "NDS 8" Then Print #1, "objectClass: ndsLoginProperties"
      If ver = "NDS 8" Then Print #1, "objectClass: top"
      Print #1,
      
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU21 + ",ou=" + OrgU15 + ",ou=" + OrgU9 + ",ou=" + OrgU3 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU21 + ",ou=" + OrgU15 + ",ou=" + OrgU9 + ",ou=" + OrgU3 + ",o=" + Org
      Print #1, "changetype: " + CType
      If ver = "NDS 7" Then Print #1, "objectClass: top"
      If ver = "NDS 8" Then Print #1, "ou: " + OrgU21
      Print #1, "objectClass: organizationalUnit"
      If ver = "NDS 8" Then Print #1, "objectClass: ndsLoginProperties"
      If ver = "NDS 8" Then Print #1, "objectClass: top"
      Print #1,
      
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU22 + ",ou=" + OrgU16 + ",ou=" + OrgU10 + ",ou=" + OrgU4 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU22 + ",ou=" + OrgU16 + ",ou=" + OrgU10 + ",ou=" + OrgU4 + ",o=" + Org
      Print #1, "changetype: " + CType
      If ver = "NDS 7" Then Print #1, "objectClass: top"
      If ver = "NDS 8" Then Print #1, "ou: " + OrgU22
      Print #1, "objectClass: organizationalUnit"
      If ver = "NDS 8" Then Print #1, "objectClass: ndsLoginProperties"
      If ver = "NDS 8" Then Print #1, "objectClass: top"
      Print #1,
      
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU23 + ",ou=" + OrgU17 + ",ou=" + OrgU11 + ",ou=" + OrgU5 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU23 + ",ou=" + OrgU17 + ",ou=" + OrgU11 + ",ou=" + OrgU5 + ",o=" + Org
      Print #1, "changetype: " + CType
      If ver = "NDS 7" Then Print #1, "objectClass: top"
      If ver = "NDS 8" Then Print #1, "ou: " + OrgU23
      Print #1, "objectClass: organizationalUnit"
      If ver = "NDS 8" Then Print #1, "objectClass: ndsLoginProperties"
      If ver = "NDS 8" Then Print #1, "objectClass: top"
      Print #1,
      
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU24 + ",ou=" + OrgU18 + ",ou=" + OrgU12 + ",ou=" + OrgU6 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU24 + ",ou=" + OrgU18 + ",ou=" + OrgU12 + ",ou=" + OrgU6 + ",o=" + Org
      Print #1, "changetype: " + CType
      If ver = "NDS 7" Then Print #1, "objectClass: top"
      If ver = "NDS 8" Then Print #1, "ou: " + OrgU24
      Print #1, "objectClass: organizationalUnit"
      If ver = "NDS 8" Then Print #1, "objectClass: ndsLoginProperties"
      If ver = "NDS 8" Then Print #1, "objectClass: top"
      Print #1,
      
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU25 + ",ou=" + OrgU19 + ",ou=" + OrgU13 + ",ou=" + OrgU7 + ",ou=" + OrgU1 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU25 + ",ou=" + OrgU19 + ",ou=" + OrgU13 + ",ou=" + OrgU7 + ",ou=" + OrgU1 + ",o=" + Org
      Print #1, "changetype: " + CType
      If ver = "NDS 7" Then Print #1, "objectClass: top"
      If ver = "NDS 8" Then Print #1, "ou: " + OrgU25
      Print #1, "objectClass: organizationalUnit"
      If ver = "NDS 8" Then Print #1, "objectClass: ndsLoginProperties"
      If ver = "NDS 8" Then Print #1, "objectClass: top"
      Print #1,
      
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU26 + ",ou=" + OrgU20 + ",ou=" + OrgU14 + ",ou=" + OrgU8 + ",ou=" + OrgU2 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU26 + ",ou=" + OrgU20 + ",ou=" + OrgU14 + ",ou=" + OrgU8 + ",ou=" + OrgU2 + ",o=" + Org
      Print #1, "changetype: " + CType
      If ver = "NDS 7" Then Print #1, "objectClass: top"
      If ver = "NDS 8" Then Print #1, "ou: " + OrgU26
      Print #1, "objectClass: organizationalUnit"
      If ver = "NDS 8" Then Print #1, "objectClass: ndsLoginProperties"
      If ver = "NDS 8" Then Print #1, "objectClass: top"
      Print #1,
      
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU27 + ",ou=" + OrgU21 + ",ou=" + OrgU15 + ",ou=" + OrgU9 + ",ou=" + OrgU3 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU27 + ",ou=" + OrgU21 + ",ou=" + OrgU15 + ",ou=" + OrgU9 + ",ou=" + OrgU3 + ",o=" + Org
      Print #1, "changetype: " + CType
      If ver = "NDS 7" Then Print #1, "objectClass: top"
      If ver = "NDS 8" Then Print #1, "ou: " + OrgU27
      Print #1, "objectClass: organizationalUnit"
      If ver = "NDS 8" Then Print #1, "objectClass: ndsLoginProperties"
      If ver = "NDS 8" Then Print #1, "objectClass: top"
      Print #1,
      
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU28 + ",ou=" + OrgU22 + ",ou=" + OrgU16 + ",ou=" + OrgU10 + ",ou=" + OrgU4 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU28 + ",ou=" + OrgU22 + ",ou=" + OrgU16 + ",ou=" + OrgU10 + ",ou=" + OrgU4 + ",o=" + Org
      Print #1, "changetype: " + CType
      If ver = "NDS 7" Then Print #1, "objectClass: top"
      If ver = "NDS 8" Then Print #1, "ou: " + OrgU28
      Print #1, "objectClass: organizationalUnit"
      If ver = "NDS 8" Then Print #1, "objectClass: ndsLoginProperties"
      If ver = "NDS 8" Then Print #1, "objectClass: top"
      Print #1,
      
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU29 + ",ou=" + OrgU23 + ",ou=" + OrgU17 + ",ou=" + OrgU11 + ",ou=" + OrgU5 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU29 + ",ou=" + OrgU23 + ",ou=" + OrgU17 + ",ou=" + OrgU11 + ",ou=" + OrgU5 + ",o=" + Org
      Print #1, "changetype: " + CType
      If ver = "NDS 7" Then Print #1, "objectClass: top"
      If ver = "NDS 8" Then Print #1, "ou: " + OrgU29
      Print #1, "objectClass: organizationalUnit"
      If ver = "NDS 8" Then Print #1, "objectClass: ndsLoginProperties"
      If ver = "NDS 8" Then Print #1, "objectClass: top"
      Print #1,
      
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU30 + ",ou=" + OrgU24 + ",ou=" + OrgU18 + ",ou=" + OrgU12 + ",ou=" + OrgU6 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU30 + ",ou=" + OrgU24 + ",ou=" + OrgU18 + ",ou=" + OrgU12 + ",ou=" + OrgU6 + ",o=" + Org
      Print #1, "changetype: " + CType
      If ver = "NDS 7" Then Print #1, "objectClass: top"
      If ver = "NDS 8" Then Print #1, "ou: " + OrgU30
      Print #1, "objectClass: organizationalUnit"
      If ver = "NDS 8" Then Print #1, "objectClass: ndsLoginProperties"
      If ver = "NDS 8" Then Print #1, "objectClass: top"
      Print #1,
      
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU31 + ",ou=" + OrgU25 + ",ou=" + OrgU19 + ",ou=" + OrgU13 + ",ou=" + OrgU7 + ",ou=" + OrgU1 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU31 + ",ou=" + OrgU25 + ",ou=" + OrgU19 + ",ou=" + OrgU13 + ",ou=" + OrgU7 + ",ou=" + OrgU1 + ",o=" + Org
      Print #1, "changetype: " + CType
      If ver = "NDS 7" Then Print #1, "objectClass: top"
      If ver = "NDS 8" Then Print #1, "ou: " + OrgU31
      Print #1, "objectClass: organizationalUnit"
      If ver = "NDS 8" Then Print #1, "objectClass: ndsLoginProperties"
      If ver = "NDS 8" Then Print #1, "objectClass: top"
      Print #1,
      
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU32 + ",ou=" + OrgU26 + ",ou=" + OrgU20 + ",ou=" + OrgU14 + ",ou=" + OrgU8 + ",ou=" + OrgU2 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU32 + ",ou=" + OrgU26 + ",ou=" + OrgU20 + ",ou=" + OrgU14 + ",ou=" + OrgU8 + ",ou=" + OrgU2 + ",o=" + Org
      Print #1, "changetype: " + CType
      If ver = "NDS 7" Then Print #1, "objectClass: top"
      If ver = "NDS 8" Then Print #1, "ou: " + OrgU32
      Print #1, "objectClass: organizationalUnit"
      If ver = "NDS 8" Then Print #1, "objectClass: ndsLoginProperties"
      If ver = "NDS 8" Then Print #1, "objectClass: top"
      Print #1,
      
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU33 + ",ou=" + OrgU27 + ",ou=" + OrgU21 + ",ou=" + OrgU15 + ",ou=" + OrgU9 + ",ou=" + OrgU3 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU33 + ",ou=" + OrgU27 + ",ou=" + OrgU21 + ",ou=" + OrgU15 + ",ou=" + OrgU9 + ",ou=" + OrgU3 + ",o=" + Org
      Print #1, "changetype: " + CType
      If ver = "NDS 7" Then Print #1, "objectClass: top"
      If ver = "NDS 8" Then Print #1, "ou: " + OrgU33
      Print #1, "objectClass: organizationalUnit"
      If ver = "NDS 8" Then Print #1, "objectClass: ndsLoginProperties"
      If ver = "NDS 8" Then Print #1, "objectClass: top"
      Print #1,
      
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU34 + ",ou=" + OrgU28 + ",ou=" + OrgU22 + ",ou=" + OrgU16 + ",ou=" + OrgU10 + ",ou=" + OrgU4 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU34 + ",ou=" + OrgU28 + ",ou=" + OrgU22 + ",ou=" + OrgU16 + ",ou=" + OrgU10 + ",ou=" + OrgU4 + ",o=" + Org
      Print #1, "changetype: " + CType
      If ver = "NDS 7" Then Print #1, "objectClass: top"
      If ver = "NDS 8" Then Print #1, "ou: " + OrgU34
      Print #1, "objectClass: organizationalUnit"
      If ver = "NDS 8" Then Print #1, "objectClass: ndsLoginProperties"
      If ver = "NDS 8" Then Print #1, "objectClass: top"
      Print #1,
      
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU35 + ",ou=" + OrgU29 + ",ou=" + OrgU23 + ",ou=" + OrgU17 + ",ou=" + OrgU11 + ",ou=" + OrgU5 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU35 + ",ou=" + OrgU29 + ",ou=" + OrgU23 + ",ou=" + OrgU17 + ",ou=" + OrgU11 + ",ou=" + OrgU5 + ",o=" + Org
      Print #1, "changetype: " + CType
      If ver = "NDS 7" Then Print #1, "objectClass: top"
      If ver = "NDS 8" Then Print #1, "ou: " + OrgU35
      Print #1, "objectClass: organizationalUnit"
      If ver = "NDS 8" Then Print #1, "objectClass: ndsLoginProperties"
      If ver = "NDS 8" Then Print #1, "objectClass: top"
      Print #1,
      
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU36 + ",ou=" + OrgU30 + ",ou=" + OrgU24 + ",ou=" + OrgU18 + ",ou=" + OrgU12 + ",ou=" + OrgU6 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU36 + ",ou=" + OrgU30 + ",ou=" + OrgU24 + ",ou=" + OrgU18 + ",ou=" + OrgU12 + ",ou=" + OrgU6 + ",o=" + Org
      Print #1, "changetype: " + CType
      If ver = "NDS 7" Then Print #1, "objectClass: top"
      If ver = "NDS 8" Then Print #1, "ou: " + OrgU36
      Print #1, "objectClass: organizationalUnit"
      If ver = "NDS 8" Then Print #1, "objectClass: ndsLoginProperties"
      If ver = "NDS 8" Then Print #1, "objectClass: top"
      'Print #1,
      
'End Add Tree Information

BypassTreeAdd:
BypassTree:
BypassTree1:
   
   For i = beginval To endval
      ProgressBar1.Value = i
      
      'Start User data
      If UserSetOneCheck = 0 Then GoTo Step10 'If Create User(s) is not checked, bypass user creation for that set
      Print #1,
      If CustContCheck = 1 Then Print #1, "dn: cn=" + userid1 + Format(i) + CustCont
      If CustContCheck = 1 Then GoTo CCStep3
      If Not CCont = "" Then Print #1, "dn: cn=" + userid1 + Format(i) + ",ou=" + OrgUnit111 + ",ou=" + OrgUnit11 + ",ou=" + OrgUnit1 + ",o=" + Org1 + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: cn=" + userid1 + Format(i) + ",ou=" + OrgUnit111 + ",ou=" + OrgUnit11 + ",ou=" + OrgUnit1 + ",o=" + Org1
CCStep3:
      Print #1, "changetype: " + CType
      If CType = "delete" Then GoTo Step10
      If ver = "NDS 7" Then Print #1, "objectclass: top"
      If ver = "NDS 7" Then Print #1, "objectclass: person"
      If ver = "NDS 7" Then Print #1, "objectclass: organizationalPerson"
      If ver = "NDS 7" Then Print #1, "objectclass: inetOrgPerson"
      If ver = "NDS 7" Then Print #1, "mail: " + userid1 + Format(i) + "@" + domainname
      If ver = "NDS 7" Then Print #1, "givenName: " + givenname1 + Format(i)
      If ver = "NDS 7" Then Print #1, "sn: " + surname1 + Format(i)
      If ver = "NDS 8" Then Print #1, "mail: " + userid1 + Format(i) + "@" + domainname
      Print #1, "uid: " + userid1 + Format(i)
      If ver = "NDS 8" Then Print #1, "givenName: " + givenname1 + Format(i)
      If ver = "NDS 8" Then Print #1, "sn: " + surname1 + Format(i)
      If ver = "NDS 8" Then Print #1, "objectclass: inetOrgPerson"
      If ver = "NDS 8" Then Print #1, "objectclass: organizationalPerson"
      If ver = "NDS 8" Then Print #1, "objectclass: person"
      If ver = "NDS 8" Then Print #1, "objectclass: ndsLoginProperties"
      If ver = "NDS 8" Then Print #1, "objectclass: top"
      If passwordm = "Yes" Then Print #1, "userpassword: " + userid1 + Format(i) Else Print #1, "userpassword: " + password1
      If homedir = "cn=server1_Vol1,ou=OrganizationalUnit,o=Container#0#\Users_directory" Then GoTo Step10
      If ver = "NDS 8" Then Print #1, "ndsHomeDirectory: " + homedir + "\" + userid1 + Format(i)
      If ver = "NDS 7" Then Print #1, "homeDirectory: " + homedir + "\" + userid1 + Format(i)
           
Step10:

      If UserSetTwoCheck = 0 Then GoTo Step11 'If Create User(s) is not checked, bypass user creation for that set
      Print #1,
      If Not CCont = "" Then Print #1, "dn: cn=" + userid2 + Format(i) + ",ou=" + OrgUnit222 + ",ou=" + OrgUnit22 + ",ou=" + OrgUnit2 + ",o=" + Org2 + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: cn=" + userid2 + Format(i) + ",ou=" + OrgUnit222 + ",ou=" + OrgUnit22 + ",ou=" + OrgUnit2 + ",o=" + Org2
      Print #1, "changetype: " + CType
      If CType = "delete" Then GoTo Step11
      If ver = "NDS 7" Then Print #1, "objectclass: top"
      If ver = "NDS 7" Then Print #1, "objectclass: person"
      If ver = "NDS 7" Then Print #1, "objectclass: organizationalPerson"
      If ver = "NDS 7" Then Print #1, "objectclass: inetOrgPerson"
      If ver = "NDS 7" Then Print #1, "mail: " + userid2 + Format(i) + "@" + domainname
      If ver = "NDS 7" Then Print #1, "givenName: " + givenname2 + Format(i)
      If ver = "NDS 7" Then Print #1, "sn: " + surname2 + Format(i)
      If ver = "NDS 8" Then Print #1, "mail: " + userid2 + Format(i) + "@" + domainname
      Print #1, "uid: " + userid2 + Format(i)
      If ver = "NDS 8" Then Print #1, "givenName: " + givenname2 + Format(i)
      If ver = "NDS 8" Then Print #1, "sn: " + surname2 + Format(i)
      If ver = "NDS 8" Then Print #1, "objectclass: inetOrgPerson"
      If ver = "NDS 8" Then Print #1, "objectclass: organizationalPerson"
      If ver = "NDS 8" Then Print #1, "objectclass: person"
      If ver = "NDS 8" Then Print #1, "objectclass: ndsLoginProperties"
      If ver = "NDS 8" Then Print #1, "objectclass: top"
      If passwordm = "Yes" Then Print #1, "userpassword: " + userid1 + Format(i) Else Print #1, "userpassword: " + password2
      If homedir = "cn=server1_Vol1,ou=OrganizationalUnit,o=Container#0#\Users_directory" Then GoTo Step11
      If ver = "NDS 8" Then Print #1, "ndsHomeDirectory: " + homedir + "\" + userid2 + Format(i)
      If ver = "NDS 7" Then Print #1, "homeDirectory: " + homedir + "\" + userid2 + Format(i)
            
Step11:
      
      If UserSetThreeCheck = 0 Then GoTo Step12 'If Create User(s) is not checked, bypass user creation for that set
      Print #1,
      If Not CCont = "" Then Print #1, "dn: cn=" + userid3 + Format(i) + ",ou=" + OrgUnit333 + ",ou=" + OrgUnit33 + ",ou=" + OrgUnit3 + ",o=" + Org3 + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: cn=" + userid3 + Format(i) + ",ou=" + OrgUnit333 + ",ou=" + OrgUnit33 + ",ou=" + OrgUnit3 + ",o=" + Org3
      Print #1, "changetype: " + CType
      If CType = "delete" Then GoTo Step12
      If ver = "NDS 7" Then Print #1, "objectclass: top"
      If ver = "NDS 7" Then Print #1, "objectclass: person"
      If ver = "NDS 7" Then Print #1, "objectclass: organizationalPerson"
      If ver = "NDS 7" Then Print #1, "objectclass: inetOrgPerson"
      If ver = "NDS 7" Then Print #1, "mail: " + userid3 + Format(i) + "@" + domainname
      If ver = "NDS 7" Then Print #1, "givenName: " + givenname3 + Format(i)
      If ver = "NDS 7" Then Print #1, "sn: " + surname3 + Format(i)
      If ver = "NDS 8" Then Print #1, "mail: " + userid3 + Format(i) + "@" + domainname
      Print #1, "uid: " + userid3 + Format(i)
      If ver = "NDS 8" Then Print #1, "givenName: " + givenname3 + Format(i)
      If ver = "NDS 8" Then Print #1, "sn: " + surname3 + Format(i)
      If ver = "NDS 8" Then Print #1, "objectclass: inetOrgPerson"
      If ver = "NDS 8" Then Print #1, "objectclass: organizationalPerson"
      If ver = "NDS 8" Then Print #1, "objectclass: person"
      If ver = "NDS 8" Then Print #1, "objectclass: ndsLoginProperties"
      If ver = "NDS 8" Then Print #1, "objectclass: top"
      If passwordm = "Yes" Then Print #1, "userpassword: " + userid3 + Format(i) Else Print #1, "userpassword: " + password3
      If homedir = "cn=server1_Vol1,ou=OrganizationalUnit,o=Container#0#\Users_directory" Then GoTo Step12
      If ver = "NDS 8" Then Print #1, "ndsHomeDirectory: " + homedir + "\" + userid3 + Format(i)
      If ver = "NDS 7" Then Print #1, "homeDirectory: " + homedir + "\" + userid3 + Format(i)
            
Step12:
            
      If UserSetFourCheck = 0 Then GoTo Step13 'If Create User(s) is not checked, bypass user creation for that set
      Print #1,
      If Not CCont = "" Then Print #1, "dn: cn=" + userid4 + Format(i) + ",ou=" + OrgUnit444 + ",ou=" + OrgUnit44 + ",ou=" + OrgUnit4 + ",o=" + Org4 + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: cn=" + userid4 + Format(i) + ",ou=" + OrgUnit444 + ",ou=" + OrgUnit44 + ",ou=" + OrgUnit4 + ",o=" + Org4
      Print #1, "changetype: " + CType
      If CType = "delete" Then GoTo Step13
      If ver = "NDS 7" Then Print #1, "objectclass: top"
      If ver = "NDS 7" Then Print #1, "objectclass: person"
      If ver = "NDS 7" Then Print #1, "objectclass: organizationalPerson"
      If ver = "NDS 7" Then Print #1, "objectclass: inetOrgPerson"
      If ver = "NDS 7" Then Print #1, "mail: " + userid4 + Format(i) + "@" + domainname
      If ver = "NDS 7" Then Print #1, "givenName: " + givenname4 + Format(i)
      If ver = "NDS 7" Then Print #1, "sn: " + surname4 + Format(i)
      If ver = "NDS 8" Then Print #1, "mail: " + userid4 + Format(i) + "@" + domainname
      Print #1, "uid: " + userid1 + Format(i)
      If ver = "NDS 8" Then Print #1, "givenName: " + givenname4 + Format(i)
      If ver = "NDS 8" Then Print #1, "sn: " + surname4 + Format(i)
      If ver = "NDS 8" Then Print #1, "objectclass: inetOrgPerson"
      If ver = "NDS 8" Then Print #1, "objectclass: organizationalPerson"
      If ver = "NDS 8" Then Print #1, "objectclass: person"
      If ver = "NDS 8" Then Print #1, "objectclass: ndsLoginProperties"
      If ver = "NDS 8" Then Print #1, "objectclass: top"
      If passwordm = "Yes" Then Print #1, "userpassword: " + userid4 + Format(i) Else Print #1, "userpassword: " + password4
      If homedir = "cn=server1_Vol1,ou=OrganizationalUnit,o=Container#0#\Users_directory" Then GoTo Step13
      If ver = "NDS 8" Then Print #1, "ndsHomeDirectory: " + homedir + "\" + userid4 + Format(i)
      If ver = "NDS 7" Then Print #1, "homeDirectory: " + homedir + "\" + userid4 + Format(i)
      'End User Data
      
Step13:
      ProgressBar1.Visible = True
      
   Next i
      
'Start Delete Tree Information
If SkipTreeCreateCheck = "0" Then GoTo BypassTree2
Step41:

   If CheckAdd = 1 Then GoTo ByPassDel

      If CheckDel = 1 Then Print #1,
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU36 + ",ou=" + OrgU30 + ",ou=" + OrgU24 + ",ou=" + OrgU18 + ",ou=" + OrgU12 + ",ou=" + OrgU6 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU36 + ",ou=" + OrgU30 + ",ou=" + OrgU24 + ",ou=" + OrgU18 + ",ou=" + OrgU12 + ",ou=" + OrgU6 + ",o=" + Org
      Print #1, "changetype: delete"
      Print #1,
      
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU35 + ",ou=" + OrgU29 + ",ou=" + OrgU23 + ",ou=" + OrgU17 + ",ou=" + OrgU11 + ",ou=" + OrgU5 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU35 + ",ou=" + OrgU29 + ",ou=" + OrgU23 + ",ou=" + OrgU17 + ",ou=" + OrgU11 + ",ou=" + OrgU5 + ",o=" + Org
      Print #1, "changetype: delete"
      Print #1,
            
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU34 + ",ou=" + OrgU28 + ",ou=" + OrgU22 + ",ou=" + OrgU16 + ",ou=" + OrgU10 + ",ou=" + OrgU4 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU34 + ",ou=" + OrgU28 + ",ou=" + OrgU22 + ",ou=" + OrgU16 + ",ou=" + OrgU10 + ",ou=" + OrgU4 + ",o=" + Org
      Print #1, "changetype: delete"
      Print #1,
    
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU33 + ",ou=" + OrgU27 + ",ou=" + OrgU21 + ",ou=" + OrgU15 + ",ou=" + OrgU9 + ",ou=" + OrgU3 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU33 + ",ou=" + OrgU27 + ",ou=" + OrgU21 + ",ou=" + OrgU15 + ",ou=" + OrgU9 + ",ou=" + OrgU3 + ",o=" + Org
      Print #1, "changetype: delete"
      Print #1,
      
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU32 + ",ou=" + OrgU26 + ",ou=" + OrgU20 + ",ou=" + OrgU14 + ",ou=" + OrgU8 + ",ou=" + OrgU2 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU32 + ",ou=" + OrgU26 + ",ou=" + OrgU20 + ",ou=" + OrgU14 + ",ou=" + OrgU8 + ",ou=" + OrgU2 + ",o=" + Org
      Print #1, "changetype: delete"
      Print #1,
      
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU31 + ",ou=" + OrgU25 + ",ou=" + OrgU19 + ",ou=" + OrgU13 + ",ou=" + OrgU7 + ",ou=" + OrgU1 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU31 + ",ou=" + OrgU25 + ",ou=" + OrgU19 + ",ou=" + OrgU13 + ",ou=" + OrgU7 + ",ou=" + OrgU1 + ",o=" + Org
      Print #1, "changetype: delete"
      Print #1,
      
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU30 + ",ou=" + OrgU24 + ",ou=" + OrgU18 + ",ou=" + OrgU12 + ",ou=" + OrgU6 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU30 + ",ou=" + OrgU24 + ",ou=" + OrgU18 + ",ou=" + OrgU12 + ",ou=" + OrgU6 + ",o=" + Org
      Print #1, "changetype: delete"
      Print #1,
      
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU29 + ",ou=" + OrgU23 + ",ou=" + OrgU17 + ",ou=" + OrgU11 + ",ou=" + OrgU5 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU29 + ",ou=" + OrgU23 + ",ou=" + OrgU17 + ",ou=" + OrgU11 + ",ou=" + OrgU5 + ",o=" + Org
      Print #1, "changetype: delete"
      Print #1,
      
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU28 + ",ou=" + OrgU22 + ",ou=" + OrgU16 + ",ou=" + OrgU10 + ",ou=" + OrgU4 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU28 + ",ou=" + OrgU22 + ",ou=" + OrgU16 + ",ou=" + OrgU10 + ",ou=" + OrgU4 + ",o=" + Org
      Print #1, "changetype: delete"
      Print #1,
      
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU27 + ",ou=" + OrgU21 + ",ou=" + OrgU15 + ",ou=" + OrgU9 + ",ou=" + OrgU3 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU27 + ",ou=" + OrgU21 + ",ou=" + OrgU15 + ",ou=" + OrgU9 + ",ou=" + OrgU3 + ",o=" + Org
      Print #1, "changetype: delete"
      Print #1,
      
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU26 + ",ou=" + OrgU20 + ",ou=" + OrgU14 + ",ou=" + OrgU8 + ",ou=" + OrgU2 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU26 + ",ou=" + OrgU20 + ",ou=" + OrgU14 + ",ou=" + OrgU8 + ",ou=" + OrgU2 + ",o=" + Org
      Print #1, "changetype: delete"
      Print #1,
      
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU25 + ",ou=" + OrgU19 + ",ou=" + OrgU13 + ",ou=" + OrgU7 + ",ou=" + OrgU1 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU25 + ",ou=" + OrgU19 + ",ou=" + OrgU13 + ",ou=" + OrgU7 + ",ou=" + OrgU1 + ",o=" + Org
      Print #1, "changetype: delete"
      Print #1,
      
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU24 + ",ou=" + OrgU18 + ",ou=" + OrgU12 + ",ou=" + OrgU6 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU24 + ",ou=" + OrgU18 + ",ou=" + OrgU12 + ",ou=" + OrgU6 + ",o=" + Org
      Print #1, "changetype: delete"
      Print #1,
      
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU23 + ",ou=" + OrgU17 + ",ou=" + OrgU11 + ",ou=" + OrgU5 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU23 + ",ou=" + OrgU17 + ",ou=" + OrgU11 + ",ou=" + OrgU5 + ",o=" + Org
      Print #1, "changetype: delete"
      Print #1,
      
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU22 + ",ou=" + OrgU16 + ",ou=" + OrgU10 + ",ou=" + OrgU4 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU22 + ",ou=" + OrgU16 + ",ou=" + OrgU10 + ",ou=" + OrgU4 + ",o=" + Org
      Print #1, "changetype: delete"
      Print #1,
      
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU21 + ",ou=" + OrgU15 + ",ou=" + OrgU9 + ",ou=" + OrgU3 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU21 + ",ou=" + OrgU15 + ",ou=" + OrgU9 + ",ou=" + OrgU3 + ",o=" + Org
      Print #1, "changetype: delete"
      Print #1,
      
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU20 + ",ou=" + OrgU14 + ",ou=" + OrgU8 + ",ou=" + OrgU2 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU20 + ",ou=" + OrgU14 + ",ou=" + OrgU8 + ",ou=" + OrgU2 + ",o=" + Org
      Print #1, "changetype: delete"
      Print #1,
      
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU19 + ",ou=" + OrgU13 + ",ou=" + OrgU7 + ",ou=" + OrgU1 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU19 + ",ou=" + OrgU13 + ",ou=" + OrgU7 + ",ou=" + OrgU1 + ",o=" + Org
      Print #1, "changetype: delete"
      Print #1,
      
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU18 + ",ou=" + OrgU12 + ",ou=" + OrgU6 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU18 + ",ou=" + OrgU12 + ",ou=" + OrgU6 + ",o=" + Org
      Print #1, "changetype: delete"
      Print #1,
      
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU17 + ",ou=" + OrgU11 + ",ou=" + OrgU5 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU17 + ",ou=" + OrgU11 + ",ou=" + OrgU5 + ",o=" + Org
      Print #1, "changetype: delete"
      Print #1,
      
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU16 + ",ou=" + OrgU10 + ",ou=" + OrgU4 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU16 + ",ou=" + OrgU10 + ",ou=" + OrgU4 + ",o=" + Org
      Print #1, "changetype: delete"
      Print #1,
      
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU15 + ",ou=" + OrgU9 + ",ou=" + OrgU3 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU15 + ",ou=" + OrgU9 + ",ou=" + OrgU3 + ",o=" + Org
      Print #1, "changetype: delete"
      Print #1,
      
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU14 + ",ou=" + OrgU8 + ",ou=" + OrgU2 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU14 + ",ou=" + OrgU8 + ",ou=" + OrgU2 + ",o=" + Org
      Print #1, "changetype: delete"
      Print #1,
      
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU13 + ",ou=" + OrgU7 + ",ou=" + OrgU1 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU13 + ",ou=" + OrgU7 + ",ou=" + OrgU1 + ",o=" + Org
      Print #1, "changetype: delete"
      Print #1,
      
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU12 + ",ou=" + OrgU6 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU12 + ",ou=" + OrgU6 + ",o=" + Org
      Print #1, "changetype: delete"
      Print #1,
      
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU11 + ",ou=" + OrgU5 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU11 + ",ou=" + OrgU5 + ",o=" + Org
      Print #1, "changetype: delete"
      Print #1,
      
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU10 + ",ou=" + OrgU4 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU10 + ",ou=" + OrgU4 + ",o=" + Org
      Print #1, "changetype: delete"
      Print #1,
      
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU9 + ",ou=" + OrgU3 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU9 + ",ou=" + OrgU3 + ",o=" + Org
      Print #1, "changetype: delete"
      Print #1,
      
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU8 + ",ou=" + OrgU2 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU8 + ",ou=" + OrgU2 + ",o=" + Org
      Print #1, "changetype: delete"
      Print #1,
      
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU7 + ",ou=" + OrgU1 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU7 + ",ou=" + OrgU1 + ",o=" + Org
      Print #1, "changetype: delete"
      Print #1,
      
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU6 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU6 + ",o=" + Org
      Print #1, "changetype: delete"
      Print #1,
      
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU5 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU5 + ",o=" + Org
      Print #1, "changetype: delete"
      Print #1,
      
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU4 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU4 + ",o=" + Org
      Print #1, "changetype: delete"
      Print #1,
      
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU3 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU3 + ",o=" + Org
      Print #1, "changetype: delete"
      Print #1,
      
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU2 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU2 + ",o=" + Org
      Print #1, "changetype: delete"
      Print #1,
      
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU1 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU1 + ",o=" + Org
      Print #1, "changetype: delete"
      Print #1,
      
      If Not CCont = "" Then Print #1, "dn: o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: o=" + Org
      Print #1, "changetype: delete"
      Print #1,
            
If CCont = "" Then GoTo ByPassDel

      Print #1, "dn: c=" + CCont
      Print #1, "changetype: delete"
      Print #1,

ByPassDel:
BypassTree2:

'End Delete Tree Information
   
   Close #1
   
Step2:
CustLDIFStep1:

'Start ice Batch
Dim Sice As String
Dim Source_Sw As String
Dim Destination_Sw As String
Dim LDIF_Sw As String
Dim LDAP_Sw As String
Dim IP_Sw As String
Dim IP As String
Dim Prt_Sw As String
Dim Prt As String
Dim Uname_Sw As String
Dim Uname As String
Dim Pwd_Sw As String
Dim Pwd As String
Dim BDN_Sw As String
Dim BDN As String
Dim Srch_Sw As String
Dim Srch As String
Dim RtCrt_Sw As String
Dim RtCrt As String
Dim StopOnErr As String
Dim F_Sw As String
Dim FOut As String
Dim FIn As String
Dim CFile As String
Dim Sp As String

Sice = "ice" + " "
Source_Sw = "-S" + " "
Destination_Sw = "-D" + " "
LDIF_Sw = "LDIF" + " "
LDAP_Sw = "LDAP" + " "
IP_Sw = "-s" + " "
IP = Text1.Text + " "
Prt_Sw = "-p" + " "
Prt = Combo1.Text + " "
Uname_Sw = "-d" + " "
Uname = Text2.Text + " "
Pwd_Sw = "-w" + " "
Pwd = Text3.Text + " "
BDN_Sw = "-b" + " "
BDN = Text14.Text + " "
Srch_Sw = "-c" + " "
Srch = Text31.Text + " "
RtCrt_Sw = "-L" + " "
RtCrt = Text4.Text + " "
F_Sw = "-f" + " "
FOut = Text6.Text + " "
FIn = Text7.Text + " "
StopOnErr = "-c" + " "
CFile = Text30.Text + " "
Sp = " "

If IP = "" + " " Then   ' IP = ""
      MyVar = MsgBox("Please specify an IPAddress.", 0, "Specify IP Address") ' Perform some action.
      If MyVar = vbOK Then GoTo Step161
End If

If Prt = "" + " " Then   ' Prt = ""
      MyVar = MsgBox("Please specify a Port.", 0, "Specify Port") ' Perform some action.
      If MyVar = vbOK Then GoTo Step161
End If

If Uname = " " Then Uname_Sw = ""
If Uname = " " Then Uname = ""
If Pwd = " " Then Pwd_Sw = ""
If Pwd = " " Then Pwd = ""
If BDN = " " Then BDN_Sw = ""
If BDN = " " Then BDN = ""
If Srch = " " Then Srch_Sw = ""
If Srch = " " Then Srch = ""
If RtCrt = " " Then RtCrt_Sw = ""
If RtCrt = " " Then RtCrt = ""
If RtCrt = " " Then RtCrt_Sw = ""
If RtCrt = " " Then RtCrt = ""
If CustomCheck = 1 Then FOut = CFile

      Open RICEBatch For Output As #2
      Print #2, "path %PATH%;" + App.Path
      Print #2, "del ice.Log"
      If RetrieveTreeCheck = 1 Then GoTo Chk2Rtv:

'Send LDAP
If StopOnErrCheck = 0 Then Print #2, Sice + Source_Sw + LDIF_Sw + F_Sw + FOut + StopOnErr + Destination_Sw + LDAP_Sw + IP_Sw + IP + Prt_Sw + Prt + Uname_Sw + Uname + Pwd_Sw + Pwd + RtCrt_Sw + RtCrt
If StopOnErrCheck = 1 Then Print #2, Sice + Source_Sw + LDIF_Sw + F_Sw + FOut + Destination_Sw + LDAP_Sw + IP_Sw + IP + Prt_Sw + Prt + Uname_Sw + Uname + Pwd_Sw + Pwd + RtCrt_Sw + RtCrt

Chk2Rtv:
'Get LDAP
If RetrieveTreeCheck = 1 Then Print #2, Sice + Source_Sw + LDAP_Sw + IP_Sw + IP + Prt_Sw + Prt + Uname_Sw + Uname + Pwd_Sw + Pwd + BDN_Sw + BDN + Srch_Sw + Srch + RtCrt_Sw + RtCrt + Destination_Sw + LDIF_Sw + F_Sw + FIn

   Close #2
'End ice Batch

'Start Display Page
 ICE.Show 0
 
   Dim JobToDo As String
      JobToDo = RICEBatch
      DoEvents: Sleep 100
      Shell32Bit JobToDo
'End Display Page

'Start User Home Directory Creation Section
   Dim beginval1 As Long
   Dim endval1 As Long
   
   beginval1 = 1
   endval1 = Val(Combo2.Text)
   
If UserHomeCheck = 0 Then GoTo Step161
If homedir = "cn=server1_Vol1,ou=OrganizationalUnit,o=Container#0#\Users_directory" Then GoTo Step161
          If (fso.FolderExists(MyPath)) Then GoTo Step1000 Else MkDir MyPath 'Make new directory or folder.
Step1000:
          ChDrive MyPath
          ChDir MyPath 'Changes the current directory or folder.
If endval = "1" Then GoTo Step150
      ProgressBar1.Min = beginval1
      ProgressBar1.Max = endval1
      ProgressBar1.Value = ProgressBar1.Min
      ProgressBar1.Visible = False

Step150:
      For j = beginval1 To endval1
         ProgressBar1.Value = j
   If userid1 = "" Then GoTo Step151
   If UserSetOneCheck = 0 Then GoTo Step151
      If (fso.FolderExists(MyPath + userid1 + Format(j))) Then GoTo CStep1 Else MkDir MyPath + userid1 + Format(j)  'Make new directory or folder.
CStep1:
      ChDir MyPath + userid1 + Format(j)
      If (fso.FolderExists(UHomeDir)) Then GoTo CStep2 Else MkDir UHomeDir  'Make new directory or folder.
CStep2:
      ChDir UHomeDir
      Open UHomeFile For Output As #2
      Print #2, userid1 + Format(j)
      ChDir "cd ..\.."
      Close #2
            
Step151:
   If userid2 = "" Then GoTo Step152
   If UserSetTwoCheck = 0 Then GoTo Step152
      If (fso.FolderExists(MyPath + userid2 + Format(j))) Then GoTo CStep3 Else MkDir MyPath + userid2 + Format(j) 'Make new directory or folder.
CStep3:
      ChDir MyPath + userid2 + Format(j)
      If (fso.FolderExists(UHomeDir)) Then GoTo CStep4 Else MkDir UHomeDir  'Make new directory or folder.
CStep4:
      ChDir UHomeDir
      Open UHomeFile For Output As #2
      Print #2, userid2 + Format(j)
      ChDir "cd ..\.."
      Close #2
         
Step152:
   If userid3 = "" Then GoTo Step153
   If UserSetThreeCheck = 0 Then GoTo Step153
      If (fso.FolderExists(MyPath + userid3 + Format(j))) Then GoTo CStep5 Else MkDir MyPath + userid3 + Format(j)  'Make new directory or folder.
CStep5:
      ChDir MyPath + userid3 + Format(j)
      If (fso.FolderExists(UHomeDir)) Then GoTo CStep6 Else MkDir UHomeDir  'Make new directory or folder.
CStep6:
      ChDir UHomeDir
      Open UHomeFile For Output As #2
      Print #2, userid3 + Format(j)
      ChDir "cd ..\.."
      Close #2
      
Step153:
   If userid4 = "" Then GoTo Step154
   If UserSetFourCheck = 0 Then GoTo Step154
      If (fso.FolderExists(MyPath + userid4 + Format(j))) Then GoTo CStep7 Else MkDir MyPath + userid4 + Format(j)  'Make new directory or folder.
CStep7:
      ChDir MyPath + userid4 + Format(j)
      If (fso.FolderExists(UHomeDir)) Then GoTo CStep8 Else MkDir UHomeDir  'Make new directory or folder.
CStep8:
      ChDir UHomeDir
      Open UHomeFile For Output As #2
      Print #2, userid4 + Format(j)
      ChDir "cd ..\.."
      Close #2
      
Step154:
      
If endval = "1" Then GoTo Step160
ProgressBar1.Visible = True
   Next j
      ChDir WPath
               
Step160:

     ChDrive "C:"
     ChDir App.Path

Step161:

'End User Home Directory Creation Section
ProgressBar1.Visible = True
ProgressBar1.Value = ProgressBar1.Max

On Error GoTo ErrorResume   'Resume next in case file is locked

'This is the start of the error handler
ErrorResume:
'in case file is locked
Resume Next
ErrorHandler:
If Err.Description = "" Then Resume Next Else MsgBox Err.Description
'This is the end of the error handler
End Sub

Private Sub CmdWrite_Click()
   
   Dim RICEBatch As String
   Dim fileout As String
   Dim filein As String
   Dim beginval As Long
   Dim endval As Long
   Dim ver As String
   Dim servername As String
   Dim domainname As String
   Dim passwordm As String
   Dim GetLDAP As String
   Dim BaseDN As String
   Dim homedir As String
   Dim userid1 As String
   Dim userid2 As String
   Dim userid3 As String
   Dim userid4 As String
   Dim givenname1 As String
   Dim givenname2 As String
   Dim givenname3 As String
   Dim givenname4 As String
   Dim surname1 As String
   Dim surname2 As String
   Dim surname3 As String
   Dim surname4 As String
   Dim password1 As String
   Dim password2 As String
   Dim password3 As String
   Dim password4 As String
   Dim CType As String
   Dim Org1 As String
   Dim Org2 As String
   Dim Org3 As String
   Dim Org4 As String
   Dim OrgUnit1 As String
   Dim OrgUnit2 As String
   Dim OrgUnit3 As String
   Dim OrgUnit4 As String
   Dim OrgUnit11 As String
   Dim OrgUnit22 As String
   Dim OrgUnit33 As String
   Dim OrgUnit44 As String
   Dim OrgUnit111 As String
   Dim OrgUnit222 As String
   Dim OrgUnit333 As String
   Dim OrgUnit444 As String
   Dim CCont As String
   Dim Org As String
   Dim OrgU As String
   Dim OrgU1 As String
   Dim OrgU2 As String
   Dim OrgU3 As String
   Dim OrgU4 As String
   Dim OrgU5 As String
   Dim OrgU6 As String
   Dim OrgU7 As String
   Dim OrgU8 As String
   Dim OrgU9 As String
   Dim OrgU10 As String
   Dim OrgU11 As String
   Dim OrgU12 As String
   Dim OrgU13 As String
   Dim OrgU14 As String
   Dim OrgU15 As String
   Dim OrgU16 As String
   Dim OrgU17 As String
   Dim OrgU18 As String
   Dim OrgU19 As String
   Dim OrgU20 As String
   Dim OrgU21 As String
   Dim OrgU22 As String
   Dim OrgU23 As String
   Dim OrgU24 As String
   Dim OrgU25 As String
   Dim OrgU26 As String
   Dim OrgU27 As String
   Dim OrgU28 As String
   Dim OrgU29 As String
   Dim OrgU30 As String
   Dim OrgU31 As String
   Dim OrgU32 As String
   Dim OrgU33 As String
   Dim OrgU34 As String
   Dim OrgU35 As String
   Dim OrgU36 As String
   Dim RetVal As Long
   Dim UHomeDir As String
   Dim UHomeFile As String
   Dim CustCont As String
   Dim CustLDIF As String
   Dim MyPath
   Dim WPath
   Dim hConsole As Long
      
   RICEBatch = Text5.Text
   fileout = Text6.Text
   filein = Text7.Text
   servername = Text8.Text
   domainname = Text9.Text
   beginval = 1
   endval = Val(Combo2.Text)
   Port = Combo1.Text
   passwordm = Combo14.Text
   ver = Combo6.Text
   WPath = Text19.Text
   homedir = Text401.Text
   MyPath = Text403.Text
   UHomeDir = Text404.Text
   UHomeFile = Text405.Text
   userid1 = Text10.Text
   userid2 = Text15.Text
   userid3 = Text20.Text
   userid4 = Text25.Text
   givenname1 = Text11.Text
   givenname2 = Text16.Text
   givenname3 = Text21.Text
   givenname4 = Text26.Text
   surname1 = Text12.Text
   surname2 = Text17.Text
   surname3 = Text22.Text
   surname4 = Text27.Text
   password1 = Text13.Text
   password2 = Text18.Text
   password3 = Text23.Text
   password4 = Text28.Text
   Org1 = Text103.Text
   Org2 = Text107.Text
   Org3 = Text111.Text
   Org4 = Text115.Text
   OrgUnit1 = Text102.Text
   OrgUnit2 = Text106.Text
   OrgUnit3 = Text110.Text
   OrgUnit4 = Text114.Text
   OrgUnit11 = Text101.Text
   OrgUnit22 = Text105.Text
   OrgUnit33 = Text109.Text
   OrgUnit44 = Text113.Text
   OrgUnit111 = Text100.Text
   OrgUnit222 = Text104.Text
   OrgUnit333 = Text108.Text
   OrgUnit444 = Text112.Text
   CCont = Combo199.Text
   Org = Text200.Text
   OrgU1 = Text1101.Text
   OrgU2 = Text1201.Text
   OrgU3 = Text1301.Text
   OrgU4 = Text1401.Text
   OrgU5 = Text1501.Text
   OrgU6 = Text1601.Text
   OrgU7 = Text1102.Text
   OrgU8 = Text1202.Text
   OrgU9 = Text1302.Text
   OrgU10 = Text1402.Text
   OrgU11 = Text1502.Text
   OrgU12 = Text1602.Text
   OrgU13 = Text1103.Text
   OrgU14 = Text1203.Text
   OrgU15 = Text1303.Text
   OrgU16 = Text1403.Text
   OrgU17 = Text1503.Text
   OrgU18 = Text1603.Text
   OrgU19 = Combo104.Text
   OrgU20 = Combo204.Text
   OrgU21 = Combo304.Text
   OrgU22 = Combo404.Text
   OrgU23 = Combo504.Text
   OrgU24 = Combo604.Text
   OrgU25 = Combo105.Text
   OrgU26 = Combo205.Text
   OrgU27 = Combo305.Text
   OrgU28 = Combo405.Text
   OrgU29 = Combo505.Text
   OrgU30 = Combo605.Text
   OrgU31 = Combo106.Text
   OrgU32 = Combo206.Text
   OrgU33 = Combo306.Text
   OrgU34 = Combo406.Text
   OrgU35 = Combo506.Text
   OrgU36 = Combo606.Text
   CustCont = Text29.Text
   CustLDIF = Text30.Text
   
   If CheckAdd = 1 Then CType = "add"
   If CheckDel = 1 Then CType = "delete"
                       
           Dim fso, Msg
           Set fso = CreateObject("Scripting.FileSystemObject")
           If (fso.FolderExists(WPath)) Then GoTo CStep180 Else MkDir WPath  'Make new directory or folder.
CStep180:
          ChDrive WPath 'Changes the current drive.
          ChDir WPath 'Changes the current directory or folder.
                              
        If RetrieveTreeCheck = 1 Then GoTo Step2

If CustomCheck = 1 Then GoTo CustLDIFStep1 'If Custom LDIF file is selected bypass the file creation.
                 
   Open fileout For Output As #1 '***File Open***
   
      Print #1, "#This file generated by Novell's Tree Builder Version 3. (Written by: Robert Foster)" 'File Header
      Print #1, "version: 1"
      If CheckDel = 0 Then Print #1,
   
If SkipTreeCreateCheck = 0 Then GoTo BypassTree

'***Need to add code to allow numbers to start at something other than zero at some point***

   If endval = "1" Then GoTo Step4 'If endval is 1 this goes around the Progress Bar Error
   ProgressBar1.Min = beginval
   ProgressBar1.Max = endval
   ProgressBar1.Value = ProgressBar1.Min
   ProgressBar1.Visible = False
         
Step4:
   
If CType = "Delete" Then GoTo Step41

'Start Add Tree Information
      If CCont = "" Then GoTo Step22
      If CType = "delete" Then GoTo Step22

      Print #1, "dn: c=" + CCont
      Print #1, "changetype: " + CType
      Print #1, "objectClass: top"
      Print #1, "objectClass: country"
      If ver = "NDS 8" Then Print #1, "c: " + CCont
      If ver = "NDS 8" Then Print #1, "objectClass: top"
      Print #1,

Step22:

If CheckDel = 1 Then GoTo BypassTreeAdd

      If Not CCont = "" Then Print #1, "dn: o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: o=" + Org
      Print #1, "changetype: " + CType
      If ver = "NDS 7" Then Print #1, "objectClass: top"
      If ver = "NDS 8" Then Print #1, "o: " + Org
      Print #1, "objectClass: organization"
      If ver = "NDS 8" Then Print #1, "objectClass: ndsLoginProperties"
      If ver = "NDS 8" Then Print #1, "objectClass: top"
      Print #1,
      
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU1 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU1 + ",o=" + Org
      Print #1, "changetype: " + CType
      If ver = "NDS 7" Then Print #1, "objectClass: top"
      If ver = "NDS 8" Then Print #1, "ou: " + OrgU1
      Print #1, "objectClass: organizationalUnit"
      If ver = "NDS 8" Then Print #1, "objectClass: ndsLoginProperties"
      If ver = "NDS 8" Then Print #1, "objectClass: top"
      Print #1,
      
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU2 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU2 + ",o=" + Org
      Print #1, "changetype: " + CType
      If ver = "NDS 7" Then Print #1, "objectClass: top"
      If ver = "NDS 8" Then Print #1, "ou: " + OrgU2
      Print #1, "objectClass: organizationalUnit"
      If ver = "NDS 8" Then Print #1, "objectClass: ndsLoginProperties"
      If ver = "NDS 8" Then Print #1, "objectClass: top"
      Print #1,
      
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU3 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU3 + ",o=" + Org
      Print #1, "changetype: " + CType
      If ver = "NDS 7" Then Print #1, "objectClass: top"
      If ver = "NDS 8" Then Print #1, "ou: " + OrgU3
      Print #1, "objectClass: organizationalUnit"
      If ver = "NDS 8" Then Print #1, "objectClass: ndsLoginProperties"
      If ver = "NDS 8" Then Print #1, "objectClass: top"
      Print #1,
      
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU4 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU4 + ",o=" + Org
      Print #1, "changetype: " + CType
      If ver = "NDS 7" Then Print #1, "objectClass: top"
      If ver = "NDS 8" Then Print #1, "ou: " + OrgU4
      Print #1, "objectClass: organizationalUnit"
      If ver = "NDS 8" Then Print #1, "objectClass: ndsLoginProperties"
      If ver = "NDS 8" Then Print #1, "objectClass: top"
      Print #1,
   
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU5 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU5 + ",o=" + Org
      Print #1, "changetype: " + CType
      If ver = "NDS 7" Then Print #1, "objectClass: top"
      If ver = "NDS 8" Then Print #1, "ou: " + OrgU5
      Print #1, "objectClass: organizationalUnit"
      If ver = "NDS 8" Then Print #1, "objectClass: ndsLoginProperties"
      If ver = "NDS 8" Then Print #1, "objectClass: top"
      Print #1,
   
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU6 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU6 + ",o=" + Org
      Print #1, "changetype: " + CType
      If ver = "NDS 7" Then Print #1, "objectClass: top"
      If ver = "NDS 8" Then Print #1, "ou: " + OrgU6
      Print #1, "objectClass: organizationalUnit"
      If ver = "NDS 8" Then Print #1, "objectClass: ndsLoginProperties"
      If ver = "NDS 8" Then Print #1, "objectClass: top"
      Print #1,
            
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU7 + ",ou=" + OrgU1 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU7 + ",ou=" + OrgU1 + ",o=" + Org
      Print #1, "changetype: " + CType
      If ver = "NDS 7" Then Print #1, "objectClass: top"
      If ver = "NDS 8" Then Print #1, "ou: " + OrgU7
      Print #1, "objectClass: organizationalUnit"
      If ver = "NDS 8" Then Print #1, "objectClass: ndsLoginProperties"
      If ver = "NDS 8" Then Print #1, "objectClass: top"
      Print #1,
      
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU8 + ",ou=" + OrgU2 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU8 + ",ou=" + OrgU2 + ",o=" + Org
      Print #1, "changetype: " + CType
      If ver = "NDS 7" Then Print #1, "objectClass: top"
      If ver = "NDS 8" Then Print #1, "ou: " + OrgU8
      Print #1, "objectClass: organizationalUnit"
      If ver = "NDS 8" Then Print #1, "objectClass: ndsLoginProperties"
      If ver = "NDS 8" Then Print #1, "objectClass: top"
      Print #1,
      
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU9 + ",ou=" + OrgU3 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU9 + ",ou=" + OrgU3 + ",o=" + Org
      Print #1, "changetype: " + CType
      If ver = "NDS 7" Then Print #1, "objectClass: top"
      If ver = "NDS 8" Then Print #1, "ou: " + OrgU9
      Print #1, "objectClass: organizationalUnit"
      If ver = "NDS 8" Then Print #1, "objectClass: ndsLoginProperties"
      If ver = "NDS 8" Then Print #1, "objectClass: top"
      Print #1,
      
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU10 + ",ou=" + OrgU4 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU10 + ",ou=" + OrgU4 + ",o=" + Org
      Print #1, "changetype: " + CType
      If ver = "NDS 7" Then Print #1, "objectClass: top"
      If ver = "NDS 8" Then Print #1, "ou: " + OrgU10
      Print #1, "objectClass: organizationalUnit"
      If ver = "NDS 8" Then Print #1, "objectClass: ndsLoginProperties"
      If ver = "NDS 8" Then Print #1, "objectClass: top"
      Print #1,
            
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU11 + ",ou=" + OrgU5 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU11 + ",ou=" + OrgU5 + ",o=" + Org
      Print #1, "changetype: " + CType
      If ver = "NDS 7" Then Print #1, "objectClass: top"
      If ver = "NDS 8" Then Print #1, "ou: " + OrgU11
      Print #1, "objectClass: organizationalUnit"
      If ver = "NDS 8" Then Print #1, "objectClass: ndsLoginProperties"
      If ver = "NDS 8" Then Print #1, "objectClass: top"
      Print #1,
            
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU12 + ",ou=" + OrgU6 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU12 + ",ou=" + OrgU6 + ",o=" + Org
      Print #1, "changetype: " + CType
      If ver = "NDS 7" Then Print #1, "objectClass: top"
      If ver = "NDS 8" Then Print #1, "ou: " + OrgU12
      Print #1, "objectClass: organizationalUnit"
      If ver = "NDS 8" Then Print #1, "objectClass: ndsLoginProperties"
      If ver = "NDS 8" Then Print #1, "objectClass: top"
      Print #1,
      
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU13 + ",ou=" + OrgU7 + ",ou=" + OrgU1 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU13 + ",ou=" + OrgU7 + ",ou=" + OrgU1 + ",o=" + Org
      Print #1, "changetype: " + CType
      If ver = "NDS 7" Then Print #1, "objectClass: top"
      If ver = "NDS 8" Then Print #1, "ou: " + OrgU13
      Print #1, "objectClass: organizationalUnit"
      If ver = "NDS 8" Then Print #1, "objectClass: ndsLoginProperties"
      If ver = "NDS 8" Then Print #1, "objectClass: top"
      Print #1,
      
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU14 + ",ou=" + OrgU8 + ",ou=" + OrgU2 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU14 + ",ou=" + OrgU8 + ",ou=" + OrgU2 + ",o=" + Org
      Print #1, "changetype: " + CType
      If ver = "NDS 7" Then Print #1, "objectClass: top"
      If ver = "NDS 8" Then Print #1, "ou: " + OrgU14
      Print #1, "objectClass: organizationalUnit"
      If ver = "NDS 8" Then Print #1, "objectClass: ndsLoginProperties"
      If ver = "NDS 8" Then Print #1, "objectClass: top"
      Print #1,
      
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU15 + ",ou=" + OrgU9 + ",ou=" + OrgU3 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU15 + ",ou=" + OrgU9 + ",ou=" + OrgU3 + ",o=" + Org
      Print #1, "changetype: " + CType
      If ver = "NDS 7" Then Print #1, "objectClass: top"
      If ver = "NDS 8" Then Print #1, "ou: " + OrgU15
      Print #1, "objectClass: organizationalUnit"
      If ver = "NDS 8" Then Print #1, "objectClass: ndsLoginProperties"
      If ver = "NDS 8" Then Print #1, "objectClass: top"
      Print #1,
      
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU16 + ",ou=" + OrgU10 + ",ou=" + OrgU4 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU16 + ",ou=" + OrgU10 + ",ou=" + OrgU4 + ",o=" + Org
      Print #1, "changetype: " + CType
      If ver = "NDS 7" Then Print #1, "objectClass: top"
      If ver = "NDS 8" Then Print #1, "ou: " + OrgU16
      Print #1, "objectClass: organizationalUnit"
      If ver = "NDS 8" Then Print #1, "objectClass: ndsLoginProperties"
      If ver = "NDS 8" Then Print #1, "objectClass: top"
      Print #1,
      
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU17 + ",ou=" + OrgU11 + ",ou=" + OrgU5 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU17 + ",ou=" + OrgU11 + ",ou=" + OrgU5 + ",o=" + Org
      Print #1, "changetype: " + CType
      If ver = "NDS 7" Then Print #1, "objectClass: top"
      If ver = "NDS 8" Then Print #1, "ou: " + OrgU17
      Print #1, "objectClass: organizationalUnit"
      If ver = "NDS 8" Then Print #1, "objectClass: ndsLoginProperties"
      If ver = "NDS 8" Then Print #1, "objectClass: top"
      Print #1,
      
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU18 + ",ou=" + OrgU12 + ",ou=" + OrgU6 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU18 + ",ou=" + OrgU12 + ",ou=" + OrgU6 + ",o=" + Org
      Print #1, "changetype: " + CType
      If ver = "NDS 7" Then Print #1, "objectClass: top"
      If ver = "NDS 8" Then Print #1, "ou: " + OrgU18
      Print #1, "objectClass: organizationalUnit"
      If ver = "NDS 8" Then Print #1, "objectClass: ndsLoginProperties"
      If ver = "NDS 8" Then Print #1, "objectClass: top"
      Print #1,
      
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU19 + ",ou=" + OrgU13 + ",ou=" + OrgU7 + ",ou=" + OrgU1 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU19 + ",ou=" + OrgU13 + ",ou=" + OrgU7 + ",ou=" + OrgU1 + ",o=" + Org
      Print #1, "changetype: " + CType
      If ver = "NDS 7" Then Print #1, "objectClass: top"
      If ver = "NDS 8" Then Print #1, "ou: " + OrgU19
      Print #1, "objectClass: organizationalUnit"
      If ver = "NDS 8" Then Print #1, "objectClass: ndsLoginProperties"
      If ver = "NDS 8" Then Print #1, "objectClass: top"
      Print #1,
      
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU20 + ",ou=" + OrgU14 + ",ou=" + OrgU8 + ",ou=" + OrgU2 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU20 + ",ou=" + OrgU14 + ",ou=" + OrgU8 + ",ou=" + OrgU2 + ",o=" + Org
      Print #1, "changetype: " + CType
      If ver = "NDS 7" Then Print #1, "objectClass: top"
      If ver = "NDS 8" Then Print #1, "ou: " + OrgU20
      Print #1, "objectClass: organizationalUnit"
      If ver = "NDS 8" Then Print #1, "objectClass: ndsLoginProperties"
      If ver = "NDS 8" Then Print #1, "objectClass: top"
      Print #1,
      
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU21 + ",ou=" + OrgU15 + ",ou=" + OrgU9 + ",ou=" + OrgU3 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU21 + ",ou=" + OrgU15 + ",ou=" + OrgU9 + ",ou=" + OrgU3 + ",o=" + Org
      Print #1, "changetype: " + CType
      If ver = "NDS 7" Then Print #1, "objectClass: top"
      If ver = "NDS 8" Then Print #1, "ou: " + OrgU21
      Print #1, "objectClass: organizationalUnit"
      If ver = "NDS 8" Then Print #1, "objectClass: ndsLoginProperties"
      If ver = "NDS 8" Then Print #1, "objectClass: top"
      Print #1,
      
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU22 + ",ou=" + OrgU16 + ",ou=" + OrgU10 + ",ou=" + OrgU4 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU22 + ",ou=" + OrgU16 + ",ou=" + OrgU10 + ",ou=" + OrgU4 + ",o=" + Org
      Print #1, "changetype: " + CType
      If ver = "NDS 7" Then Print #1, "objectClass: top"
      If ver = "NDS 8" Then Print #1, "ou: " + OrgU22
      Print #1, "objectClass: organizationalUnit"
      If ver = "NDS 8" Then Print #1, "objectClass: ndsLoginProperties"
      If ver = "NDS 8" Then Print #1, "objectClass: top"
      Print #1,
      
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU23 + ",ou=" + OrgU17 + ",ou=" + OrgU11 + ",ou=" + OrgU5 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU23 + ",ou=" + OrgU17 + ",ou=" + OrgU11 + ",ou=" + OrgU5 + ",o=" + Org
      Print #1, "changetype: " + CType
      If ver = "NDS 7" Then Print #1, "objectClass: top"
      If ver = "NDS 8" Then Print #1, "ou: " + OrgU23
      Print #1, "objectClass: organizationalUnit"
      If ver = "NDS 8" Then Print #1, "objectClass: ndsLoginProperties"
      If ver = "NDS 8" Then Print #1, "objectClass: top"
      Print #1,
      
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU24 + ",ou=" + OrgU18 + ",ou=" + OrgU12 + ",ou=" + OrgU6 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU24 + ",ou=" + OrgU18 + ",ou=" + OrgU12 + ",ou=" + OrgU6 + ",o=" + Org
      Print #1, "changetype: " + CType
      If ver = "NDS 7" Then Print #1, "objectClass: top"
      If ver = "NDS 8" Then Print #1, "ou: " + OrgU24
      Print #1, "objectClass: organizationalUnit"
      If ver = "NDS 8" Then Print #1, "objectClass: ndsLoginProperties"
      If ver = "NDS 8" Then Print #1, "objectClass: top"
      Print #1,
      
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU25 + ",ou=" + OrgU19 + ",ou=" + OrgU13 + ",ou=" + OrgU7 + ",ou=" + OrgU1 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU25 + ",ou=" + OrgU19 + ",ou=" + OrgU13 + ",ou=" + OrgU7 + ",ou=" + OrgU1 + ",o=" + Org
      Print #1, "changetype: " + CType
      If ver = "NDS 7" Then Print #1, "objectClass: top"
      If ver = "NDS 8" Then Print #1, "ou: " + OrgU25
      Print #1, "objectClass: organizationalUnit"
      If ver = "NDS 8" Then Print #1, "objectClass: ndsLoginProperties"
      If ver = "NDS 8" Then Print #1, "objectClass: top"
      Print #1,
      
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU26 + ",ou=" + OrgU20 + ",ou=" + OrgU14 + ",ou=" + OrgU8 + ",ou=" + OrgU2 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU26 + ",ou=" + OrgU20 + ",ou=" + OrgU14 + ",ou=" + OrgU8 + ",ou=" + OrgU2 + ",o=" + Org
      Print #1, "changetype: " + CType
      If ver = "NDS 7" Then Print #1, "objectClass: top"
      If ver = "NDS 8" Then Print #1, "ou: " + OrgU26
      Print #1, "objectClass: organizationalUnit"
      If ver = "NDS 8" Then Print #1, "objectClass: ndsLoginProperties"
      If ver = "NDS 8" Then Print #1, "objectClass: top"
      Print #1,
      
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU27 + ",ou=" + OrgU21 + ",ou=" + OrgU15 + ",ou=" + OrgU9 + ",ou=" + OrgU3 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU27 + ",ou=" + OrgU21 + ",ou=" + OrgU15 + ",ou=" + OrgU9 + ",ou=" + OrgU3 + ",o=" + Org
      Print #1, "changetype: " + CType
      If ver = "NDS 7" Then Print #1, "objectClass: top"
      If ver = "NDS 8" Then Print #1, "ou: " + OrgU27
      Print #1, "objectClass: organizationalUnit"
      If ver = "NDS 8" Then Print #1, "objectClass: ndsLoginProperties"
      If ver = "NDS 8" Then Print #1, "objectClass: top"
      Print #1,
      
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU28 + ",ou=" + OrgU22 + ",ou=" + OrgU16 + ",ou=" + OrgU10 + ",ou=" + OrgU4 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU28 + ",ou=" + OrgU22 + ",ou=" + OrgU16 + ",ou=" + OrgU10 + ",ou=" + OrgU4 + ",o=" + Org
      Print #1, "changetype: " + CType
      If ver = "NDS 7" Then Print #1, "objectClass: top"
      If ver = "NDS 8" Then Print #1, "ou: " + OrgU28
      Print #1, "objectClass: organizationalUnit"
      If ver = "NDS 8" Then Print #1, "objectClass: ndsLoginProperties"
      If ver = "NDS 8" Then Print #1, "objectClass: top"
      Print #1,
      
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU29 + ",ou=" + OrgU23 + ",ou=" + OrgU17 + ",ou=" + OrgU11 + ",ou=" + OrgU5 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU29 + ",ou=" + OrgU23 + ",ou=" + OrgU17 + ",ou=" + OrgU11 + ",ou=" + OrgU5 + ",o=" + Org
      Print #1, "changetype: " + CType
      If ver = "NDS 7" Then Print #1, "objectClass: top"
      If ver = "NDS 8" Then Print #1, "ou: " + OrgU29
      Print #1, "objectClass: organizationalUnit"
      If ver = "NDS 8" Then Print #1, "objectClass: ndsLoginProperties"
      If ver = "NDS 8" Then Print #1, "objectClass: top"
      Print #1,
      
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU30 + ",ou=" + OrgU24 + ",ou=" + OrgU18 + ",ou=" + OrgU12 + ",ou=" + OrgU6 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU30 + ",ou=" + OrgU24 + ",ou=" + OrgU18 + ",ou=" + OrgU12 + ",ou=" + OrgU6 + ",o=" + Org
      Print #1, "changetype: " + CType
      If ver = "NDS 7" Then Print #1, "objectClass: top"
      If ver = "NDS 8" Then Print #1, "ou: " + OrgU30
      Print #1, "objectClass: organizationalUnit"
      If ver = "NDS 8" Then Print #1, "objectClass: ndsLoginProperties"
      If ver = "NDS 8" Then Print #1, "objectClass: top"
      Print #1,
      
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU31 + ",ou=" + OrgU25 + ",ou=" + OrgU19 + ",ou=" + OrgU13 + ",ou=" + OrgU7 + ",ou=" + OrgU1 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU31 + ",ou=" + OrgU25 + ",ou=" + OrgU19 + ",ou=" + OrgU13 + ",ou=" + OrgU7 + ",ou=" + OrgU1 + ",o=" + Org
      Print #1, "changetype: " + CType
      If ver = "NDS 7" Then Print #1, "objectClass: top"
      If ver = "NDS 8" Then Print #1, "ou: " + OrgU31
      Print #1, "objectClass: organizationalUnit"
      If ver = "NDS 8" Then Print #1, "objectClass: ndsLoginProperties"
      If ver = "NDS 8" Then Print #1, "objectClass: top"
      Print #1,
      
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU32 + ",ou=" + OrgU26 + ",ou=" + OrgU20 + ",ou=" + OrgU14 + ",ou=" + OrgU8 + ",ou=" + OrgU2 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU32 + ",ou=" + OrgU26 + ",ou=" + OrgU20 + ",ou=" + OrgU14 + ",ou=" + OrgU8 + ",ou=" + OrgU2 + ",o=" + Org
      Print #1, "changetype: " + CType
      If ver = "NDS 7" Then Print #1, "objectClass: top"
      If ver = "NDS 8" Then Print #1, "ou: " + OrgU32
      Print #1, "objectClass: organizationalUnit"
      If ver = "NDS 8" Then Print #1, "objectClass: ndsLoginProperties"
      If ver = "NDS 8" Then Print #1, "objectClass: top"
      Print #1,
      
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU33 + ",ou=" + OrgU27 + ",ou=" + OrgU21 + ",ou=" + OrgU15 + ",ou=" + OrgU9 + ",ou=" + OrgU3 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU33 + ",ou=" + OrgU27 + ",ou=" + OrgU21 + ",ou=" + OrgU15 + ",ou=" + OrgU9 + ",ou=" + OrgU3 + ",o=" + Org
      Print #1, "changetype: " + CType
      If ver = "NDS 7" Then Print #1, "objectClass: top"
      If ver = "NDS 8" Then Print #1, "ou: " + OrgU33
      Print #1, "objectClass: organizationalUnit"
      If ver = "NDS 8" Then Print #1, "objectClass: ndsLoginProperties"
      If ver = "NDS 8" Then Print #1, "objectClass: top"
      Print #1,
      
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU34 + ",ou=" + OrgU28 + ",ou=" + OrgU22 + ",ou=" + OrgU16 + ",ou=" + OrgU10 + ",ou=" + OrgU4 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU34 + ",ou=" + OrgU28 + ",ou=" + OrgU22 + ",ou=" + OrgU16 + ",ou=" + OrgU10 + ",ou=" + OrgU4 + ",o=" + Org
      Print #1, "changetype: " + CType
      If ver = "NDS 7" Then Print #1, "objectClass: top"
      If ver = "NDS 8" Then Print #1, "ou: " + OrgU34
      Print #1, "objectClass: organizationalUnit"
      If ver = "NDS 8" Then Print #1, "objectClass: ndsLoginProperties"
      If ver = "NDS 8" Then Print #1, "objectClass: top"
      Print #1,
      
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU35 + ",ou=" + OrgU29 + ",ou=" + OrgU23 + ",ou=" + OrgU17 + ",ou=" + OrgU11 + ",ou=" + OrgU5 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU35 + ",ou=" + OrgU29 + ",ou=" + OrgU23 + ",ou=" + OrgU17 + ",ou=" + OrgU11 + ",ou=" + OrgU5 + ",o=" + Org
      Print #1, "changetype: " + CType
      If ver = "NDS 7" Then Print #1, "objectClass: top"
      If ver = "NDS 8" Then Print #1, "ou: " + OrgU35
      Print #1, "objectClass: organizationalUnit"
      If ver = "NDS 8" Then Print #1, "objectClass: ndsLoginProperties"
      If ver = "NDS 8" Then Print #1, "objectClass: top"
      Print #1,
      
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU36 + ",ou=" + OrgU30 + ",ou=" + OrgU24 + ",ou=" + OrgU18 + ",ou=" + OrgU12 + ",ou=" + OrgU6 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU36 + ",ou=" + OrgU30 + ",ou=" + OrgU24 + ",ou=" + OrgU18 + ",ou=" + OrgU12 + ",ou=" + OrgU6 + ",o=" + Org
      Print #1, "changetype: " + CType
      If ver = "NDS 7" Then Print #1, "objectClass: top"
      If ver = "NDS 8" Then Print #1, "ou: " + OrgU36
      Print #1, "objectClass: organizationalUnit"
      If ver = "NDS 8" Then Print #1, "objectClass: ndsLoginProperties"
      If ver = "NDS 8" Then Print #1, "objectClass: top"
'End Add Tree Information

BypassTreeAdd:
BypassTree:
BypassTree1:
   
   For i = beginval To endval
      ProgressBar1.Value = i
      
      'Start User data
      If UserSetOneCheck = 0 Then GoTo Step10 'If Create User(s) is not checked, bypass user creation for that set
      Print #1,
      If CustContCheck = 1 Then Print #1, "dn: cn=" + userid1 + Format(i) + CustCont
      If CustContCheck = 1 Then GoTo CCStep3
      If Not CCont = "" Then Print #1, "dn: cn=" + userid1 + Format(i) + ",ou=" + OrgUnit111 + ",ou=" + OrgUnit11 + ",ou=" + OrgUnit1 + ",o=" + Org1 + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: cn=" + userid1 + Format(i) + ",ou=" + OrgUnit111 + ",ou=" + OrgUnit11 + ",ou=" + OrgUnit1 + ",o=" + Org1
CCStep3:
      Print #1, "changetype: " + CType
      If CType = "delete" Then GoTo Step10
      If ver = "NDS 7" Then Print #1, "objectclass: top"
      If ver = "NDS 7" Then Print #1, "objectclass: person"
      If ver = "NDS 7" Then Print #1, "objectclass: organizationalPerson"
      If ver = "NDS 7" Then Print #1, "objectclass: inetOrgPerson"
      If ver = "NDS 7" Then Print #1, "mail: " + userid1 + Format(i) + "@" + domainname
      If ver = "NDS 7" Then Print #1, "givenName: " + givenname1 + Format(i)
      If ver = "NDS 7" Then Print #1, "sn: " + surname1 + Format(i)
      If ver = "NDS 8" Then Print #1, "mail: " + userid1 + Format(i) + "@" + domainname
      Print #1, "uid: " + userid1 + Format(i)
      If ver = "NDS 8" Then Print #1, "givenName: " + givenname1 + Format(i)
      If ver = "NDS 8" Then Print #1, "sn: " + surname1 + Format(i)
      If ver = "NDS 8" Then Print #1, "objectclass: inetOrgPerson"
      If ver = "NDS 8" Then Print #1, "objectclass: organizationalPerson"
      If ver = "NDS 8" Then Print #1, "objectclass: person"
      If ver = "NDS 8" Then Print #1, "objectclass: ndsLoginProperties"
      If ver = "NDS 8" Then Print #1, "objectclass: top"
      If passwordm = "Yes" Then Print #1, "userpassword: " + userid1 + Format(i) Else Print #1, "userpassword: " + password1
      If homedir = "cn=server1_Vol1,ou=OrganizationalUnit,o=Container#0#\Users_directory" Then GoTo Step10
      If ver = "NDS 8" Then Print #1, "ndsHomeDirectory: " + homedir + "\" + userid1 + Format(i)
      If ver = "NDS 7" Then Print #1, "homeDirectory: " + homedir + "\" + userid1 + Format(i)
           
Step10:

      If UserSetTwoCheck = 0 Then GoTo Step11 'If Create User(s) is not checked, bypass user creation for that set
      Print #1,
      If Not CCont = "" Then Print #1, "dn: cn=" + userid2 + Format(i) + ",ou=" + OrgUnit222 + ",ou=" + OrgUnit22 + ",ou=" + OrgUnit2 + ",o=" + Org2 + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: cn=" + userid2 + Format(i) + ",ou=" + OrgUnit222 + ",ou=" + OrgUnit22 + ",ou=" + OrgUnit2 + ",o=" + Org2
      Print #1, "changetype: " + CType
      If CType = "delete" Then GoTo Step11
      If ver = "NDS 7" Then Print #1, "objectclass: top"
      If ver = "NDS 7" Then Print #1, "objectclass: person"
      If ver = "NDS 7" Then Print #1, "objectclass: organizationalPerson"
      If ver = "NDS 7" Then Print #1, "objectclass: inetOrgPerson"
      If ver = "NDS 7" Then Print #1, "mail: " + userid2 + Format(i) + "@" + domainname
      If ver = "NDS 7" Then Print #1, "givenName: " + givenname2 + Format(i)
      If ver = "NDS 7" Then Print #1, "sn: " + surname2 + Format(i)
      If ver = "NDS 8" Then Print #1, "mail: " + userid2 + Format(i) + "@" + domainname
      Print #1, "uid: " + userid2 + Format(i)
      If ver = "NDS 8" Then Print #1, "givenName: " + givenname2 + Format(i)
      If ver = "NDS 8" Then Print #1, "sn: " + surname2 + Format(i)
      If ver = "NDS 8" Then Print #1, "objectclass: inetOrgPerson"
      If ver = "NDS 8" Then Print #1, "objectclass: organizationalPerson"
      If ver = "NDS 8" Then Print #1, "objectclass: person"
      If ver = "NDS 8" Then Print #1, "objectclass: ndsLoginProperties"
      If ver = "NDS 8" Then Print #1, "objectclass: top"
      If passwordm = "Yes" Then Print #1, "userpassword: " + userid1 + Format(i) Else Print #1, "userpassword: " + password2
      If homedir = "cn=server1_Vol1,ou=OrganizationalUnit,o=Container#0#\Users_directory" Then GoTo Step11
      If ver = "NDS 8" Then Print #1, "ndsHomeDirectory: " + homedir + "\" + userid2 + Format(i)
      If ver = "NDS 7" Then Print #1, "homeDirectory: " + homedir + "\" + userid2 + Format(i)
            
Step11:
      
      If UserSetThreeCheck = 0 Then GoTo Step12 'If Create User(s) is not checked, bypass user creation for that set
      Print #1,
      If Not CCont = "" Then Print #1, "dn: cn=" + userid3 + Format(i) + ",ou=" + OrgUnit333 + ",ou=" + OrgUnit33 + ",ou=" + OrgUnit3 + ",o=" + Org3 + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: cn=" + userid3 + Format(i) + ",ou=" + OrgUnit333 + ",ou=" + OrgUnit33 + ",ou=" + OrgUnit3 + ",o=" + Org3
      Print #1, "changetype: " + CType
      If CType = "delete" Then GoTo Step12
      If ver = "NDS 7" Then Print #1, "objectclass: top"
      If ver = "NDS 7" Then Print #1, "objectclass: person"
      If ver = "NDS 7" Then Print #1, "objectclass: organizationalPerson"
      If ver = "NDS 7" Then Print #1, "objectclass: inetOrgPerson"
      If ver = "NDS 7" Then Print #1, "mail: " + userid3 + Format(i) + "@" + domainname
      If ver = "NDS 7" Then Print #1, "givenName: " + givenname3 + Format(i)
      If ver = "NDS 7" Then Print #1, "sn: " + surname3 + Format(i)
      If ver = "NDS 8" Then Print #1, "mail: " + userid3 + Format(i) + "@" + domainname
      Print #1, "uid: " + userid3 + Format(i)
      If ver = "NDS 8" Then Print #1, "givenName: " + givenname3 + Format(i)
      If ver = "NDS 8" Then Print #1, "sn: " + surname3 + Format(i)
      If ver = "NDS 8" Then Print #1, "objectclass: inetOrgPerson"
      If ver = "NDS 8" Then Print #1, "objectclass: organizationalPerson"
      If ver = "NDS 8" Then Print #1, "objectclass: person"
      If ver = "NDS 8" Then Print #1, "objectclass: ndsLoginProperties"
      If ver = "NDS 8" Then Print #1, "objectclass: top"
      If passwordm = "Yes" Then Print #1, "userpassword: " + userid3 + Format(i) Else Print #1, "userpassword: " + password3
      If homedir = "cn=server1_Vol1,ou=OrganizationalUnit,o=Container#0#\Users_directory" Then GoTo Step12
      If ver = "NDS 8" Then Print #1, "ndsHomeDirectory: " + homedir + "\" + userid3 + Format(i)
      If ver = "NDS 7" Then Print #1, "homeDirectory: " + homedir + "\" + userid3 + Format(i)
            
Step12:
            
      If UserSetFourCheck = 0 Then GoTo Step13 'If Create User(s) is not checked, bypass user creation for that set
      Print #1,
      If Not CCont = "" Then Print #1, "dn: cn=" + userid4 + Format(i) + ",ou=" + OrgUnit444 + ",ou=" + OrgUnit44 + ",ou=" + OrgUnit4 + ",o=" + Org4 + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: cn=" + userid4 + Format(i) + ",ou=" + OrgUnit444 + ",ou=" + OrgUnit44 + ",ou=" + OrgUnit4 + ",o=" + Org4
      Print #1, "changetype: " + CType
      If CType = "delete" Then GoTo Step13
      If ver = "NDS 7" Then Print #1, "objectclass: top"
      If ver = "NDS 7" Then Print #1, "objectclass: person"
      If ver = "NDS 7" Then Print #1, "objectclass: organizationalPerson"
      If ver = "NDS 7" Then Print #1, "objectclass: inetOrgPerson"
      If ver = "NDS 7" Then Print #1, "mail: " + userid4 + Format(i) + "@" + domainname
      If ver = "NDS 7" Then Print #1, "givenName: " + givenname4 + Format(i)
      If ver = "NDS 7" Then Print #1, "sn: " + surname4 + Format(i)
      If ver = "NDS 8" Then Print #1, "mail: " + userid4 + Format(i) + "@" + domainname
      Print #1, "uid: " + userid1 + Format(i)
      If ver = "NDS 8" Then Print #1, "givenName: " + givenname4 + Format(i)
      If ver = "NDS 8" Then Print #1, "sn: " + surname4 + Format(i)
      If ver = "NDS 8" Then Print #1, "objectclass: inetOrgPerson"
      If ver = "NDS 8" Then Print #1, "objectclass: organizationalPerson"
      If ver = "NDS 8" Then Print #1, "objectclass: person"
      If ver = "NDS 8" Then Print #1, "objectclass: ndsLoginProperties"
      If ver = "NDS 8" Then Print #1, "objectclass: top"
      If passwordm = "Yes" Then Print #1, "userpassword: " + userid4 + Format(i) Else Print #1, "userpassword: " + password4
      If homedir = "cn=server1_Vol1,ou=OrganizationalUnit,o=Container#0#\Users_directory" Then GoTo Step13
      If ver = "NDS 8" Then Print #1, "ndsHomeDirectory: " + homedir + "\" + userid4 + Format(i)
      If ver = "NDS 7" Then Print #1, "homeDirectory: " + homedir + "\" + userid4 + Format(i)
            
      'End User Data
      
Step13:
      ProgressBar1.Visible = True
      
   Next i
      
'Start Delete Tree Information
If SkipTreeCreateCheck = "0" Then GoTo BypassTree2
Step41:

   If CheckAdd = 1 Then GoTo ByPassDel

      If CheckDel = 1 Then Print #1,
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU36 + ",ou=" + OrgU30 + ",ou=" + OrgU24 + ",ou=" + OrgU18 + ",ou=" + OrgU12 + ",ou=" + OrgU6 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU36 + ",ou=" + OrgU30 + ",ou=" + OrgU24 + ",ou=" + OrgU18 + ",ou=" + OrgU12 + ",ou=" + OrgU6 + ",o=" + Org
      Print #1, "changetype: delete"
      Print #1,
      
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU35 + ",ou=" + OrgU29 + ",ou=" + OrgU23 + ",ou=" + OrgU17 + ",ou=" + OrgU11 + ",ou=" + OrgU5 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU35 + ",ou=" + OrgU29 + ",ou=" + OrgU23 + ",ou=" + OrgU17 + ",ou=" + OrgU11 + ",ou=" + OrgU5 + ",o=" + Org
      Print #1, "changetype: delete"
      Print #1,
            
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU34 + ",ou=" + OrgU28 + ",ou=" + OrgU22 + ",ou=" + OrgU16 + ",ou=" + OrgU10 + ",ou=" + OrgU4 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU34 + ",ou=" + OrgU28 + ",ou=" + OrgU22 + ",ou=" + OrgU16 + ",ou=" + OrgU10 + ",ou=" + OrgU4 + ",o=" + Org
      Print #1, "changetype: delete"
      Print #1,
    
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU33 + ",ou=" + OrgU27 + ",ou=" + OrgU21 + ",ou=" + OrgU15 + ",ou=" + OrgU9 + ",ou=" + OrgU3 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU33 + ",ou=" + OrgU27 + ",ou=" + OrgU21 + ",ou=" + OrgU15 + ",ou=" + OrgU9 + ",ou=" + OrgU3 + ",o=" + Org
      Print #1, "changetype: delete"
      Print #1,
      
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU32 + ",ou=" + OrgU26 + ",ou=" + OrgU20 + ",ou=" + OrgU14 + ",ou=" + OrgU8 + ",ou=" + OrgU2 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU32 + ",ou=" + OrgU26 + ",ou=" + OrgU20 + ",ou=" + OrgU14 + ",ou=" + OrgU8 + ",ou=" + OrgU2 + ",o=" + Org
      Print #1, "changetype: delete"
      Print #1,
      
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU31 + ",ou=" + OrgU25 + ",ou=" + OrgU19 + ",ou=" + OrgU13 + ",ou=" + OrgU7 + ",ou=" + OrgU1 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU31 + ",ou=" + OrgU25 + ",ou=" + OrgU19 + ",ou=" + OrgU13 + ",ou=" + OrgU7 + ",ou=" + OrgU1 + ",o=" + Org
      Print #1, "changetype: delete"
      Print #1,
      
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU30 + ",ou=" + OrgU24 + ",ou=" + OrgU18 + ",ou=" + OrgU12 + ",ou=" + OrgU6 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU30 + ",ou=" + OrgU24 + ",ou=" + OrgU18 + ",ou=" + OrgU12 + ",ou=" + OrgU6 + ",o=" + Org
      Print #1, "changetype: delete"
      Print #1,
      
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU29 + ",ou=" + OrgU23 + ",ou=" + OrgU17 + ",ou=" + OrgU11 + ",ou=" + OrgU5 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU29 + ",ou=" + OrgU23 + ",ou=" + OrgU17 + ",ou=" + OrgU11 + ",ou=" + OrgU5 + ",o=" + Org
      Print #1, "changetype: delete"
      Print #1,
      
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU28 + ",ou=" + OrgU22 + ",ou=" + OrgU16 + ",ou=" + OrgU10 + ",ou=" + OrgU4 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU28 + ",ou=" + OrgU22 + ",ou=" + OrgU16 + ",ou=" + OrgU10 + ",ou=" + OrgU4 + ",o=" + Org
      Print #1, "changetype: delete"
      Print #1,
      
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU27 + ",ou=" + OrgU21 + ",ou=" + OrgU15 + ",ou=" + OrgU9 + ",ou=" + OrgU3 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU27 + ",ou=" + OrgU21 + ",ou=" + OrgU15 + ",ou=" + OrgU9 + ",ou=" + OrgU3 + ",o=" + Org
      Print #1, "changetype: delete"
      Print #1,
      
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU26 + ",ou=" + OrgU20 + ",ou=" + OrgU14 + ",ou=" + OrgU8 + ",ou=" + OrgU2 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU26 + ",ou=" + OrgU20 + ",ou=" + OrgU14 + ",ou=" + OrgU8 + ",ou=" + OrgU2 + ",o=" + Org
      Print #1, "changetype: delete"
      Print #1,
      
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU25 + ",ou=" + OrgU19 + ",ou=" + OrgU13 + ",ou=" + OrgU7 + ",ou=" + OrgU1 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU25 + ",ou=" + OrgU19 + ",ou=" + OrgU13 + ",ou=" + OrgU7 + ",ou=" + OrgU1 + ",o=" + Org
      Print #1, "changetype: delete"
      Print #1,
      
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU24 + ",ou=" + OrgU18 + ",ou=" + OrgU12 + ",ou=" + OrgU6 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU24 + ",ou=" + OrgU18 + ",ou=" + OrgU12 + ",ou=" + OrgU6 + ",o=" + Org
      Print #1, "changetype: delete"
      Print #1,
      
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU23 + ",ou=" + OrgU17 + ",ou=" + OrgU11 + ",ou=" + OrgU5 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU23 + ",ou=" + OrgU17 + ",ou=" + OrgU11 + ",ou=" + OrgU5 + ",o=" + Org
      Print #1, "changetype: delete"
      Print #1,
      
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU22 + ",ou=" + OrgU16 + ",ou=" + OrgU10 + ",ou=" + OrgU4 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU22 + ",ou=" + OrgU16 + ",ou=" + OrgU10 + ",ou=" + OrgU4 + ",o=" + Org
      Print #1, "changetype: delete"
      Print #1,
      
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU21 + ",ou=" + OrgU15 + ",ou=" + OrgU9 + ",ou=" + OrgU3 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU21 + ",ou=" + OrgU15 + ",ou=" + OrgU9 + ",ou=" + OrgU3 + ",o=" + Org
      Print #1, "changetype: delete"
      Print #1,
      
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU20 + ",ou=" + OrgU14 + ",ou=" + OrgU8 + ",ou=" + OrgU2 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU20 + ",ou=" + OrgU14 + ",ou=" + OrgU8 + ",ou=" + OrgU2 + ",o=" + Org
      Print #1, "changetype: delete"
      Print #1,
      
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU19 + ",ou=" + OrgU13 + ",ou=" + OrgU7 + ",ou=" + OrgU1 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU19 + ",ou=" + OrgU13 + ",ou=" + OrgU7 + ",ou=" + OrgU1 + ",o=" + Org
      Print #1, "changetype: delete"
      Print #1,
      
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU18 + ",ou=" + OrgU12 + ",ou=" + OrgU6 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU18 + ",ou=" + OrgU12 + ",ou=" + OrgU6 + ",o=" + Org
      Print #1, "changetype: delete"
      Print #1,
      
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU17 + ",ou=" + OrgU11 + ",ou=" + OrgU5 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU17 + ",ou=" + OrgU11 + ",ou=" + OrgU5 + ",o=" + Org
      Print #1, "changetype: delete"
      Print #1,
      
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU16 + ",ou=" + OrgU10 + ",ou=" + OrgU4 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU16 + ",ou=" + OrgU10 + ",ou=" + OrgU4 + ",o=" + Org
      Print #1, "changetype: delete"
      Print #1,
      
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU15 + ",ou=" + OrgU9 + ",ou=" + OrgU3 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU15 + ",ou=" + OrgU9 + ",ou=" + OrgU3 + ",o=" + Org
      Print #1, "changetype: delete"
      Print #1,
      
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU14 + ",ou=" + OrgU8 + ",ou=" + OrgU2 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU14 + ",ou=" + OrgU8 + ",ou=" + OrgU2 + ",o=" + Org
      Print #1, "changetype: delete"
      Print #1,
      
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU13 + ",ou=" + OrgU7 + ",ou=" + OrgU1 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU13 + ",ou=" + OrgU7 + ",ou=" + OrgU1 + ",o=" + Org
      Print #1, "changetype: delete"
      Print #1,
      
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU12 + ",ou=" + OrgU6 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU12 + ",ou=" + OrgU6 + ",o=" + Org
      Print #1, "changetype: delete"
      Print #1,
      
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU11 + ",ou=" + OrgU5 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU11 + ",ou=" + OrgU5 + ",o=" + Org
      Print #1, "changetype: delete"
      Print #1,
      
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU10 + ",ou=" + OrgU4 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU10 + ",ou=" + OrgU4 + ",o=" + Org
      Print #1, "changetype: delete"
      Print #1,
      
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU9 + ",ou=" + OrgU3 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU9 + ",ou=" + OrgU3 + ",o=" + Org
      Print #1, "changetype: delete"
      Print #1,
      
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU8 + ",ou=" + OrgU2 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU8 + ",ou=" + OrgU2 + ",o=" + Org
      Print #1, "changetype: delete"
      Print #1,
      
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU7 + ",ou=" + OrgU1 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU7 + ",ou=" + OrgU1 + ",o=" + Org
      Print #1, "changetype: delete"
      Print #1,
      
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU6 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU6 + ",o=" + Org
      Print #1, "changetype: delete"
      Print #1,
      
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU5 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU5 + ",o=" + Org
      Print #1, "changetype: delete"
      Print #1,
      
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU4 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU4 + ",o=" + Org
      Print #1, "changetype: delete"
      Print #1,
      
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU3 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU3 + ",o=" + Org
      Print #1, "changetype: delete"
      Print #1,
      
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU2 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU2 + ",o=" + Org
      Print #1, "changetype: delete"
      Print #1,
      
      If Not CCont = "" Then Print #1, "dn: ou=" + OrgU1 + ",o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: ou=" + OrgU1 + ",o=" + Org
      Print #1, "changetype: delete"
      Print #1,
      
      If Not CCont = "" Then Print #1, "dn: o=" + Org + ",c=" + CCont
      If CCont = "" Then Print #1, "dn: o=" + Org
      Print #1, "changetype: delete"
      Print #1,
            
If CCont = "" Then GoTo ByPassDel

      Print #1, "dn: c=" + CCont
      Print #1, "changetype: delete"
      Print #1,

ByPassDel:
BypassTree2:

'End Delete Tree Information
   
   Close #1
   
Step2:
CustLDIFStep1:

'Start ice Batch
Dim Sice As String
Dim Source_Sw As String
Dim Destination_Sw As String
Dim LDIF_Sw As String
Dim LDAP_Sw As String
Dim IP_Sw As String
Dim IP As String
Dim Prt_Sw As String
Dim Prt As String
Dim Uname_Sw As String
Dim Uname As String
Dim Pwd_Sw As String
Dim Pwd As String
Dim BDN_Sw As String
Dim BDN As String
Dim Srch_Sw As String
Dim Srch As String
Dim RtCrt_Sw As String
Dim RtCrt As String
Dim StopOnErr As String
Dim F_Sw As String
Dim FOut As String
Dim FIn As String
Dim CFile As String
Dim Sp As String

Sice = "ice" + " "
Source_Sw = "-S" + " "
Destination_Sw = "-D" + " "
LDIF_Sw = "LDIF" + " "
LDAP_Sw = "LDAP" + " "
IP_Sw = "-s" + " "
IP = Text1.Text + " "
Prt_Sw = "-p" + " "
Prt = Combo1.Text + " "
Uname_Sw = "-d" + " "
Uname = Text2.Text + " "
Pwd_Sw = "-w" + " "
Pwd = Text3.Text + " "
BDN_Sw = "-b" + " "
BDN = Text14.Text + " "
Srch_Sw = "-c" + " "
Srch = Text31.Text + " "
RtCrt_Sw = "-L" + " "
RtCrt = Text4.Text + " "
F_Sw = "-f" + " "
FOut = Text6.Text + " "
FIn = Text7.Text + " "
StopOnErr = "-c" + " "
CFile = Text30.Text + " "
Sp = " "

If IP = "" + " " Then   ' IP = ""
      MyVar = MsgBox("Please specify an IPAddress.", 0, "Specify IP Address") ' Perform some action.
      If MyVar = vbOK Then GoTo Step161
End If

If Prt = "" + " " Then   ' Prt = ""
      MyVar = MsgBox("Please specify a Port.", 0, "Specify Port") ' Perform some action.
      If MyVar = vbOK Then GoTo Step161
End If

If Uname = " " Then Uname_Sw = ""
If Uname = " " Then Uname = ""
If Pwd = " " Then Pwd_Sw = ""
If Pwd = " " Then Pwd = ""
If BDN = " " Then BDN_Sw = ""
If BDN = " " Then BDN = ""
If Srch = " " Then Srch_Sw = ""
If Srch = " " Then Srch = ""
If RtCrt = " " Then RtCrt_Sw = ""
If RtCrt = " " Then RtCrt = ""
If RtCrt = " " Then RtCrt_Sw = ""
If RtCrt = " " Then RtCrt = ""
If CustomCheck = 1 Then FOut = CFile

      Open RICEBatch For Output As #2
      Print #2, "path %PATH%;" + App.Path
      Print #2, "del ice.Log"
      If RetrieveTreeCheck = 1 Then GoTo Chk2Rtv:

'Send LDAP
If StopOnErrCheck = 0 Then Print #2, Sice + Source_Sw + LDIF_Sw + F_Sw + FOut + StopOnErr + Destination_Sw + LDAP_Sw + IP_Sw + IP + Prt_Sw + Prt + Uname_Sw + Uname + Pwd_Sw + Pwd + RtCrt_Sw + RtCrt
If StopOnErrCheck = 1 Then Print #2, Sice + Source_Sw + LDIF_Sw + F_Sw + FOut + Destination_Sw + LDAP_Sw + IP_Sw + IP + Prt_Sw + Prt + Uname_Sw + Uname + Pwd_Sw + Pwd + RtCrt_Sw + RtCrt

Chk2Rtv:
'Get LDAP
If RetrieveTreeCheck = 1 Then Print #2, Sice + Source_Sw + LDAP_Sw + IP_Sw + IP + Prt_Sw + Prt + Uname_Sw + Uname + Pwd_Sw + Pwd + BDN_Sw + BDN + Srch_Sw + Srch + RtCrt_Sw + RtCrt + Destination_Sw + LDIF_Sw + F_Sw + FIn

   Close #2
'End ice Batch

'Start Display Page
'ICE.Show 0
   Dim JobToDo As String
      JobToDo = RICEBatch
      DoEvents: Sleep 100
'      Shell32Bit JobToDo
'End Display Page

'Start User Home Directory Creation Section
   Dim beginval1 As Long
   Dim endval1 As Long
   
   beginval1 = 1
   endval1 = Val(Combo2.Text)
   
If UserHomeCheck = 0 Then GoTo Step161
If homedir = "cn=server1_Vol1,ou=OrganizationalUnit,o=Container#0#\Users_directory" Then GoTo Step161
          If (fso.FolderExists(MyPath)) Then GoTo Step1000 Else MkDir MyPath 'Make new directory or folder.
Step1000:
          ChDrive MyPath
          ChDir MyPath 'Changes the current directory or folder.
If endval = "1" Then GoTo Step150
      ProgressBar1.Min = beginval1
      ProgressBar1.Max = endval1
      ProgressBar1.Value = ProgressBar1.Min
      ProgressBar1.Visible = False

Step150:
      For j = beginval1 To endval1
         ProgressBar1.Value = j
   If userid1 = "" Then GoTo Step151
   If UserSetOneCheck = 0 Then GoTo Step151
      If (fso.FolderExists(MyPath + userid1 + Format(j))) Then GoTo CStep1 Else MkDir MyPath + userid1 + Format(j)  'Make new directory or folder.
CStep1:
      ChDir MyPath + userid1 + Format(j)
      If (fso.FolderExists(UHomeDir)) Then GoTo CStep2 Else MkDir UHomeDir  'Make new directory or folder.
CStep2:
      ChDir UHomeDir
      Open UHomeFile For Output As #2
      Print #2, userid1 + Format(j)
      ChDir "cd ..\.."
      Close #2
            
Step151:
   If userid2 = "" Then GoTo Step152
   If UserSetTwoCheck = 0 Then GoTo Step152
      If (fso.FolderExists(MyPath + userid2 + Format(j))) Then GoTo CStep3 Else MkDir MyPath + userid2 + Format(j) 'Make new directory or folder.
CStep3:
      ChDir MyPath + userid2 + Format(j)
      If (fso.FolderExists(UHomeDir)) Then GoTo CStep4 Else MkDir UHomeDir  'Make new directory or folder.
CStep4:
      ChDir UHomeDir
      Open UHomeFile For Output As #2
      Print #2, userid2 + Format(j)
      ChDir "cd ..\.."
      Close #2
         
Step152:
   If userid3 = "" Then GoTo Step153
   If UserSetThreeCheck = 0 Then GoTo Step153
      If (fso.FolderExists(MyPath + userid3 + Format(j))) Then GoTo CStep5 Else MkDir MyPath + userid3 + Format(j)  'Make new directory or folder.
CStep5:
      ChDir MyPath + userid3 + Format(j)
      If (fso.FolderExists(UHomeDir)) Then GoTo CStep6 Else MkDir UHomeDir  'Make new directory or folder.
CStep6:
      ChDir UHomeDir
      Open UHomeFile For Output As #2
      Print #2, userid3 + Format(j)
      ChDir "cd ..\.."
      Close #2
      
Step153:
   If userid4 = "" Then GoTo Step154
   If UserSetFourCheck = 0 Then GoTo Step154
      If (fso.FolderExists(MyPath + userid4 + Format(j))) Then GoTo CStep7 Else MkDir MyPath + userid4 + Format(j)  'Make new directory or folder.
CStep7:
      ChDir MyPath + userid4 + Format(j)
      If (fso.FolderExists(UHomeDir)) Then GoTo CStep8 Else MkDir UHomeDir  'Make new directory or folder.
CStep8:
      ChDir UHomeDir
      Open UHomeFile For Output As #2
      Print #2, userid4 + Format(j)
      ChDir "cd ..\.."
      Close #2
      
Step154:
      
If endval = "1" Then GoTo Step160
ProgressBar1.Visible = True
   Next j
      ChDir WPath
Step160:
     ChDrive "C:"
     ChDir App.Path
Step161:
'End User Home Directory Creation Section

ProgressBar1.Visible = True
ProgressBar1.Value = ProgressBar1.Max
End Sub

Sub Shell32Bit(ByVal JobToDo As String)
    Dim hProcess As Long
    Dim RetVal As Long
    Dim WPath
    Dim IceLog As String
    Dim MyString
        
    WPath = Text19.Text
    IceLog = "ice.log"
    
    hProcess = OpenProcess(PROCESS_QUERY_INFORMATION, False, Shell(JobToDo, 0)) 'The next line launches JobToDo as icon, captures process ID
    Do
       GetExitCodeProcess hProcess, RetVal 'Get the status of the process
            DoEvents: Sleep 1000 'Sleep command recommended as well as DoEvents

            Dim Num_Apps As Integer, NewFile As Integer
            Dim File_Data As String, DosCmd As String

            NewFile = FreeFile 'Display the filenames in the Text Box.
            Sleep 100
            
               Do While FileExists(WPath & IceLog) = "False" 'Make sure the file exists
               Sleep 1000
               Loop

                  Open (WPath & IceLog) For Input As #NewFile
                  ICE.Text1.Text = ""
                  While Not EOF(NewFile)
                  Line Input #NewFile, File_Data
                  ICE.Text1.Text = ICE.Text1.Text & File_Data & Chr(13) & Chr(10)
                  Wend
                  Close #NewFile
          
               Loop While RetVal = STILL_ACTIVE 'Loop while the process is active
               MyVar = MsgBox("The Update is Complete, Press OK to close the ICE Log or Cancel to review.", vbOKCancel, "Update Complete!")
               If MyVar = vbOK Then
               Unload ICE 'MsgBox ("OK Pressed")
               End If
      
End Sub
