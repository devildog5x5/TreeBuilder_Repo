VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tree Builder Version 3"
   ClientHeight    =   8070
   ClientLeft      =   2565
   ClientTop       =   1500
   ClientWidth     =   9990
   Icon            =   "frmOptions1.frx":0000
   KeyPreview      =   -1  'True
   ScaleHeight     =   8070
   ScaleWidth      =   9990
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   240
      TabIndex        =   166
      Top             =   7080
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton CmdWrite 
      Caption         =   "Write Files"
      Height          =   495
      Left            =   1800
      TabIndex        =   35
      Top             =   7440
      Width           =   1335
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
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample4 
         Caption         =   "Sample 4"
         Height          =   1785
         Left            =   2100
         TabIndex        =   43
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
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample3 
         Caption         =   "Sample 3"
         Height          =   1785
         Left            =   1545
         TabIndex        =   42
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
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample2 
         Caption         =   "Sample 2"
         Height          =   1785
         Left            =   645
         TabIndex        =   41
         Top             =   300
         Width           =   2055
      End
   End
   Begin VB.CommandButton CmdViewICE 
      Caption         =   "View ICE Log"
      Height          =   495
      Left            =   3360
      TabIndex        =   36
      Top             =   7440
      Width           =   1335
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
      Left            =   4920
      TabIndex        =   37
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
      TabIndex        =   34
      Top             =   7440
      Width           =   1335
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8055
      Left            =   0
      TabIndex        =   44
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
      TabPicture(0)   =   "frmOptions1.frx":0CCA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label54"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label53"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label51"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label50"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label49"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label48"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label28"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label30"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label10"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label11"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label9"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label14"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label12"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label18"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label13"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Check1"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "CmdBrowse2"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "CmdBrowse1"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Text5"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Text4"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Text3"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Text2"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Text1"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Combo2"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Combo6"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Text6"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "CmdBrowse3"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "Check2"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "Check3"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "Text7"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "CmdBrowse4"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "Check4"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "Text14"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "Check5"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "cmdViewExp"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "cmdViewImp"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "Check6"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "Combo1"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "Text19"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "CmdBrowse5"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "cmdViewRice"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "cmdClean"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "cmdViewCust"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "CmdBrowse10"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "Text30"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "Check9"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).ControlCount=   46
      TabCaption(1)   =   "User Configuration"
      TabPicture(1)   =   "frmOptions1.frx":0CE6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label44"
      Tab(1).Control(1)=   "Label43"
      Tab(1).Control(2)=   "Label47"
      Tab(1).Control(3)=   "Label31"
      Tab(1).Control(4)=   "Label29"
      Tab(1).Control(5)=   "Label27"
      Tab(1).Control(6)=   "Label25"
      Tab(1).Control(7)=   "Label20"
      Tab(1).Control(8)=   "Label15"
      Tab(1).Control(9)=   "Label8"
      Tab(1).Control(10)=   "Label4"
      Tab(1).Control(11)=   "Label7"
      Tab(1).Control(12)=   "Label6"
      Tab(1).Control(13)=   "Label5"
      Tab(1).Control(14)=   "Label100"
      Tab(1).Control(15)=   "Label101"
      Tab(1).Control(16)=   "Label3"
      Tab(1).Control(17)=   "Label1"
      Tab(1).Control(18)=   "Label2"
      Tab(1).Control(19)=   "Label16"
      Tab(1).Control(20)=   "Text103"
      Tab(1).Control(21)=   "Text107"
      Tab(1).Control(22)=   "Text111"
      Tab(1).Control(23)=   "Text115"
      Tab(1).Control(24)=   "Text102"
      Tab(1).Control(25)=   "Text106"
      Tab(1).Control(26)=   "Text110"
      Tab(1).Control(27)=   "Text114"
      Tab(1).Control(28)=   "Text100"
      Tab(1).Control(29)=   "Text104"
      Tab(1).Control(30)=   "Text108"
      Tab(1).Control(31)=   "Text112"
      Tab(1).Control(32)=   "Text9"
      Tab(1).Control(33)=   "Text8"
      Tab(1).Control(34)=   "Combo14"
      Tab(1).Control(35)=   "Text28"
      Tab(1).Control(36)=   "Text25"
      Tab(1).Control(37)=   "Text26"
      Tab(1).Control(38)=   "Text27"
      Tab(1).Control(39)=   "Text23"
      Tab(1).Control(40)=   "Text20"
      Tab(1).Control(41)=   "Text21"
      Tab(1).Control(42)=   "Text22"
      Tab(1).Control(43)=   "Text18"
      Tab(1).Control(44)=   "Text15"
      Tab(1).Control(45)=   "Text16"
      Tab(1).Control(46)=   "Text17"
      Tab(1).Control(47)=   "Text13"
      Tab(1).Control(48)=   "Text12"
      Tab(1).Control(49)=   "Text11"
      Tab(1).Control(50)=   "Text10"
      Tab(1).Control(51)=   "Check21"
      Tab(1).Control(52)=   "Check22"
      Tab(1).Control(53)=   "Check23"
      Tab(1).Control(54)=   "Check24"
      Tab(1).Control(55)=   "Text113"
      Tab(1).Control(56)=   "Text109"
      Tab(1).Control(57)=   "Text105"
      Tab(1).Control(58)=   "Text101"
      Tab(1).Control(59)=   "Check7"
      Tab(1).Control(60)=   "Text29"
      Tab(1).Control(61)=   "Check8"
      Tab(1).ControlCount=   62
      TabCaption(2)   =   "Tree Configuration"
      TabPicture(2)   =   "frmOptions1.frx":0D02
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label64"
      Tab(2).Control(1)=   "Label65"
      Tab(2).Control(2)=   "Label66"
      Tab(2).Control(3)=   "Label17"
      Tab(2).Control(4)=   "Label23"
      Tab(2).Control(5)=   "Combo199"
      Tab(2).Control(6)=   "Combo104"
      Tab(2).Control(7)=   "Combo105"
      Tab(2).Control(8)=   "Combo106"
      Tab(2).Control(9)=   "Combo204"
      Tab(2).Control(10)=   "Combo205"
      Tab(2).Control(11)=   "Combo206"
      Tab(2).Control(12)=   "Combo304"
      Tab(2).Control(13)=   "Combo305"
      Tab(2).Control(14)=   "Combo306"
      Tab(2).Control(15)=   "Combo404"
      Tab(2).Control(16)=   "Combo405"
      Tab(2).Control(17)=   "Combo406"
      Tab(2).Control(18)=   "Combo504"
      Tab(2).Control(19)=   "Combo505"
      Tab(2).Control(20)=   "Combo506"
      Tab(2).Control(21)=   "Combo604"
      Tab(2).Control(22)=   "Combo605"
      Tab(2).Control(23)=   "Combo606"
      Tab(2).Control(24)=   "Text200"
      Tab(2).Control(25)=   "Text1101"
      Tab(2).Control(26)=   "Text1201"
      Tab(2).Control(27)=   "Text1301"
      Tab(2).Control(28)=   "Text1401"
      Tab(2).Control(29)=   "Text1501"
      Tab(2).Control(30)=   "Text1601"
      Tab(2).Control(31)=   "Text1102"
      Tab(2).Control(32)=   "Text1202"
      Tab(2).Control(33)=   "Text1302"
      Tab(2).Control(34)=   "Text1402"
      Tab(2).Control(35)=   "Text1502"
      Tab(2).Control(36)=   "Text1602"
      Tab(2).Control(37)=   "Text1103"
      Tab(2).Control(38)=   "Text1203"
      Tab(2).Control(39)=   "Text1303"
      Tab(2).Control(40)=   "Text1403"
      Tab(2).Control(41)=   "Text1503"
      Tab(2).Control(42)=   "Text1603"
      Tab(2).Control(43)=   "Text24"
      Tab(2).Control(44)=   "Check10"
      Tab(2).ControlCount=   45
      TabCaption(3)   =   "Home Dir Configuration"
      TabPicture(3)   =   "frmOptions1.frx":0D1E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label45"
      Tab(3).Control(1)=   "Label41"
      Tab(3).Control(2)=   "Label40"
      Tab(3).Control(3)=   "Label19"
      Tab(3).Control(4)=   "Label21"
      Tab(3).Control(5)=   "Label22"
      Tab(3).Control(6)=   "CmdBrowse6"
      Tab(3).Control(7)=   "Text403"
      Tab(3).Control(8)=   "Text401"
      Tab(3).Control(9)=   "Text404"
      Tab(3).Control(10)=   "Text405"
      Tab(3).ControlCount=   11
      Begin VB.CheckBox Check10 
         Caption         =   "Create Tree/Container Information"
         Height          =   255
         Left            =   -69600
         TabIndex        =   175
         Top             =   1560
         Value           =   1  'Checked
         Width           =   2775
      End
      Begin VB.CheckBox Check9 
         Caption         =   "Use This Custom LDIF File"
         Height          =   255
         Left            =   240
         TabIndex        =   174
         Top             =   4560
         Width           =   2655
      End
      Begin VB.TextBox Text30 
         Height          =   285
         Left            =   3240
         TabIndex        =   173
         Text            =   "C:\Temp\ldif_cust.ldi"
         Top             =   4560
         Width           =   3255
      End
      Begin VB.CommandButton CmdBrowse10 
         Caption         =   "Browse"
         Height          =   375
         Left            =   6600
         TabIndex        =   172
         Top             =   4560
         Width           =   1335
      End
      Begin VB.CommandButton cmdViewCust 
         Caption         =   "View Export File"
         Height          =   375
         Left            =   8160
         TabIndex        =   171
         Top             =   4560
         Width           =   1335
      End
      Begin VB.CheckBox Check8 
         Caption         =   "Add User Set 1 to this Context"
         Height          =   255
         Left            =   -74760
         TabIndex        =   170
         Top             =   3360
         Width           =   3135
      End
      Begin VB.TextBox Text29 
         Height          =   285
         Left            =   -71520
         TabIndex        =   169
         Text            =   ",ou=Provo,o=novell"
         Top             =   3360
         Width           =   4095
      End
      Begin VB.TextBox Text24 
         Height          =   285
         Left            =   -72960
         TabIndex        =   167
         Text            =   "123456789012"
         Top             =   5880
         Width           =   1215
      End
      Begin VB.CommandButton cmdClean 
         Caption         =   "Delete Files"
         Height          =   375
         Left            =   8160
         TabIndex        =   19
         Top             =   4080
         Width           =   1335
      End
      Begin VB.CommandButton cmdViewRice 
         Caption         =   "View ICE Batch"
         Height          =   375
         Left            =   8160
         TabIndex        =   10
         Top             =   2640
         Width           =   1335
      End
      Begin VB.CommandButton CmdBrowse5 
         Caption         =   "Browse"
         Height          =   375
         Left            =   6600
         TabIndex        =   18
         Top             =   4080
         Width           =   1335
      End
      Begin VB.TextBox Text19 
         Height          =   285
         Left            =   3240
         TabIndex        =   17
         Text            =   "C:\Temp\"
         Top             =   4080
         Width           =   3255
      End
      Begin VB.TextBox Text405 
         Height          =   285
         Left            =   -70920
         TabIndex        =   163
         Text            =   "index.html"
         Top             =   2520
         Width           =   2895
      End
      Begin VB.TextBox Text404 
         Height          =   285
         Left            =   -70920
         TabIndex        =   158
         Text            =   "public_html"
         Top             =   2040
         Width           =   2895
      End
      Begin VB.TextBox Text1603 
         Height          =   285
         Left            =   -66840
         TabIndex        =   151
         Text            =   "Level3FullDepthContainer"
         Top             =   3360
         Width           =   1455
      End
      Begin VB.TextBox Text1503 
         Height          =   285
         Left            =   -68400
         TabIndex        =   145
         Text            =   "Users"
         Top             =   3360
         Width           =   1455
      End
      Begin VB.TextBox Text1403 
         Height          =   285
         Left            =   -69960
         TabIndex        =   139
         Text            =   "Users"
         Top             =   3360
         Width           =   1455
      End
      Begin VB.TextBox Text1303 
         Height          =   285
         Left            =   -71520
         TabIndex        =   133
         Text            =   "Users"
         Top             =   3360
         Width           =   1455
      End
      Begin VB.TextBox Text1203 
         Height          =   285
         Left            =   -73080
         TabIndex        =   127
         Text            =   "Users"
         Top             =   3360
         Width           =   1455
      End
      Begin VB.TextBox Text1103 
         Height          =   285
         Left            =   -74640
         TabIndex        =   121
         Text            =   "Users"
         Top             =   3360
         Width           =   1455
      End
      Begin VB.TextBox Text1602 
         Height          =   285
         Left            =   -66840
         TabIndex        =   150
         Text            =   "Level2FullDepthContainer"
         Top             =   2880
         Width           =   1455
      End
      Begin VB.TextBox Text1502 
         Height          =   285
         Left            =   -68400
         TabIndex        =   144
         Text            =   "Marketing"
         Top             =   2880
         Width           =   1455
      End
      Begin VB.TextBox Text1402 
         Height          =   285
         Left            =   -69960
         TabIndex        =   138
         Text            =   "ProtocolEngineering"
         Top             =   2880
         Width           =   1455
      End
      Begin VB.TextBox Text1302 
         Height          =   285
         Left            =   -71520
         TabIndex        =   132
         Text            =   "Engineering"
         Top             =   2880
         Width           =   1455
      End
      Begin VB.TextBox Text1202 
         Height          =   285
         Left            =   -73080
         TabIndex        =   126
         Text            =   "Internationalization"
         Top             =   2880
         Width           =   1455
      End
      Begin VB.TextBox Text1102 
         Height          =   285
         Left            =   -74640
         TabIndex        =   120
         Text            =   "Accounting"
         Top             =   2880
         Width           =   1455
      End
      Begin VB.TextBox Text1601 
         Height          =   285
         Left            =   -66840
         TabIndex        =   149
         Text            =   "Level1DuesseldorfFullDepthContainer"
         Top             =   2400
         Width           =   1455
      End
      Begin VB.TextBox Text1501 
         Height          =   285
         Left            =   -68400
         TabIndex        =   143
         Text            =   "Cambridge"
         Top             =   2400
         Width           =   1455
      End
      Begin VB.TextBox Text1401 
         Height          =   285
         Left            =   -69960
         TabIndex        =   137
         Text            =   "Bangalore"
         Top             =   2400
         Width           =   1455
      End
      Begin VB.TextBox Text1301 
         Height          =   285
         Left            =   -71520
         TabIndex        =   131
         Text            =   "Provo"
         Top             =   2400
         Width           =   1455
      End
      Begin VB.TextBox Text1201 
         Height          =   285
         Left            =   -73080
         TabIndex        =   125
         Text            =   "Dublin"
         Top             =   2400
         Width           =   1455
      End
      Begin VB.TextBox Text1101 
         Height          =   285
         Left            =   -74640
         TabIndex        =   119
         Text            =   "Boston"
         Top             =   2400
         Width           =   1455
      End
      Begin VB.TextBox Text200 
         Height          =   285
         Left            =   -71760
         TabIndex        =   118
         Text            =   "Novell"
         Top             =   1560
         Width           =   1695
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "frmOptions1.frx":0D3A
         Left            =   3240
         List            =   "frmOptions1.frx":0D3C
         Style           =   1  'Simple Combo
         TabIndex        =   1
         Text            =   "636"
         Top             =   840
         Width           =   975
      End
      Begin VB.CheckBox Check7 
         Caption         =   "Select All"
         Height          =   255
         Left            =   -74760
         TabIndex        =   27
         Top             =   720
         Width           =   1095
      End
      Begin VB.CheckBox Check6 
         Caption         =   "Delete"
         Height          =   255
         Left            =   1320
         TabIndex        =   25
         Top             =   5640
         Width           =   975
      End
      Begin VB.CommandButton cmdViewImp 
         Caption         =   "View Import File"
         Height          =   375
         Left            =   8160
         TabIndex        =   16
         Top             =   3600
         Width           =   1335
      End
      Begin VB.CommandButton cmdViewExp 
         Caption         =   "View Export File"
         Height          =   375
         Left            =   8160
         TabIndex        =   13
         Top             =   3120
         Width           =   1335
      End
      Begin VB.CheckBox Check5 
         Caption         =   "Anonymous Bind And/Or Non-SSL"
         Height          =   255
         Left            =   4440
         TabIndex        =   2
         Top             =   840
         Width           =   3135
      End
      Begin VB.TextBox Text14 
         Height          =   285
         Left            =   4440
         TabIndex        =   5
         Top             =   1440
         Width           =   1695
      End
      Begin VB.ComboBox Combo606 
         Height          =   315
         ItemData        =   "frmOptions1.frx":0D3E
         Left            =   -66840
         List            =   "frmOptions1.frx":0D96
         Style           =   1  'Simple Combo
         TabIndex        =   154
         Text            =   "Level6FullDepthContainerFullDepthContainerFullDepth"
         Top             =   4800
         Width           =   1455
      End
      Begin VB.ComboBox Combo605 
         Height          =   315
         ItemData        =   "frmOptions1.frx":0ED4
         Left            =   -66840
         List            =   "frmOptions1.frx":0F2C
         Style           =   1  'Simple Combo
         TabIndex        =   153
         Text            =   "Level5FullDepthContainer"
         Top             =   4320
         Width           =   1455
      End
      Begin VB.ComboBox Combo604 
         Height          =   315
         ItemData        =   "frmOptions1.frx":106A
         Left            =   -66840
         List            =   "frmOptions1.frx":10C2
         Style           =   1  'Simple Combo
         TabIndex        =   152
         Text            =   "Level4FullDepthContainer"
         Top             =   3840
         Width           =   1455
      End
      Begin VB.ComboBox Combo506 
         Height          =   315
         ItemData        =   "frmOptions1.frx":1200
         Left            =   -68400
         List            =   "frmOptions1.frx":1258
         Style           =   1  'Simple Combo
         TabIndex        =   148
         Text            =   "Level6Users"
         Top             =   4800
         Width           =   1455
      End
      Begin VB.ComboBox Combo505 
         Height          =   315
         ItemData        =   "frmOptions1.frx":1396
         Left            =   -68400
         List            =   "frmOptions1.frx":13EE
         Style           =   1  'Simple Combo
         TabIndex        =   147
         Text            =   "Level5Users"
         Top             =   4320
         Width           =   1455
      End
      Begin VB.ComboBox Combo504 
         Height          =   315
         ItemData        =   "frmOptions1.frx":152C
         Left            =   -68400
         List            =   "frmOptions1.frx":1584
         Style           =   1  'Simple Combo
         TabIndex        =   146
         Text            =   "Level4Users"
         Top             =   3840
         Width           =   1455
      End
      Begin VB.ComboBox Combo406 
         Height          =   315
         ItemData        =   "frmOptions1.frx":16C2
         Left            =   -69960
         List            =   "frmOptions1.frx":171A
         Style           =   1  'Simple Combo
         TabIndex        =   142
         Text            =   "Level6Users"
         Top             =   4800
         Width           =   1455
      End
      Begin VB.ComboBox Combo405 
         Height          =   315
         ItemData        =   "frmOptions1.frx":1858
         Left            =   -69960
         List            =   "frmOptions1.frx":18B0
         Style           =   1  'Simple Combo
         TabIndex        =   141
         Text            =   "Level5Users"
         Top             =   4320
         Width           =   1455
      End
      Begin VB.ComboBox Combo404 
         Height          =   315
         ItemData        =   "frmOptions1.frx":19EE
         Left            =   -69960
         List            =   "frmOptions1.frx":1A46
         Style           =   1  'Simple Combo
         TabIndex        =   140
         Text            =   "Level4Users"
         Top             =   3840
         Width           =   1455
      End
      Begin VB.ComboBox Combo306 
         Height          =   315
         ItemData        =   "frmOptions1.frx":1B84
         Left            =   -71520
         List            =   "frmOptions1.frx":1BDC
         Style           =   1  'Simple Combo
         TabIndex        =   136
         Text            =   "Level6Users"
         Top             =   4800
         Width           =   1455
      End
      Begin VB.ComboBox Combo305 
         Height          =   315
         ItemData        =   "frmOptions1.frx":1D1A
         Left            =   -71520
         List            =   "frmOptions1.frx":1D72
         Style           =   1  'Simple Combo
         TabIndex        =   135
         Text            =   "Level5Users"
         Top             =   4320
         Width           =   1455
      End
      Begin VB.ComboBox Combo304 
         Height          =   315
         ItemData        =   "frmOptions1.frx":1EB0
         Left            =   -71520
         List            =   "frmOptions1.frx":1F08
         Style           =   1  'Simple Combo
         TabIndex        =   134
         Text            =   "Level4Users"
         Top             =   3840
         Width           =   1455
      End
      Begin VB.ComboBox Combo206 
         Height          =   315
         ItemData        =   "frmOptions1.frx":2046
         Left            =   -73080
         List            =   "frmOptions1.frx":209E
         Style           =   1  'Simple Combo
         TabIndex        =   130
         Text            =   "Level6Users"
         Top             =   4800
         Width           =   1455
      End
      Begin VB.ComboBox Combo205 
         Height          =   315
         ItemData        =   "frmOptions1.frx":21DC
         Left            =   -73080
         List            =   "frmOptions1.frx":2234
         Style           =   1  'Simple Combo
         TabIndex        =   129
         Text            =   "Level5Users"
         Top             =   4320
         Width           =   1455
      End
      Begin VB.ComboBox Combo204 
         Height          =   315
         ItemData        =   "frmOptions1.frx":2372
         Left            =   -73080
         List            =   "frmOptions1.frx":23CA
         Style           =   1  'Simple Combo
         TabIndex        =   128
         Text            =   "Level4Users"
         Top             =   3840
         Width           =   1455
      End
      Begin VB.ComboBox Combo106 
         Height          =   315
         ItemData        =   "frmOptions1.frx":2508
         Left            =   -74640
         List            =   "frmOptions1.frx":2560
         Style           =   1  'Simple Combo
         TabIndex        =   124
         Text            =   "Level6Users"
         Top             =   4800
         Width           =   1455
      End
      Begin VB.ComboBox Combo105 
         Height          =   315
         ItemData        =   "frmOptions1.frx":269E
         Left            =   -74640
         List            =   "frmOptions1.frx":26F6
         Style           =   1  'Simple Combo
         TabIndex        =   123
         Text            =   "Level5Users"
         Top             =   4320
         Width           =   1455
      End
      Begin VB.ComboBox Combo104 
         Height          =   315
         ItemData        =   "frmOptions1.frx":2834
         Left            =   -74640
         List            =   "frmOptions1.frx":288C
         Style           =   1  'Simple Combo
         TabIndex        =   122
         Text            =   "Level4Users"
         Top             =   3840
         Width           =   1455
      End
      Begin VB.TextBox Text101 
         Height          =   285
         Left            =   -70920
         TabIndex        =   95
         Text            =   "Engineering"
         Top             =   4080
         Width           =   1575
      End
      Begin VB.TextBox Text105 
         Height          =   285
         Left            =   -70920
         TabIndex        =   101
         Text            =   "Accounting"
         Top             =   4440
         Width           =   1575
      End
      Begin VB.TextBox Text109 
         Height          =   285
         Left            =   -70920
         TabIndex        =   106
         Text            =   "ProtocolEngineering"
         Top             =   4800
         Width           =   1575
      End
      Begin VB.TextBox Text113 
         Height          =   285
         Left            =   -70920
         TabIndex        =   110
         Text            =   "Internationalization"
         Top             =   5160
         Width           =   1575
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Create User Home Directories?   Select the ""Home Dir Configuration"" Tab and modify the NDS Home Directory Path."
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   6000
         Width           =   9015
      End
      Begin VB.CommandButton CmdBrowse4 
         Caption         =   "Browse"
         Height          =   375
         Left            =   6600
         TabIndex        =   15
         Top             =   3600
         Width           =   1335
      End
      Begin VB.TextBox Text7 
         Height          =   285
         Left            =   3240
         TabIndex        =   14
         Text            =   "C:\Temp\ldif_imp.ldi"
         Top             =   3600
         Width           =   3255
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Add"
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   5640
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Retrieve Tree Information"
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   5280
         Width           =   2295
      End
      Begin VB.CommandButton CmdBrowse3 
         Caption         =   "Browse"
         Height          =   375
         Left            =   6600
         TabIndex        =   12
         Top             =   3120
         Width           =   1335
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   3240
         TabIndex        =   11
         Text            =   "C:\Temp\ldif_exp.ldi"
         Top             =   3120
         Width           =   3255
      End
      Begin VB.CheckBox Check24 
         Caption         =   "Create User(s)"
         Height          =   255
         Left            =   -67200
         TabIndex        =   31
         Top             =   720
         Width           =   1575
      End
      Begin VB.CheckBox Check23 
         Caption         =   "Create User(s)"
         Height          =   255
         Left            =   -69360
         TabIndex        =   30
         Top             =   720
         Width           =   1575
      End
      Begin VB.CheckBox Check22 
         Caption         =   "Create User(s)"
         Height          =   255
         Left            =   -71520
         TabIndex        =   29
         Top             =   720
         Width           =   1575
      End
      Begin VB.CheckBox Check21 
         Caption         =   "Create User(s)"
         Height          =   255
         Left            =   -73680
         TabIndex        =   28
         Top             =   720
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.ComboBox Combo6 
         Height          =   315
         ItemData        =   "frmOptions1.frx":29CA
         Left            =   8520
         List            =   "frmOptions1.frx":29D4
         TabIndex        =   21
         Text            =   "NDS 8"
         Top             =   2160
         Width           =   975
      End
      Begin VB.ComboBox Combo199 
         Height          =   315
         ItemData        =   "frmOptions1.frx":29E6
         Left            =   -71760
         List            =   "frmOptions1.frx":2CC3
         TabIndex        =   117
         Top             =   960
         Width           =   735
      End
      Begin VB.TextBox Text401 
         Height          =   285
         Left            =   -74760
         TabIndex        =   155
         Text            =   "cn=server1_Vol1,ou=OrganizationalUnit,o=Container#0#\Users_directory"
         Top             =   1080
         Width           =   8175
      End
      Begin VB.TextBox Text403 
         Height          =   285
         Left            =   -70920
         TabIndex        =   156
         Text            =   "C:\Users\"
         Top             =   1560
         Width           =   2895
      End
      Begin VB.CommandButton CmdBrowse6 
         Caption         =   "Browse"
         Height          =   375
         Left            =   -67920
         TabIndex        =   157
         Top             =   1560
         Width           =   1335
      End
      Begin VB.TextBox Text10 
         Height          =   285
         Left            =   -73680
         TabIndex        =   32
         Text            =   "ProvoTestUser"
         Top             =   1320
         Width           =   1935
      End
      Begin VB.TextBox Text11 
         Height          =   285
         Left            =   -73680
         TabIndex        =   33
         Text            =   "John"
         Top             =   1680
         Width           =   1935
      End
      Begin VB.TextBox Text12 
         Height          =   285
         Left            =   -73680
         TabIndex        =   66
         Text            =   "Doe"
         Top             =   2040
         Width           =   1935
      End
      Begin VB.TextBox Text13 
         Height          =   285
         Left            =   -73680
         TabIndex        =   68
         Text            =   "test"
         Top             =   2400
         Width           =   1935
      End
      Begin VB.TextBox Text17 
         Height          =   285
         Left            =   -71520
         TabIndex        =   73
         Text            =   "Malone"
         Top             =   2040
         Width           =   1935
      End
      Begin VB.TextBox Text16 
         Height          =   285
         Left            =   -71520
         TabIndex        =   72
         Text            =   "Jack"
         Top             =   1680
         Width           =   1935
      End
      Begin VB.TextBox Text15 
         Height          =   285
         Left            =   -71520
         TabIndex        =   70
         Text            =   "BostonTestUser"
         Top             =   1320
         Width           =   1935
      End
      Begin VB.TextBox Text18 
         Height          =   285
         Left            =   -71520
         TabIndex        =   75
         Text            =   "test"
         Top             =   2400
         Width           =   1935
      End
      Begin VB.TextBox Text22 
         Height          =   285
         Left            =   -69360
         TabIndex        =   81
         Text            =   "Sarkar"
         Top             =   2040
         Width           =   1935
      End
      Begin VB.TextBox Text21 
         Height          =   285
         Left            =   -69360
         TabIndex        =   79
         Text            =   "Sudarshan"
         Top             =   1680
         Width           =   1935
      End
      Begin VB.TextBox Text20 
         Height          =   285
         Left            =   -69360
         TabIndex        =   77
         Text            =   "BangaloreTestUser"
         Top             =   1320
         Width           =   1935
      End
      Begin VB.TextBox Text23 
         Height          =   285
         Left            =   -69360
         TabIndex        =   82
         Text            =   "test"
         Top             =   2400
         Width           =   1935
      End
      Begin VB.TextBox Text27 
         Height          =   285
         Left            =   -67200
         TabIndex        =   90
         Text            =   "O'Hare"
         Top             =   2040
         Width           =   1935
      End
      Begin VB.TextBox Text26 
         Height          =   285
         Left            =   -67200
         TabIndex        =   88
         Text            =   "Patrick"
         Top             =   1680
         Width           =   1935
      End
      Begin VB.TextBox Text25 
         Height          =   285
         Left            =   -67200
         TabIndex        =   86
         Text            =   "DublinTestUser"
         Top             =   1320
         Width           =   1935
      End
      Begin VB.TextBox Text28 
         Height          =   285
         Left            =   -67200
         TabIndex        =   92
         Text            =   "test"
         Top             =   2400
         Width           =   1935
      End
      Begin VB.ComboBox Combo14 
         Height          =   315
         ItemData        =   "frmOptions1.frx":3093
         Left            =   -71520
         List            =   "frmOptions1.frx":309D
         TabIndex        =   69
         Text            =   "No"
         Top             =   2880
         Width           =   855
      End
      Begin VB.TextBox Text8 
         Height          =   285
         Left            =   -72840
         TabIndex        =   113
         Text            =   "mh.novell.com"
         Top             =   5640
         Width           =   1935
      End
      Begin VB.TextBox Text9 
         Height          =   285
         Left            =   -72840
         TabIndex        =   114
         Text            =   "novell.com"
         Top             =   6000
         Width           =   1935
      End
      Begin VB.TextBox Text112 
         Height          =   285
         Left            =   -72840
         TabIndex        =   109
         Text            =   "Users"
         Top             =   5160
         Width           =   1575
      End
      Begin VB.TextBox Text108 
         Height          =   285
         Left            =   -72840
         TabIndex        =   105
         Text            =   "Users"
         Top             =   4800
         Width           =   1575
      End
      Begin VB.TextBox Text104 
         Height          =   285
         Left            =   -72840
         TabIndex        =   100
         Text            =   "Users"
         Top             =   4440
         Width           =   1575
      End
      Begin VB.TextBox Text100 
         Height          =   285
         Left            =   -72840
         TabIndex        =   94
         Text            =   "Users"
         Top             =   4080
         Width           =   1575
      End
      Begin VB.TextBox Text114 
         Height          =   285
         Left            =   -69000
         TabIndex        =   111
         Text            =   "Dublin"
         Top             =   5160
         Width           =   1575
      End
      Begin VB.TextBox Text110 
         Height          =   285
         Left            =   -69000
         TabIndex        =   107
         Text            =   "Bangalore"
         Top             =   4800
         Width           =   1575
      End
      Begin VB.TextBox Text106 
         Height          =   285
         Left            =   -69000
         TabIndex        =   103
         Text            =   "Boston"
         Top             =   4440
         Width           =   1575
      End
      Begin VB.TextBox Text102 
         Height          =   285
         Left            =   -69000
         TabIndex        =   97
         Text            =   "Provo"
         Top             =   4080
         Width           =   1575
      End
      Begin VB.TextBox Text115 
         Height          =   285
         Left            =   -67080
         TabIndex        =   112
         Text            =   "Novell"
         Top             =   5160
         Width           =   1575
      End
      Begin VB.TextBox Text111 
         Height          =   285
         Left            =   -67080
         TabIndex        =   108
         Text            =   "Novell"
         Top             =   4800
         Width           =   1575
      End
      Begin VB.TextBox Text107 
         Height          =   285
         Left            =   -67080
         TabIndex        =   104
         Text            =   "Novell"
         Top             =   4440
         Width           =   1575
      End
      Begin VB.TextBox Text103 
         Height          =   285
         Left            =   -67080
         TabIndex        =   98
         Text            =   "Novell"
         Top             =   4080
         Width           =   1575
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "frmOptions1.frx":30AA
         Left            =   8520
         List            =   "frmOptions1.frx":30C3
         TabIndex        =   20
         Text            =   "10"
         Top             =   1800
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   240
         TabIndex        =   0
         Text            =   "255.255.255.255"
         Top             =   840
         Width           =   2655
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   240
         TabIndex        =   3
         Text            =   "cn=admin,o=novell"
         Top             =   1440
         Width           =   2655
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   3240
         TabIndex        =   4
         Text            =   "test"
         Top             =   1440
         Width           =   975
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   240
         TabIndex        =   6
         Text            =   "F:\Public\Rootcert.der"
         Top             =   2040
         Width           =   3975
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   3240
         TabIndex        =   8
         Text            =   "C:\temp\RICE.BAT"
         Top             =   2640
         Width           =   3255
      End
      Begin VB.CommandButton CmdBrowse1 
         Caption         =   "Browse"
         Height          =   375
         Left            =   4440
         TabIndex        =   7
         Top             =   2040
         Width           =   1335
      End
      Begin VB.CommandButton CmdBrowse2 
         Caption         =   "Browse"
         Height          =   375
         Left            =   6600
         TabIndex        =   9
         Top             =   2640
         Width           =   1335
      End
      Begin VB.CheckBox Check1 
         Caption         =   "&Stop on Error(s)"
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   4920
         Width           =   1935
      End
      Begin VB.Label Label23 
         Caption         =   $"frmOptions1.frx":30F8
         Height          =   495
         Left            =   -74640
         TabIndex        =   168
         Top             =   6240
         Width           =   8655
      End
      Begin VB.Label Label13 
         Caption         =   "Working Directory"
         Height          =   255
         Left            =   240
         TabIndex        =   165
         Top             =   4080
         Width           =   2895
      End
      Begin VB.Label Label22 
         Caption         =   "User Home File Name"
         Height          =   255
         Left            =   -74760
         TabIndex        =   164
         Top             =   2520
         Width           =   3015
      End
      Begin VB.Label Label21 
         Caption         =   $"frmOptions1.frx":31B0
         Height          =   495
         Left            =   -74760
         TabIndex        =   162
         Top             =   5520
         Width           =   8895
      End
      Begin VB.Label Label19 
         Caption         =   "User Home Sub-Dir Name"
         Height          =   255
         Left            =   -74760
         TabIndex        =   159
         Top             =   2040
         Width           =   3015
      End
      Begin VB.Label Label18 
         Caption         =   "NOTE 2: To Create User Home Directories, select the ""Home Dir Configuration"" Tab and modify the NDS Home Directory Path."
         Height          =   255
         Left            =   240
         TabIndex        =   161
         Top             =   6720
         Width           =   9255
      End
      Begin VB.Label Label12 
         Caption         =   "Base DN for LDAP Search (Typically the Organizational container)"
         Height          =   255
         Left            =   4440
         TabIndex        =   116
         Top             =   1200
         Width           =   4815
      End
      Begin VB.Label Label17 
         Caption         =   $"frmOptions1.frx":328C
         Height          =   375
         Left            =   -74640
         TabIndex        =   115
         Top             =   5400
         Width           =   8655
      End
      Begin VB.Label Label16 
         Caption         =   "Organizational Unit"
         Height          =   255
         Left            =   -70920
         TabIndex        =   102
         Top             =   3720
         Width           =   1455
      End
      Begin VB.Label Label14 
         Caption         =   "LDIF Path and File Name for Import"
         Height          =   255
         Left            =   240
         TabIndex        =   99
         Top             =   3600
         Width           =   2895
      End
      Begin VB.Label Label9 
         Caption         =   "LDIF Path and File Name for Export"
         Height          =   255
         Left            =   240
         TabIndex        =   96
         Top             =   3120
         Width           =   2895
      End
      Begin VB.Label Label11 
         Caption         =   "(Port 389 non-SSL, Port 636 SSL)"
         Height          =   255
         Left            =   4200
         TabIndex        =   93
         Top             =   600
         Width           =   2415
      End
      Begin VB.Label Label10 
         Caption         =   "NOTE 1: Items in BOLD must be modified prior to selecting the ""Run Update"" button."
         Height          =   255
         Left            =   240
         TabIndex        =   160
         Top             =   6360
         Width           =   9015
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
         TabIndex        =   91
         Top             =   2160
         Width           =   2295
      End
      Begin VB.Label Label66 
         Caption         =   "Country Container"
         Height          =   255
         Left            =   -74400
         TabIndex        =   89
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label Label65 
         Caption         =   "Organizational Container"
         Height          =   255
         Left            =   -74400
         TabIndex        =   87
         Top             =   1560
         Width           =   1935
      End
      Begin VB.Label Label64 
         Caption         =   "Organizational Units (OU's)"
         Height          =   255
         Left            =   -71520
         TabIndex        =   85
         Top             =   2040
         Width           =   2295
      End
      Begin VB.Label Label40 
         Caption         =   $"frmOptions1.frx":3330
         Height          =   735
         Left            =   -74760
         TabIndex        =   84
         Top             =   6120
         Width           =   9015
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
         TabIndex        =   83
         Top             =   840
         Width           =   7215
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
         TabIndex        =   80
         Top             =   1560
         Width           =   4335
      End
      Begin VB.Label Label2 
         Caption         =   "Given Name"
         Height          =   255
         Left            =   -74760
         TabIndex        =   78
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "User ID"
         Height          =   255
         Left            =   -74760
         TabIndex        =   76
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Surname"
         Height          =   255
         Left            =   -74760
         TabIndex        =   74
         Top             =   2040
         Width           =   1575
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
         TabIndex        =   71
         Top             =   480
         Width           =   3255
      End
      Begin VB.Label Label100 
         Caption         =   "User Set 1"
         Height          =   255
         Left            =   -73680
         TabIndex        =   67
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label Label5 
         Caption         =   "User Set 1, Context 1"
         Height          =   255
         Left            =   -74760
         TabIndex        =   65
         Top             =   4080
         Width           =   2055
      End
      Begin VB.Label Label6 
         Caption         =   "User Set 2, Context 2"
         Height          =   255
         Left            =   -74760
         TabIndex        =   64
         Top             =   4440
         Width           =   2055
      End
      Begin VB.Label Label7 
         Caption         =   "User Set 3, Context 3"
         Height          =   255
         Left            =   -74760
         TabIndex        =   63
         Top             =   4800
         Width           =   2055
      End
      Begin VB.Label Label4 
         Caption         =   "Password"
         Height          =   255
         Left            =   -74760
         TabIndex        =   62
         Top             =   2400
         Width           =   1095
      End
      Begin VB.Label Label8 
         Caption         =   "User Set 4, Context 4"
         Height          =   255
         Left            =   -74760
         TabIndex        =   61
         Top             =   5160
         Width           =   2055
      End
      Begin VB.Label Label15 
         Caption         =   "User Set 2"
         Height          =   255
         Left            =   -71520
         TabIndex        =   60
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label Label20 
         Caption         =   "User Set 3"
         Height          =   255
         Left            =   -69360
         TabIndex        =   59
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label Label25 
         Caption         =   "User Set 4"
         Height          =   255
         Left            =   -67200
         TabIndex        =   58
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label Label27 
         Caption         =   "Make All Password Match User Name"
         Height          =   255
         Left            =   -74760
         TabIndex        =   57
         Top             =   2880
         Width           =   3255
      End
      Begin VB.Label Label29 
         Caption         =   "Mail Server Name"
         Height          =   255
         Left            =   -74760
         TabIndex        =   56
         Top             =   5640
         Width           =   1455
      End
      Begin VB.Label Label31 
         Caption         =   "Domain Name"
         Height          =   255
         Left            =   -74760
         TabIndex        =   55
         Top             =   6000
         Width           =   1215
      End
      Begin VB.Label Label47 
         Caption         =   "Organizational Unit"
         Height          =   255
         Left            =   -72840
         TabIndex        =   54
         Top             =   3720
         Width           =   1695
      End
      Begin VB.Label Label43 
         Caption         =   "Organizational Unit"
         Height          =   255
         Left            =   -69000
         TabIndex        =   53
         Top             =   3720
         Width           =   1455
      End
      Begin VB.Label Label44 
         Caption         =   "Organization"
         Height          =   255
         Left            =   -67080
         TabIndex        =   52
         Top             =   3720
         Width           =   1695
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
         TabIndex        =   51
         Top             =   1800
         Width           =   2295
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
         TabIndex        =   50
         Top             =   600
         Width           =   2415
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
         TabIndex        =   49
         Top             =   600
         Width           =   1215
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
         TabIndex        =   48
         Top             =   1200
         Width           =   1215
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
         Left            =   3240
         TabIndex        =   47
         Top             =   1200
         Width           =   975
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
         TabIndex        =   46
         Top             =   1800
         Width           =   2055
      End
      Begin VB.Label Label54 
         Caption         =   "Rapid ICE Batch File"
         Height          =   255
         Left            =   240
         TabIndex        =   45
         Top             =   2640
         Width           =   1935
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

Private Sub Check2_Click()
If Check2 = 1 Then Text14.Text = "Novell"
If Check2 = 0 Then Text14.Text = ""
If Check2 = 1 Then Check3 = 0
If Check2 = 1 Then Check6 = 0
End Sub

Private Sub Check3_Click()
If Check3 = 1 Then Check6 = 0
If Check3 = 1 Then Check2 = 0
If Check3 = 1 Then CType = "add"
If Check3 = 0 Then CType = "delete"
End Sub

Private Sub Check5_Click()
If Check5 = 1 Or 0 And Combo1.Text = "389" Then Text4.Text = ""
If Check5 = 1 Then Text4.Text = ""
If Check5 = 1 Then Combo1.Text = "389"

If Check5 = 0 Then Text2.Text = "cn=admin,o=novell"
If Check5 = 0 Then Text3.Text = "test"
If Check5 = 0 Then Text4.Text = "F:\Public\Rootcert.der"
If Check5 = 0 Then Text14.Text = ""
If Check5 = 0 Then Combo1.Text = "636"
End Sub

Private Sub Check6_Click()
If Check6 = 1 Then Check3 = 0
If Check6 = 1 Then Check2 = 0
If Check6 = 1 Then CType = "delete"
If Check6 = 0 Then CType = "add"
End Sub

Private Sub Check7_Click()
If Check7 = 1 Then Check21 = 1
If Check7 = 1 Then Check22 = 1
If Check7 = 1 Then Check23 = 1
If Check7 = 1 Then Check24 = 1
If Check7 = 0 Then Check21 = 0
If Check7 = 0 Then Check22 = 0
If Check7 = 0 Then Check23 = 0
If Check7 = 0 Then Check24 = 0
End Sub

Private Sub Check8_Click()
If Check8 = 1 Then Check7 = 0
If Check8 = 1 Then Check21 = 1
If Check8 = 1 Then Check22 = 0
If Check8 = 1 Then Check23 = 0
If Check8 = 1 Then Check24 = 0
If Check8 = 0 Then Check7 = 0
If Check8 = 0 Then Check21 = 1
If Check8 = 0 Then Check22 = 0
If Check8 = 0 Then Check23 = 0
If Check8 = 0 Then Check24 = 0
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
If Text30.Text = "" Then Text30.Text = "C:\Temp\ldif_cust.ldi"
End Sub

Private Sub CmdBrowse2_Click()
cmnDialog.FileName = Text5.Text
cmnDialog.ShowOpen
Text5.Text = cmnDialog.FileName
If Text5.Text = "" Then Text5.Text = "C:\temp\RICE.BAT"
End Sub

Private Sub CmdBrowse3_Click()
cmnDialog.FileName = Text6.Text
cmnDialog.ShowOpen
Text6.Text = cmnDialog.FileName
If Text6.Text = "" Then Text6.Text = "C:\Temp\ldif_file.ldi"
End Sub

Private Sub CmdBrowse4_Click()
cmnDialog.FileName = Text7.Text
cmnDialog.ShowOpen
Text7.Text = cmnDialog.FileName
If Text7.Text = "" Then Text7.Text = "C:\Temp\ldif_imp.ldi"
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
   
   Dim IPAddress As String
   Dim Port As String
   Dim Username As String
   Dim Password As String
   Dim Rootcert As String
   Dim RICEBatch As String
   Dim fileout As String
   Dim filein As String
   Dim beginval As Long
   Dim endval As Long
   Dim ver As String
   Dim servername As String
   Dim domainname As String
   Dim passwordm As String
   Dim Estop As String
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
      
   IPAddress = Text1.Text
   Username = Text2.Text
   Password = Text3.Text
   BaseDN = Text14.Text
   Rootcert = Text4.Text
   RICEBatch = Text5.Text
   fileout = Text6.Text
   filein = Text7.Text
   servername = Text8.Text
   domainname = Text9.Text
   beginval = 1
   Port = Combo1.Text
   endval = Val(Combo2.Text)
   passwordm = Combo14.Text
   ver = Combo6.Text
   Estop = ""
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
   
   If Check3 = 1 Then CType = "add"
   If Check3 = 0 Then CType = "delete"
   If Check6 = 1 Then CType = "delete"
   If Check6 = 0 Then CType = "add"
                    
           Dim fso, msg
           Set fso = CreateObject("Scripting.FileSystemObject")
           If (fso.FolderExists(WPath)) Then GoTo CStep180 Else MkDir WPath  'Make new directory or folder.
CStep180:
          ChDrive WPath 'Changes the current drive.
          ChDir WPath 'Changes the current directory or folder.
                              
        If Check2 = 1 Then GoTo Step2
                 
   Open fileout For Output As #1
   
If Check8 = 1 Then GoTo CCStep1

If Check9 = 1 Then GoTo CustLDIFStep1 'If Custom LDIF file is selected bypass the file creation.
   
'Start Tree Information
If Check3 = 0 Then GoTo Step3

Step3:
'End Tree Information
        
   If endval = "1" Then GoTo Step4 'If endval is 1 this goes around the Progress Bar Error
   
   ProgressBar1.Min = beginval
   ProgressBar1.Max = endval
   ProgressBar1.Value = ProgressBar1.Min
   ProgressBar1.Visible = True
   
Step4:
   
If CType = "Delete" Then GoTo Step41
      Print #1, "#This file generated by Tree Builder Version 3." 'File Header
If Check10 = 0 Then GoTo BypassTree
      'Start Add Tree Information

      If CCont = "" Then GoTo Step22
      If CType = "delete" Then GoTo Step22

      Print #1,
      Print #1, "dn: c=" + CCont
      Print #1, "changetype: " + CType
      Print #1, "objectClass: top"
      Print #1, "objectClass: country"
      If ver = "NDS 8" Then Print #1, "c: " + CCont
      If ver = "NDS 8" Then Print #1, "objectClass: top"
      Print #1,

Step22:

If Check3 = 0 Then GoTo Step20

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
      Print #1,
      
Step20:
   'End Add Tree Information
   
CCStep1:
BypassTree:
   For i = beginval To endval
   ProgressBar1.Value = i
      
      'Start User data
      If Check21 = 0 Then GoTo Step10 'If Create User(s) is not checked, bypass user creation for that set
      Print #1,
      If Check8 = 1 Then Print #1, "dn: cn=" + userid1 + Format(i) + CustCont
      If Check8 = 1 Then GoTo CCStep3
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

      If Check22 = 0 Then GoTo Step11 'If Create User(s) is not checked, bypass user creation for that set
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
      
      If Check23 = 0 Then GoTo Step12 'If Create User(s) is not checked, bypass user creation for that set
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
            
      If Check24 = 0 Then GoTo Step13 'If Create User(s) is not checked, bypass user creation for that set
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
      
   Next i
   
   'End Tree Information
        
If endval = "1" Then GoTo Step41 'If endval is 1 this goes around the Progress Bar Error
   
   ProgressBar1.Min = beginval
   ProgressBar1.Max = endval
   ProgressBar1.Value = ProgressBar1.Min
   ProgressBar1.Visible = True
   
If Check8 = 1 Then GoTo CCStep2
   
'Start Delete Tree Information

Step41:

   If Check6 = 0 Then GoTo Step40
If Check10 = "0" Then GoTo BypassTree1

      Print #1,
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
            
If CCont = "" Then GoTo Step40

      Print #1,
      Print #1, "dn: c=" + CCont
      Print #1, "changetype: delete"
      Print #1,

Step40:

   'End Delete Tree Information
   
CCStep2:
   
   ProgressBar1.Visible = True
   
   Close #1

Step2:
CustLDIFStep1:
BypassTree1:

If Check9 = 1 Then fileout = CustLDIF

      Open RICEBatch For Output As #2
      'Print #2, "path %PATH%;C:\Program Files\TreeBldr3\"
      Print #2, "path %PATH%;" + App.Path
      Print #2, "del ice.Log"
      If Check1 = 0 Then Estop = " -c" Else Estop = ""
      If Check2 = 1 Then GoTo Step300
      If Check5 = 1 Then Username = ""
      If Check5 = 1 Then Password = ""
      If Username = "" And Password = "" Then Port = "389"
      
      If Username = "" And Password = "" Then Print #2, "ice -S LDIF -f " + fileout + Estop + " -D LDAP -s " + IPAddress + " -p " + Port
      If Username <> "" And Password <> "" And Port = "636" Then Print #2, "ice -S LDIF -f " + fileout + Estop + " -D LDAP -s " + IPAddress + " -p " + Port + " -d " + Username + " -w " + Password + " -L " + Rootcert
      If Username <> "" And Password <> "" And Port = "389" Then Print #2, "ice -S LDIF -f " + fileout + Estop + " -D LDAP -s " + IPAddress + " -p " + Port + " -d " + Username + " -w " + Password
      
Step300:
      If Check2 = 0 Then GoTo Step301
      If Check5 = 1 Then Username = ""
      If Check5 = 1 Then Password = ""
      If Username = "" And Password = "" Then Port = "389"
      If Username = "" And Password = "" Then Print #2, "ice -S LDAP -s " + IPAddress + " -p " + Port + " -b o=" + BaseDN + " -c sub" + " -D LDIF -f " + filein
      If Username <> "" And Password <> "" Then Print #2, "ice -S LDAP -s " + IPAddress + " -p " + Port + " -d " + Username + " -w " + Password + " -b o=" + BaseDN + " -c sub -L " + Rootcert + " -D LDIF -f " + filein
      Print #2,
      
Step301:

   Close #2

'Call Shell(RICEBatch, 1)   'Call the ICE Batch file.
'Start Display Code

   ICE.Show 0
 
   Dim JobToDo As String
      JobToDo = RICEBatch
      DoEvents: Sleep 100
      Shell32Bit JobToDo

'End Display Code

'Start User Home Directory Creation Section

If Check4 = 0 Then GoTo Step161
If homedir = "cn=server1_Vol1,ou=OrganizationalUnit,o=Container#0#\Users_directory" Then GoTo Step161

          'Dim fso, msg
          'Set fso = CreateObject("Scripting.FileSystemObject")
          If (fso.FolderExists(MyPath)) Then GoTo Step1000 Else MkDir MyPath 'Make new directory or folder.

Step1000:
          
          ChDrive MyPath
          ChDir MyPath 'Changes the current directory or folder.

If endval = "1" Then GoTo Step150
   
   ProgressBar1.Min = beginval
   ProgressBar1.Max = endval
   ProgressBar1.Value = ProgressBar1.Min
   ProgressBar1.Visible = True
   
Step150:
   
   For i = beginval To endval
   ProgressBar1.Value = i
      
   If userid1 = "" Then GoTo Step151
   If Check21 = 0 Then GoTo Step151
      If (fso.FolderExists(MyPath + userid1 + Format(i))) Then GoTo CStep1 Else MkDir MyPath + userid1 + Format(i)  'Make new directory or folder.
CStep1:
      ChDir MyPath + userid1 + Format(i)
      If (fso.FolderExists(UHomeDir)) Then GoTo CStep2 Else MkDir UHomeDir  'Make new directory or folder.
CStep2:
      ChDir UHomeDir
      Open UHomeFile For Output As #2
      Print #2, userid1 + Format(i)
      ChDir "cd ..\.."
      Close #2
            
Step151:
   If userid2 = "" Then GoTo Step152
   If Check22 = 0 Then GoTo Step152
      If (fso.FolderExists(MyPath + userid2 + Format(i))) Then GoTo CStep3 Else MkDir MyPath + userid2 + Format(i) 'Make new directory or folder.
CStep3:
      ChDir MyPath + userid2 + Format(i)
      If (fso.FolderExists(UHomeDir)) Then GoTo CStep4 Else MkDir UHomeDir  'Make new directory or folder.
CStep4:
      ChDir UHomeDir
      Open UHomeFile For Output As #2
      Print #2, userid2 + Format(i)
      ChDir "cd ..\.."
      Close #2
         
Step152:
   If userid3 = "" Then GoTo Step153
   If Check23 = 0 Then GoTo Step153
      If (fso.FolderExists(MyPath + userid3 + Format(i))) Then GoTo CStep5 Else MkDir MyPath + userid3 + Format(i)  'Make new directory or folder.
CStep5:
      ChDir MyPath + userid3 + Format(i)
      If (fso.FolderExists(UHomeDir)) Then GoTo CStep6 Else MkDir UHomeDir  'Make new directory or folder.
CStep6:
      ChDir UHomeDir
      Open UHomeFile For Output As #2
      Print #2, userid3 + Format(i)
      ChDir "cd ..\.."
      Close #2
      
Step153:
   If userid4 = "" Then GoTo Step154
   If Check24 = 0 Then GoTo Step154
      If (fso.FolderExists(MyPath + userid4 + Format(i))) Then GoTo CStep7 Else MkDir MyPath + userid4 + Format(i)  'Make new directory or folder.
CStep7:
      ChDir MyPath + userid4 + Format(i)
      If (fso.FolderExists(UHomeDir)) Then GoTo CStep8 Else MkDir UHomeDir  'Make new directory or folder.
CStep8:
      ChDir UHomeDir
      Open UHomeFile For Output As #2
      Print #2, userid4 + Format(i)
      ChDir "cd ..\.."
      Close #2
      
Step154:
      
If endval = "1" Then GoTo Step160

ProgressBar1.Visible = True

   Next i
      ChDir WPath
      'ChDir "..\..\..\..\..\..\.."
         
Step160:

     ChDrive "C:"
     ChDir App.Path

Step161:
'End User Home Directory Creation Section

End Sub

Private Sub CmdWrite_Click()
   
   Dim IPAddress As String
   Dim Port As String
   Dim Username As String
   Dim Password As String
   Dim Rootcert As String
   Dim RICEBatch As String
   Dim fileout As String
   Dim filein As String
   Dim beginval As Long
   Dim endval As Long
   Dim ver As String
   Dim servername As String
   Dim domainname As String
   Dim passwordm As String
   Dim Estop As String
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
      
   IPAddress = Text1.Text
   Username = Text2.Text
   Password = Text3.Text
   BaseDN = Text14.Text
   Rootcert = Text4.Text
   RICEBatch = Text5.Text
   fileout = Text6.Text
   filein = Text7.Text
   servername = Text8.Text
   domainname = Text9.Text
   beginval = 1
   Port = Combo1.Text
   endval = Val(Combo2.Text)
   passwordm = Combo14.Text
   ver = Combo6.Text
   Estop = ""
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
   
   If Check3 = 1 Then CType = "add"
   If Check3 = 0 Then CType = "delete"
   If Check6 = 1 Then CType = "delete"
   If Check6 = 0 Then CType = "add"
                    
           Dim fso, msg
           Set fso = CreateObject("Scripting.FileSystemObject")
           If (fso.FolderExists(WPath)) Then GoTo CStep180 Else MkDir WPath  'Make new directory or folder.
CStep180:
          ChDrive WPath 'Changes the current drive.
          ChDir WPath 'Changes the current directory or folder.
                              
        If Check2 = 1 Then GoTo Step2
                 
   Open fileout For Output As #1
   
If Check8 = 1 Then GoTo CCStep1

If Check9 = 1 Then GoTo CustLDIFStep1 'If Custom LDIF file is selected bypass the file creation.
   
'Start Tree Information
If Check3 = 0 Then GoTo Step3

Step3:
'End Tree Information
        
   If endval = "1" Then GoTo Step4 'If endval is 1 this goes around the Progress Bar Error
   
   ProgressBar1.Min = beginval
   ProgressBar1.Max = endval
   ProgressBar1.Value = ProgressBar1.Min
   ProgressBar1.Visible = True
   
Step4:
   
If CType = "Delete" Then GoTo Step41
      Print #1, "#This file generated by Tree Builder Version 3." 'File Header
If Check10 = 0 Then GoTo BypassTree
      'Start Add Tree Information

      If CCont = "" Then GoTo Step22
      If CType = "delete" Then GoTo Step22

      Print #1,
      Print #1, "dn: c=" + CCont
      Print #1, "changetype: " + CType
      Print #1, "objectClass: top"
      Print #1, "objectClass: country"
      If ver = "NDS 8" Then Print #1, "c: " + CCont
      If ver = "NDS 8" Then Print #1, "objectClass: top"
      Print #1,

Step22:

If Check3 = 0 Then GoTo Step20

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
      Print #1,
      
Step20:
   'End Add Tree Information
   
CCStep1:
BypassTree:
   For i = beginval To endval
   ProgressBar1.Value = i
      
      'Start User data
      If Check21 = 0 Then GoTo Step10 'If Create User(s) is not checked, bypass user creation for that set
      Print #1,
      If Check8 = 1 Then Print #1, "dn: cn=" + userid1 + Format(i) + CustCont
      If Check8 = 1 Then GoTo CCStep3
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

      If Check22 = 0 Then GoTo Step11 'If Create User(s) is not checked, bypass user creation for that set
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
      
      If Check23 = 0 Then GoTo Step12 'If Create User(s) is not checked, bypass user creation for that set
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
            
      If Check24 = 0 Then GoTo Step13 'If Create User(s) is not checked, bypass user creation for that set
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
      
   Next i
   
   'End Tree Information
        
If endval = "1" Then GoTo Step41 'If endval is 1 this goes around the Progress Bar Error
   
   ProgressBar1.Min = beginval
   ProgressBar1.Max = endval
   ProgressBar1.Value = ProgressBar1.Min
   ProgressBar1.Visible = True
   
If Check8 = 1 Then GoTo CCStep2
   
'Start Delete Tree Information

Step41:

   If Check6 = 0 Then GoTo Step40
If Check10 = 0 Then GoTo BypassTree1

      Print #1,
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
            
If CCont = "" Then GoTo Step40

      Print #1,
      Print #1, "dn: c=" + CCont
      Print #1, "changetype: delete"
      Print #1,

Step40:

   'End Delete Tree Information
   
CCStep2:
   
   ProgressBar1.Visible = True
   
   Close #1

Step2:
CustLDIFStep1:
BypassTree1:

If Check9 = 1 Then fileout = CustLDIF

      Open RICEBatch For Output As #2
      'Print #2, "path %PATH%;C:\Program Files\TreeBldr3\"
      Print #2, "path %PATH%;" + App.Path
      Print #2, "del ice.Log"
      If Check1 = 0 Then Estop = " -c" Else Estop = ""
      If Check2 = 1 Then GoTo Step300
      If Check5 = 1 Then Username = ""
      If Check5 = 1 Then Password = ""
      If Username = "" And Password = "" Then Port = "389"
      
      If Username = "" And Password = "" Then Print #2, "ice -S LDIF -f " + fileout + Estop + " -D LDAP -s " + IPAddress + " -p " + Port
      If Username <> "" And Password <> "" And Port = "636" Then Print #2, "ice -S LDIF -f " + fileout + Estop + " -D LDAP -s " + IPAddress + " -p " + Port + " -d " + Username + " -w " + Password + " -L " + Rootcert
      If Username <> "" And Password <> "" And Port = "389" Then Print #2, "ice -S LDIF -f " + fileout + Estop + " -D LDAP -s " + IPAddress + " -p " + Port + " -d " + Username + " -w " + Password
      
Step300:
      If Check2 = 0 Then GoTo Step301
      If Check5 = 1 Then Username = ""
      If Check5 = 1 Then Password = ""
      If Username = "" And Password = "" Then Port = "389"
      If Username = "" And Password = "" Then Print #2, "ice -S LDAP -s " + IPAddress + " -p " + Port + " -b o=" + BaseDN + " -c sub" + " -D LDIF -f " + filein
      If Username <> "" And Password <> "" Then Print #2, "ice -S LDAP -s " + IPAddress + " -p " + Port + " -d " + Username + " -w " + Password + " -b o=" + BaseDN + " -c sub -L " + Rootcert + " -D LDIF -f " + filein
      Print #2,
      
Step301:

   Close #2

'Call Shell(RICEBatch, 1)   'Call the ICE Batch file.
'Start Display Code

'   ICE.Show 0
 
'   Dim JobToDo As String
'      JobToDo = RICEBatch
'      DoEvents: Sleep 100
'      Shell32Bit JobToDo

'End Display Code

'Start User Home Directory Creation Section

If Check4 = 0 Then GoTo Step161
If homedir = "cn=server1_Vol1,ou=OrganizationalUnit,o=Container#0#\Users_directory" Then GoTo Step161

          'Dim fso, msg
          'Set fso = CreateObject("Scripting.FileSystemObject")
          If (fso.FolderExists(MyPath)) Then GoTo Step1000 Else MkDir MyPath 'Make new directory or folder.

Step1000:
          
          ChDrive MyPath
          ChDir MyPath 'Changes the current directory or folder.

If endval = "1" Then GoTo Step150
   
   ProgressBar1.Min = beginval
   ProgressBar1.Max = endval
   ProgressBar1.Value = ProgressBar1.Min
   ProgressBar1.Visible = True
   
Step150:
   
   For i = beginval To endval
   ProgressBar1.Value = i
      
   If userid1 = "" Then GoTo Step151
   If Check21 = 0 Then GoTo Step151
      If (fso.FolderExists(MyPath + userid1 + Format(i))) Then GoTo CStep1 Else MkDir MyPath + userid1 + Format(i)  'Make new directory or folder.
CStep1:
      ChDir MyPath + userid1 + Format(i)
      If (fso.FolderExists(UHomeDir)) Then GoTo CStep2 Else MkDir UHomeDir  'Make new directory or folder.
CStep2:
      ChDir UHomeDir
      Open UHomeFile For Output As #2
      Print #2, userid1 + Format(i)
      ChDir "cd ..\.."
      Close #2
            
Step151:
   If userid2 = "" Then GoTo Step152
   If Check22 = 0 Then GoTo Step152
      If (fso.FolderExists(MyPath + userid2 + Format(i))) Then GoTo CStep3 Else MkDir MyPath + userid2 + Format(i) 'Make new directory or folder.
CStep3:
      ChDir MyPath + userid2 + Format(i)
      If (fso.FolderExists(UHomeDir)) Then GoTo CStep4 Else MkDir UHomeDir  'Make new directory or folder.
CStep4:
      ChDir UHomeDir
      Open UHomeFile For Output As #2
      Print #2, userid2 + Format(i)
      ChDir "cd ..\.."
      Close #2
         
Step152:
   If userid3 = "" Then GoTo Step153
   If Check23 = 0 Then GoTo Step153
      If (fso.FolderExists(MyPath + userid3 + Format(i))) Then GoTo CStep5 Else MkDir MyPath + userid3 + Format(i)  'Make new directory or folder.
CStep5:
      ChDir MyPath + userid3 + Format(i)
      If (fso.FolderExists(UHomeDir)) Then GoTo CStep6 Else MkDir UHomeDir  'Make new directory or folder.
CStep6:
      ChDir UHomeDir
      Open UHomeFile For Output As #2
      Print #2, userid3 + Format(i)
      ChDir "cd ..\.."
      Close #2
      
Step153:
   If userid4 = "" Then GoTo Step154
   If Check24 = 0 Then GoTo Step154
      If (fso.FolderExists(MyPath + userid4 + Format(i))) Then GoTo CStep7 Else MkDir MyPath + userid4 + Format(i)  'Make new directory or folder.
CStep7:
      ChDir MyPath + userid4 + Format(i)
      If (fso.FolderExists(UHomeDir)) Then GoTo CStep8 Else MkDir UHomeDir  'Make new directory or folder.
CStep8:
      ChDir UHomeDir
      Open UHomeFile For Output As #2
      Print #2, userid4 + Format(i)
      ChDir "cd ..\.."
      Close #2
      
Step154:
      
If endval = "1" Then GoTo Step160

ProgressBar1.Visible = True

   Next i
      ChDir WPath
      'ChDir "..\..\..\..\..\..\.."
         
Step160:

     ChDrive "C:"
     ChDir App.Path

Step161:
'End User Home Directory Creation Section

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
