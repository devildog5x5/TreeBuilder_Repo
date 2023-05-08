VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form RICE 
   Caption         =   "RICE.EXE version 1.0"
   ClientHeight    =   3780
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8415
   LinkTopic       =   "Form1"
   ScaleHeight     =   3780
   ScaleWidth      =   8415
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text8 
      Height          =   285
      Left            =   4920
      TabIndex        =   19
      Text            =   "C:\Temp\RapidICE.BAT"
      Top             =   2160
      Width           =   2895
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Create"
      Height          =   375
      Left            =   5760
      TabIndex        =   18
      Top             =   2640
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Browse"
      Height          =   375
      Left            =   3480
      TabIndex        =   16
      Top             =   3120
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Browse"
      Height          =   375
      Left            =   3480
      TabIndex        =   15
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Browse"
      Height          =   375
      Left            =   3480
      TabIndex        =   14
      Top             =   1320
      Width           =   975
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   360
      TabIndex        =   12
      Text            =   "C:\Temp\RICE_log.txt"
      Top             =   3240
      Width           =   2895
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   360
      TabIndex        =   10
      Text            =   "F:\public\Rootcert.der"
      Top             =   2400
      Width           =   2895
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   360
      TabIndex        =   8
      Text            =   "C:\Temp\ldif_file.txt"
      Top             =   1440
      Width           =   2895
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   5640
      TabIndex        =   6
      Text            =   "test"
      Top             =   600
      Width           =   1095
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   3840
      TabIndex        =   4
      Text            =   "cn=admin,o=novell"
      Top             =   600
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   2760
      TabIndex        =   2
      Text            =   "636"
      Top             =   600
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   360
      TabIndex        =   0
      Text            =   "111.222.333.444"
      Top             =   600
      Width           =   2175
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   10200
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label8 
      Caption         =   "Build ICE Batch File"
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
      Left            =   5400
      TabIndex        =   17
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Label Label7 
      Caption         =   "Log File Name"
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
      Left            =   360
      TabIndex        =   13
      Top             =   2880
      Width           =   1815
   End
   Begin VB.Label Label6 
      Caption         =   "Location of RootCert.der File"
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
      Left            =   360
      TabIndex        =   11
      Top             =   2040
      Width           =   2655
   End
   Begin VB.Label Label5 
      Caption         =   "LDIF File Name"
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
      Left            =   360
      TabIndex        =   9
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label Label4 
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
      Left            =   5640
      TabIndex        =   7
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Username"
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
      Left            =   3840
      TabIndex        =   5
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Port"
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
      Left            =   2760
      TabIndex        =   3
      Top             =   240
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Server IP Address"
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
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   1815
   End
End
Attribute VB_Name = "RICE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
   CommonDialog1.ShowOpen
   Text5.Text = CommonDialog1.filename
End Sub
Private Sub Command2_Click()
   CommonDialog1.ShowOpen
   Text6.Text = CommonDialog1.filename
End Sub
Private Sub Command3_Click()
   CommonDialog1.ShowOpen
   Text7.Text = CommonDialog1.filename
End Sub

Private Sub Command4_Click()
   
   Dim IPAddress As String
   Dim Port As String
   Dim Username As String
   Dim Password As String
   Dim filename As String
   Dim LOGFile As String
   Dim RICEBatch As String
         
   IPAddress = Text1.Text
   Port = Text2.Text
   Username = Text3.Text
   Password = Text4.Text
   LDIFfile = Text5.Text
   RootCert = Text6.Text
   LOGFile = Text7.Text
   RICEBatch = Text8.Text
                     
'Start Main ICE Batch File Creation

   Open RICEBatch For Output As #1
      
      Print #1, "path %PATH%;C:\Program Files\RapidICE"
      Print #1, "ice -S LDIF -f " + LDIFfile + " -D LDAP -s " + IPAddress + " -p " + Port + " -d " + Username + " -w " + Password + " -L " + RootCert; " > " + LOGFile
      Print #1,
      
   Close #1
End Sub
