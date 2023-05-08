VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   Caption         =   "User Set Builder"
   ClientHeight    =   9900
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13215
   BeginProperty Font 
      Name            =   "Roman"
      Size            =   9.75
      Charset         =   255
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   9900
   ScaleWidth      =   13215
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   "LDIF File Generator"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13215
      Begin VB.ComboBox Combo13 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "Form1.frx":0000
         Left            =   8040
         List            =   "Form1.frx":000A
         TabIndex        =   126
         Text            =   "No"
         Top             =   7080
         Width           =   735
      End
      Begin VB.ComboBox Combo12 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "Form1.frx":0017
         Left            =   4920
         List            =   "Form1.frx":0021
         TabIndex        =   125
         Text            =   "No"
         Top             =   7080
         Width           =   1095
      End
      Begin VB.CommandButton Command13 
         Caption         =   "View ICE.LOG"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   10920
         TabIndex        =   123
         Top             =   8040
         Width           =   1935
      End
      Begin VB.CommandButton Command12 
         Caption         =   "Browse"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8880
         TabIndex        =   122
         Top             =   7080
         Width           =   975
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Run HomeDir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   10920
         TabIndex        =   61
         Top             =   8520
         Width           =   1935
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Browse"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   11880
         TabIndex        =   55
         Top             =   7080
         Width           =   975
      End
      Begin VB.CommandButton Command11 
         Caption         =   "Browse"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9840
         TabIndex        =   60
         Top             =   8040
         Width           =   975
      End
      Begin VB.TextBox Text52 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5640
         TabIndex        =   59
         Text            =   "C:\Users\"
         Top             =   8160
         Width           =   3975
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1200
         TabIndex        =   4
         Text            =   "English"
         Top             =   1800
         Width           =   1935
      End
      Begin VB.TextBox Text9 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4320
         TabIndex        =   14
         Text            =   "English"
         Top             =   1800
         Width           =   1935
      End
      Begin VB.TextBox Text14 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7320
         TabIndex        =   24
         Text            =   "English"
         Top             =   1800
         Width           =   1935
      End
      Begin VB.TextBox Text19 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   10440
         TabIndex        =   34
         Text            =   "English"
         Top             =   1800
         Width           =   1935
      End
      Begin VB.TextBox Text74 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6120
         TabIndex        =   53
         Text            =   "C:\temp\RICE.BAT"
         Top             =   7080
         Width           =   1815
      End
      Begin VB.TextBox Text76 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   9960
         TabIndex        =   54
         Text            =   "F:\Public\Rootcert.der"
         Top             =   7080
         Width           =   1815
      End
      Begin VB.TextBox Text73 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3960
         TabIndex        =   52
         Text            =   "test"
         Top             =   7080
         Width           =   855
      End
      Begin VB.TextBox Text72 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2160
         TabIndex        =   51
         Text            =   "cn=admin,o=novell"
         Top             =   7080
         Width           =   1695
      End
      Begin VB.TextBox Text71 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1560
         TabIndex        =   50
         Text            =   "636"
         Top             =   7080
         Width           =   495
      End
      Begin VB.TextBox Text70 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   49
         Text            =   "111.222.333.444"
         Top             =   7080
         Width           =   1335
      End
      Begin VB.TextBox Text105 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6840
         TabIndex        =   8
         Text            =   "Novell"
         Top             =   3840
         Width           =   1935
      End
      Begin VB.TextBox Text110 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6840
         TabIndex        =   18
         Text            =   "Novell"
         Top             =   4200
         Width           =   1935
      End
      Begin VB.TextBox Text115 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6840
         TabIndex        =   28
         Text            =   "Novell"
         Top             =   4560
         Width           =   1935
      End
      Begin VB.TextBox Text120 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6840
         TabIndex        =   38
         Text            =   "Novell"
         Top             =   4920
         Width           =   1935
      End
      Begin VB.TextBox Text104 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4440
         TabIndex        =   9
         Text            =   "Provo"
         Top             =   3840
         Width           =   1935
      End
      Begin VB.TextBox Text109 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4440
         TabIndex        =   19
         Text            =   "Boston"
         Top             =   4200
         Width           =   1935
      End
      Begin VB.TextBox Text114 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4440
         TabIndex        =   29
         Text            =   "Bangalore"
         Top             =   4560
         Width           =   1935
      End
      Begin VB.TextBox Text119 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4440
         TabIndex        =   39
         Text            =   "Dublin"
         Top             =   4920
         Width           =   1935
      End
      Begin VB.TextBox Text103 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2040
         TabIndex        =   10
         Text            =   "Users"
         Top             =   3840
         Width           =   1935
      End
      Begin VB.TextBox Text108 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2040
         TabIndex        =   20
         Text            =   "Users"
         Top             =   4200
         Width           =   1935
      End
      Begin VB.TextBox Text113 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2040
         TabIndex        =   30
         Text            =   "Users"
         Top             =   4560
         Width           =   1935
      End
      Begin VB.TextBox Text118 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2040
         TabIndex        =   40
         Text            =   "Users"
         Top             =   4920
         Width           =   1935
      End
      Begin VB.TextBox Text63 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   46
         Text            =   "cn=server1_Vol1,ou=OrganizationalUnit,o=Container#0#\Users_directory"
         Top             =   6480
         Width           =   12255
      End
      Begin VB.TextBox Text51 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   57
         Text            =   "C:\Users\hdir.bat"
         Top             =   8160
         Width           =   3975
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Browse"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4320
         TabIndex        =   58
         Top             =   8040
         Width           =   1095
      End
      Begin VB.ComboBox Combo11 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "Form1.frx":002E
         Left            =   11400
         List            =   "Form1.frx":0038
         TabIndex        =   37
         Text            =   "Yes"
         Top             =   4920
         Width           =   975
      End
      Begin VB.ComboBox Combo10 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "Form1.frx":0045
         Left            =   11400
         List            =   "Form1.frx":004F
         TabIndex        =   27
         Text            =   "Yes"
         Top             =   4560
         Width           =   975
      End
      Begin VB.ComboBox Combo8 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "Form1.frx":005C
         Left            =   11400
         List            =   "Form1.frx":0066
         TabIndex        =   7
         Text            =   "Yes"
         Top             =   3840
         Width           =   975
      End
      Begin VB.ComboBox Combo9 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "Form1.frx":0073
         Left            =   11400
         List            =   "Form1.frx":007D
         TabIndex        =   17
         Text            =   "Yes"
         Top             =   4200
         Width           =   975
      End
      Begin VB.ComboBox Combo7 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "Form1.frx":008A
         Left            =   10800
         List            =   "Form1.frx":0094
         TabIndex        =   36
         Text            =   "add"
         Top             =   2520
         Width           =   1575
      End
      Begin VB.ComboBox Combo6 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "Form1.frx":00A5
         Left            =   7680
         List            =   "Form1.frx":00AF
         TabIndex        =   26
         Text            =   "add"
         Top             =   2520
         Width           =   1575
      End
      Begin VB.ComboBox Combo5 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "Form1.frx":00C0
         Left            =   4680
         List            =   "Form1.frx":00CA
         TabIndex        =   16
         Text            =   "add"
         Top             =   2520
         Width           =   1575
      End
      Begin VB.ComboBox Combo4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "Form1.frx":00DB
         Left            =   1560
         List            =   "Form1.frx":00E5
         TabIndex        =   6
         Text            =   "add"
         Top             =   2520
         Width           =   1575
      End
      Begin VB.TextBox Text62 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7320
         TabIndex        =   45
         Text            =   "novell.com"
         Top             =   6000
         Width           =   1935
      End
      Begin VB.TextBox Text61 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7320
         TabIndex        =   44
         Text            =   "mh.novell.com"
         Top             =   5640
         Width           =   1935
      End
      Begin VB.ComboBox Combo3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "Form1.frx":00F6
         Left            =   7320
         List            =   "Form1.frx":0100
         TabIndex        =   43
         Text            =   "NetWare 6"
         Top             =   5280
         Width           =   1935
      End
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "Form1.frx":011C
         Left            =   3960
         List            =   "Form1.frx":0135
         TabIndex        =   42
         Text            =   "1"
         Top             =   5640
         Width           =   1215
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "Form1.frx":016A
         Left            =   3960
         List            =   "Form1.frx":0174
         TabIndex        =   41
         Text            =   "No"
         Top             =   5280
         Width           =   735
      End
      Begin VB.TextBox Text20 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   10440
         TabIndex        =   35
         Text            =   "test"
         Top             =   2160
         Width           =   1935
      End
      Begin VB.TextBox Text16 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   10440
         TabIndex        =   31
         Text            =   "DublinTestUser"
         Top             =   720
         Width           =   1935
      End
      Begin VB.TextBox Text17 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   10440
         TabIndex        =   32
         Text            =   "Patrick"
         Top             =   1080
         Width           =   1935
      End
      Begin VB.TextBox Text18 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   10440
         TabIndex        =   33
         Text            =   "O'Hare"
         Top             =   1440
         Width           =   1935
      End
      Begin VB.TextBox Text15 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7320
         TabIndex        =   25
         Text            =   "test"
         Top             =   2160
         Width           =   1935
      End
      Begin VB.TextBox Text11 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7320
         TabIndex        =   21
         Text            =   "BangaloreTestUser"
         Top             =   720
         Width           =   1935
      End
      Begin VB.TextBox Text12 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7320
         TabIndex        =   22
         Text            =   "Sudarshan"
         Top             =   1080
         Width           =   1935
      End
      Begin VB.TextBox Text13 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7320
         TabIndex        =   23
         Text            =   "Sarkar"
         Top             =   1440
         Width           =   1935
      End
      Begin VB.TextBox Text10 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4320
         TabIndex        =   15
         Text            =   "test"
         Top             =   2160
         Width           =   1935
      End
      Begin VB.TextBox Text6 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4320
         TabIndex        =   11
         Text            =   "BostonTestUser"
         Top             =   720
         Width           =   1935
      End
      Begin VB.TextBox Text7 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4320
         TabIndex        =   12
         Text            =   "Jack"
         Top             =   1080
         Width           =   1935
      End
      Begin VB.TextBox Text8 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4320
         TabIndex        =   13
         Text            =   "Malone"
         Top             =   1440
         Width           =   1935
      End
      Begin VB.TextBox Text5 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1200
         TabIndex        =   5
         Text            =   "test"
         Top             =   2160
         Width           =   1935
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1200
         TabIndex        =   3
         Text            =   "Doe"
         Top             =   1440
         Width           =   1935
      End
      Begin ComctlLib.ProgressBar ProgressBar1 
         Height          =   255
         Left            =   240
         TabIndex        =   66
         Top             =   7560
         Width           =   10575
         _ExtentX        =   18653
         _ExtentY        =   450
         _Version        =   327682
         Appearance      =   1
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   12360
         Top             =   120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Browse"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   11640
         TabIndex        =   48
         Top             =   6000
         Width           =   1095
      End
      Begin VB.TextBox Text50 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   9600
         TabIndex        =   47
         Text            =   "C:\Temp\ldif_file.ldi"
         Top             =   6000
         Width           =   1935
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1200
         TabIndex        =   2
         Text            =   "John"
         Top             =   1080
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1200
         TabIndex        =   1
         Text            =   "ProvoTestUser"
         Top             =   720
         Width           =   1935
      End
      Begin VB.CommandButton Command2 
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
         Height          =   375
         Left            =   10920
         TabIndex        =   56
         Top             =   7560
         Width           =   1935
      End
      Begin ComctlLib.ProgressBar ProgressBar2 
         Height          =   255
         Left            =   240
         TabIndex        =   104
         Top             =   8640
         Width           =   10575
         _ExtentX        =   18653
         _ExtentY        =   450
         _Version        =   327682
         Appearance      =   1
      End
      Begin VB.Label Label46 
         Caption         =   "Get LDAP"
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
         Left            =   8040
         TabIndex        =   127
         Top             =   6840
         Width           =   975
      End
      Begin VB.Label Label42 
         Caption         =   "Stop on Error"
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
         Left            =   4920
         TabIndex        =   124
         Top             =   6840
         Width           =   1215
      End
      Begin VB.Label Label45 
         Caption         =   "Enter Directory to Create User Directories"
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
         TabIndex        =   121
         Top             =   7920
         Width           =   4095
      End
      Begin VB.Label Label13 
         Caption         =   "Enter Path and File Name"
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
         Left            =   120
         TabIndex        =   120
         Top             =   7920
         Width           =   2295
      End
      Begin VB.Label Label67 
         Caption         =   "Language"
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
         Left            =   120
         TabIndex        =   119
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label Label63 
         Caption         =   "Language"
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
         TabIndex        =   118
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label Label60 
         Caption         =   "Language"
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
         Left            =   6240
         TabIndex        =   117
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label Label57 
         Caption         =   "Language"
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
         Left            =   9360
         TabIndex        =   116
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label Label54 
         Caption         =   "Rapid ICE Batch File"
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
         Left            =   6120
         TabIndex        =   115
         Top             =   6840
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
         Left            =   9960
         TabIndex        =   114
         Top             =   6840
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
         Left            =   3960
         TabIndex        =   113
         Top             =   6840
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
         Left            =   2160
         TabIndex        =   112
         Top             =   6840
         Width           =   1215
      End
      Begin VB.Label Label49 
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
         Left            =   1560
         TabIndex        =   111
         Top             =   6840
         Width           =   495
      End
      Begin VB.Label Label48 
         Caption         =   "IP Address"
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
         Left            =   120
         TabIndex        =   110
         Top             =   6840
         Width           =   1215
      End
      Begin VB.Label Label44 
         Caption         =   "Organization"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6840
         TabIndex        =   109
         Top             =   3480
         Width           =   1695
      End
      Begin VB.Label Label43 
         Caption         =   "Organizational Unit"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4440
         TabIndex        =   108
         Top             =   3480
         Width           =   1695
      End
      Begin VB.Label Label47 
         Caption         =   "Organizational Unit"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2040
         TabIndex        =   107
         Top             =   3480
         Width           =   1695
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
         Left            =   120
         TabIndex        =   106
         Top             =   6240
         Width           =   7215
      End
      Begin VB.Label Label40 
         Caption         =   "NOTE:This should be the location on disk where home directories are to be created"
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
         Left            =   120
         TabIndex        =   105
         Top             =   9000
         Width           =   7215
      End
      Begin VB.Label Label39 
         Caption         =   "Create Context?"
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
         Left            =   9360
         TabIndex        =   103
         Top             =   4920
         Width           =   1575
      End
      Begin VB.Label Label38 
         Caption         =   "Create Context?"
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
         Left            =   9360
         TabIndex        =   102
         Top             =   4560
         Width           =   1575
      End
      Begin VB.Label Label37 
         Caption         =   "Create Context?"
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
         Left            =   9360
         TabIndex        =   101
         Top             =   3840
         Width           =   1575
      End
      Begin VB.Label Label36 
         Caption         =   "Create Context?"
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
         Left            =   9360
         TabIndex        =   100
         Top             =   4200
         Width           =   1575
      End
      Begin VB.Label Label35 
         Caption         =   "Change Type"
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
         Left            =   9360
         TabIndex        =   99
         Top             =   2520
         Width           =   1455
      End
      Begin VB.Label Label34 
         Caption         =   "Change Type"
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
         Left            =   6240
         TabIndex        =   98
         Top             =   2520
         Width           =   1455
      End
      Begin VB.Label Label33 
         Caption         =   "Change Type"
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
         TabIndex        =   97
         Top             =   2520
         Width           =   1455
      End
      Begin VB.Label Label32 
         Caption         =   "Change Type"
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
         Left            =   120
         TabIndex        =   96
         Top             =   2520
         Width           =   1455
      End
      Begin VB.Label Label31 
         Caption         =   "Domain Name"
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
         TabIndex        =   95
         Top             =   6000
         Width           =   1215
      End
      Begin VB.Label Label30 
         Caption         =   "Select NetWare OS"
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
         TabIndex        =   94
         Top             =   5280
         Width           =   2055
      End
      Begin VB.Label Label29 
         Caption         =   "Mail Server Name"
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
         TabIndex        =   93
         Top             =   5640
         Width           =   2055
      End
      Begin VB.Label Label28 
         Caption         =   "Number of Users to Create in each Context"
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
         Left            =   120
         TabIndex        =   92
         Top             =   5640
         Width           =   3735
      End
      Begin VB.Label Label27 
         Caption         =   "Make Password Match User Name"
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
         Left            =   120
         TabIndex        =   91
         Top             =   5280
         Width           =   3015
      End
      Begin VB.Label Label26 
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
         Left            =   9360
         TabIndex        =   90
         Top             =   2160
         Width           =   1575
      End
      Begin VB.Label Label25 
         Caption         =   "User Set 4"
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
         Left            =   10440
         TabIndex        =   89
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label24 
         Caption         =   "Given Name"
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
         Left            =   9360
         TabIndex        =   88
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label23 
         Caption         =   "User ID"
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
         Left            =   9360
         TabIndex        =   87
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label22 
         Caption         =   "Surname"
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
         Left            =   9360
         TabIndex        =   86
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label21 
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
         Left            =   6240
         TabIndex        =   85
         Top             =   2160
         Width           =   1575
      End
      Begin VB.Label Label20 
         Caption         =   "User Set 3"
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
         Left            =   7320
         TabIndex        =   84
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label19 
         Caption         =   "Given Name"
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
         Left            =   6240
         TabIndex        =   83
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label18 
         Caption         =   "User ID"
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
         Left            =   6240
         TabIndex        =   82
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label17 
         Caption         =   "Surname"
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
         Left            =   6240
         TabIndex        =   81
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label16 
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
         TabIndex        =   80
         Top             =   2160
         Width           =   1575
      End
      Begin VB.Label Label15 
         Caption         =   "User Set 2"
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
         Left            =   4320
         TabIndex        =   79
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label14 
         Caption         =   "Given Name"
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
         TabIndex        =   78
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label11 
         Caption         =   "User ID"
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
         TabIndex        =   77
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label10 
         Caption         =   "Surname"
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
         TabIndex        =   76
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label8 
         Caption         =   "User Set 4, Context 4"
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
         Left            =   120
         TabIndex        =   75
         Top             =   4920
         Width           =   2055
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
         Left            =   120
         TabIndex        =   74
         Top             =   2160
         Width           =   1575
      End
      Begin VB.Label Label7 
         Caption         =   "User Set 3, Context 3"
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
         Left            =   120
         TabIndex        =   73
         Top             =   4560
         Width           =   2055
      End
      Begin VB.Label Label6 
         Caption         =   "User Set 2, Context 2"
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
         Left            =   120
         TabIndex        =   72
         Top             =   4200
         Width           =   2055
      End
      Begin VB.Label Label5 
         Caption         =   "User Set 1, Context 1"
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
         Left            =   120
         TabIndex        =   71
         Top             =   3840
         Width           =   2055
      End
      Begin VB.Label Label100 
         Caption         =   "User Set 1"
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
         Left            =   1200
         TabIndex        =   70
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label55 
         Caption         =   $"Form1.frx":0181
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   69
         Top             =   3000
         Width           =   12255
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
         Left            =   4680
         TabIndex        =   68
         Top             =   240
         Width           =   3255
      End
      Begin VB.Label Label3 
         Caption         =   "Surname"
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
         Left            =   120
         TabIndex        =   65
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label9 
         Caption         =   "Enter LDIF Path and File Name"
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
         Left            =   9600
         TabIndex        =   64
         Top             =   5640
         Width           =   2655
      End
      Begin VB.Label Label1 
         Caption         =   "User ID"
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
         Left            =   120
         TabIndex        =   63
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Given Name"
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
         Left            =   120
         TabIndex        =   62
         Top             =   1080
         Width           =   1575
      End
   End
   Begin VB.Label Label12 
      Caption         =   "User Context"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   67
      Top             =   3960
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Text50.Text = "C:\Temp\ldif_file.ldi"
CommonDialog1.ShowOpen
Text50.Text = CommonDialog1.filename
End Sub

Private Sub Command2_Click()

   Dim userid1 As String
   'Dim userid2 As String
   'Dim userid3 As String
   'Dim userid4 As String
   Dim givenname1 As String
   'Dim givenname2 As String
   'Dim givenname3 As String
   'Dim givenname4 As String
   Dim surname1 As String
   'Dim surname2 As String
   'Dim surname3 As String
   'Dim surname4 As String
   Dim password1 As String
   'Dim password2 As String
   'Dim password3 As String
   'Dim password4 As String
   Dim Lang1 As String
   'Dim Lang2 As String
   'Dim Lang3 As String
   'Dim Lang4 As String
   Dim Org1 As String
   'Dim Org2 As String
   'Dim Org3 As String
   'Dim Org4 As String
   Dim OrgUnit1 As String
   'Dim OrgUnit2 As String
   'Dim OrgUnit3 As String
   'Dim OrgUnit4 As String
   Dim OrgUnit11 As String
   'Dim OrgUnit22 As String
   'Dim OrgUnit33 As String
   'Dim OrgUnit44 As String
   Dim OrgUnit111 As String
   'Dim OrgUnit222 As String
   'Dim OrgUnit333 As String
   'Dim OrgUnit444 As String
   Dim OrgUnit1111 As String
   'Dim OrgUnit2222 As String
   'Dim OrgUnit3333 As String
   'Dim OrgUnit4444 As String
   Dim cType1 As String
   'Dim cType2 As String
   'Dim cType3 As String
   'Dim cType4 As String
   Dim servername As String
   Dim domainname As String
   Dim filename As String
   Dim homedir As String
   Dim beginval As Long
   Dim endval As Long
   Dim ver As String
   Dim passwordm As String
   Dim context1 As String
   'Dim context2 As String
   'Dim context3 As String
   'Dim context4 As String
   Dim IPAddress As String
   Dim Port As String
   Dim Username As String
   Dim Password As String
   Dim bfilename As String
   Dim RICEBatch As String
   Dim Estop As String
   Dim GetLDAP As String
            
   userid1 = Text1.Text
   givenname1 = Text2.Text
   surname1 = Text3.Text
   Lang1 = Text4.Text
   password1 = Text5.Text
   cType1 = Combo4.Text
   Org1 = Text105.Text
   OrgUnit1 = Text104.Text
   OrgUnit11 = Text103.Text
   'userid2 = Text6.Text
   'givenname2 = Text7.Text
   'surname2 = Text8.Text
   'Lang2 = Text9.Text
   'password2 = Text10.Text
   'cType2 = Combo5.Text
   'Org2 = Text110.Text
   'OrgUnit2 = Text109.Text
   'OrgUnit22 = Text108.Text
   'userid3 = Text11.Text
   'givenname3 = Text12.Text
   'surname3 = Text13.Text
   'Lang3 = Text14.Text
   'password3 = Text15.Text
   'cType3 = Combo6.Text
   'Org3 = Text115.Text
   'OrgUnit3 = Text114.Text
   'OrgUnit33 = Text113.Text
   'userid4 = Text16.Text
   'givenname4 = Text17.Text
   'surname4 = Text18.Text
   'Lang4 = Text19.Text
   'password4 = Text20.Text
   'cType4 = Combo7.Text
   'Org4 = Text120.Text
   'OrgUnit4 = Text119.Text
   'OrgUnit44 = Text118.Text
   servername = Text61.Text
   domainname = Text62.Text
   filename = Text50.Text
   homedir = Text63.Text
   IPAddress = Text70.Text
   Port = Text71.Text
   Username = Text72.Text
   Password = Text73.Text
   'LDIFfile = Text75.Text
   Rootcert = Text76.Text
   RICEBatch = Text74.Text
   beginval = 1
   endval = Val(Combo2.Text)
   passwordm = Combo1.Text
   ContC1 = Combo8.Text
   'ContC2 = Combo9.Text
   'ContC3 = Combo10.Text
   'ContC4 = Combo11.Text
   ver = Combo3.Text
   Estop = Combo12.Text
   GetLDAP = Combo12.Text
   
   If GetLDAP = "Yes" Then GoTo Step299
   
'Start Main LDIF

   Open filename For Output As #1

'*** Start Organizational and Organizational Unit data

If cType1 = "delete" Then GoTo Step240
If cType1 = "modify" Then GoTo Step241
If ContC1 = "Yes" And cType1 = "add" Then Print #1,
If ContC1 = "Yes" And cType1 = "add" Then Print #1, "dn: o=" + Org1
If ContC1 = "Yes" And cType1 = "add" Then Print #1, "changetype: " + cType1
If ContC1 = "Yes" And cType1 = "add" Then Print #1, "objectClass: organization"
If ContC1 = "Yes" And cType1 = "add" Then Print #1, "objectClass: ndsLoginProperties"
If ContC1 = "Yes" And cType1 = "add" Then Print #1, "objectClass: ndsContainerLoginProperties"
If ContC1 = "Yes" And cType1 = "add" Then Print #1, "objectClass: Top"
If ContC1 = "Yes" And cType1 = "add" Then Print #1,
If ContC1 = "Yes" And cType1 = "add" Then Print #1, "dn: ou=" + OrgUnit1 + ",o=" + Org1
If ContC1 = "Yes" And cType1 = "add" Then Print #1, "changetype: " + cType1
If ContC1 = "Yes" And cType1 = "add" Then Print #1, "ou: " + OrgUnit1
If ContC1 = "Yes" And cType1 = "add" Then Print #1, "objectClass: organizationalUnit"
If ContC1 = "Yes" And cType1 = "add" Then Print #1, "objectClass: ndsLoginProperties"
If ContC1 = "Yes" And cType1 = "add" Then Print #1, "objectClass: Top"
If ContC1 = "Yes" And cType1 = "add" Then Print #1,
If ContC1 = "Yes" And cType1 = "add" Then Print #1, "dn: ou=" + OrgUnit11 + ",ou=" + OrgUnit1 + ",o=" + Org1
If ContC1 = "Yes" And cType1 = "add" Then Print #1, "changetype: " + cType1
If ContC1 = "Yes" And cType1 = "add" Then Print #1, "ou: " + OrgUnit11
If ContC1 = "Yes" And cType1 = "add" Then Print #1, "objectClass: organizationalUnit"
If ContC1 = "Yes" And cType1 = "add" Then Print #1, "objectClass: ndsLoginProperties"
If ContC1 = "Yes" And cType1 = "add" Then Print #1, "objectClass: Top"
Step240:
Step241:

'If cType2 = "delete" Then GoTo Step242
'If cType2 = "modify" Then GoTo Step243
'If ContC2 = "Yes" And cType2 = "add" Then Print #1,
'If ContC2 = "Yes" And cType2 = "add" Then Print #1, "dn: o=" + Org2
'If ContC2 = "Yes" And cType2 = "add" Then Print #1, "changetype: " + cType2
'If ContC1 = "Yes" And cType1 = "add" Then Print #1, "objectClass: organization"
'If ContC1 = "Yes" And cType1 = "add" Then Print #1, "objectClass: ndsLoginProperties"
'If ContC1 = "Yes" And cType1 = "add" Then Print #1, "objectClass: ndsContainerLoginProperties"
'If ContC1 = "Yes" And cType1 = "add" Then Print #1, "objectClass: Top"
'If ContC2 = "Yes" And cType2 = "add" Then Print #1,
'If ContC2 = "Yes" And cType2 = "add" Then Print #1, "dn: ou=" + OrgUnit2 + ",o=" + Org2
'If ContC2 = "Yes" And cType2 = "add" Then Print #1, "changetype: " + cType2
'If ContC2 = "Yes" And cType2 = "add" Then Print #1, "ou: " + OrgUnit2
'If ContC2 = "Yes" And cType2 = "add" Then Print #1, "objectClass: organizationalUnit"
'If ContC2 = "Yes" And cType2 = "add" Then Print #1, "objectClass: ndsLoginProperties"
'If ContC2 = "Yes" And cType2 = "add" Then Print #1, "objectClass: Top"
'If ContC2 = "Yes" And cType2 = "add" Then Print #1,
'If ContC2 = "Yes" And cType2 = "add" Then Print #1, "dn: ou=" + OrgUnit22 + ",ou=" + OrgUnit2 + ",o=" + Org2
'If ContC2 = "Yes" And cType2 = "add" Then Print #1, "changetype: " + cType2
'If ContC2 = "Yes" And cType2 = "add" Then Print #1, "ou: " + OrgUnit22
'If ContC2 = "Yes" And cType2 = "add" Then Print #1, "objectClass: organizationalUnit"
'If ContC2 = "Yes" And cType2 = "add" Then Print #1, "objectClass: ndsLoginProperties"
'If ContC2 = "Yes" And cType2 = "add" Then Print #1, "objectClass: Top"
'Step242:
'Step243:

'If cType3 = "delete" Then GoTo Step244
'If cType3 = "modify" Then GoTo Step245
'If ContC3 = "Yes" And cType3 = "add" Then Print #1,
'If ContC3 = "Yes" And cType3 = "add" Then Print #1, "dn: o=" + Org3
'If ContC3 = "Yes" And cType3 = "add" Then Print #1, "changetype: " + cType3
'If ContC1 = "Yes" And cType1 = "add" Then Print #1, "objectClass: organization"
'If ContC1 = "Yes" And cType1 = "add" Then Print #1, "objectClass: ndsLoginProperties"
'If ContC1 = "Yes" And cType1 = "add" Then Print #1, "objectClass: ndsContainerLoginProperties"
'If ContC1 = "Yes" And cType1 = "add" Then Print #1, "objectClass: Top"
'If ContC3 = "Yes" And cType3 = "add" Then Print #1,
'If ContC3 = "Yes" And cType3 = "add" Then Print #1, "dn: ou=" + OrgUnit3 + ",o=" + Org3
'If ContC3 = "Yes" And cType3 = "add" Then Print #1, "changetype: " + cType3
'If ContC3 = "Yes" And cType3 = "add" Then Print #1, "ou: " + OrgUnit3
'If ContC3 = "Yes" And cType3 = "add" Then Print #1, "objectClass: organizationalUnit"
'If ContC3 = "Yes" And cType3 = "add" Then Print #1, "objectClass: ndsLoginProperties"
'If ContC3 = "Yes" And cType3 = "add" Then Print #1, "objectClass: Top"
'If ContC3 = "Yes" And cType3 = "add" Then Print #1,
'If ContC3 = "Yes" And cType3 = "add" Then Print #1, "dn: ou=" + OrgUnit33 + ",ou=" + OrgUnit3 + ",o=" + Org3
'If ContC3 = "Yes" And cType3 = "add" Then Print #1, "changetype: " + cType3
'If ContC3 = "Yes" And cType3 = "add" Then Print #1, "ou: " + OrgUnit33
'If ContC3 = "Yes" And cType3 = "add" Then Print #1, "objectClass: organizationalUnit"
'If ContC3 = "Yes" And cType3 = "add" Then Print #1, "objectClass: ndsLoginProperties"
'If ContC3 = "Yes" And cType3 = "add" Then Print #1, "objectClass: Top"
'Step244:
'Step245:

'If cType4 = "delete" Then GoTo Step246
'If cType4 = "modify" Then GoTo Step247
'If ContC4 = "Yes" And cType4 = "add" Then Print #1,
'If ContC4 = "Yes" And cType4 = "add" Then Print #1, "dn: o=" + Org4
'If ContC4 = "Yes" And cType4 = "add" Then Print #1, "changetype: " + cType4
'If ContC1 = "Yes" And cType1 = "add" Then Print #1, "objectClass: organization"
'If ContC1 = "Yes" And cType1 = "add" Then Print #1, "objectClass: ndsLoginProperties"
'If ContC1 = "Yes" And cType1 = "add" Then Print #1, "objectClass: ndsContainerLoginProperties"
'If ContC1 = "Yes" And cType1 = "add" Then Print #1, "objectClass: Top"
'If ContC4 = "Yes" And cType4 = "add" Then Print #1,
'If ContC4 = "Yes" And cType4 = "add" Then Print #1, "dn: ou=" + OrgUnit4 + ",o=" + Org4
'If ContC4 = "Yes" And cType4 = "add" Then Print #1, "changetype: " + cType4
'If ContC4 = "Yes" And cType4 = "add" Then Print #1, "ou: " + OrgUnit4
'If ContC4 = "Yes" And cType4 = "add" Then Print #1, "objectClass: organizationalUnit"
'If ContC4 = "Yes" And cType4 = "add" Then Print #1, "objectClass: ndsLoginProperties"
'If ContC4 = "Yes" And cType4 = "add" Then Print #1, "objectClass: Top"
'If ContC4 = "Yes" And cType4 = "add" Then Print #1,
'If ContC4 = "Yes" And cType4 = "add" Then Print #1, "dn: ou=" + OrgUnit44 + ",ou=" + OrgUnit4 + ",o=" + Org4
'If ContC4 = "Yes" And cType4 = "add" Then Print #1, "changetype: " + cType4
'If ContC4 = "Yes" And cType4 = "add" Then Print #1, "ou: " + OrgUnit44
'If ContC4 = "Yes" And cType4 = "add" Then Print #1, "objectClass: organizationalUnit"
'If ContC4 = "Yes" And cType4 = "add" Then Print #1, "objectClass: ndsLoginProperties"
'If ContC4 = "Yes" And cType4 = "add" Then Print #1, "objectClass: Top"
'Step246:
'Step247:

'*** End Organizational and Organizational Unit data
        
If endval = "1" Then GoTo Step11
    
         ProgressBar1.Min = beginval
         ProgressBar1.Max = endval
         ProgressBar1.Value = ProgressBar1.Min
         ProgressBar1.Visible = False

Step11:

         For i = beginval To endval
         ProgressBar1.Value = i

'*** Start User data

      If userid1 = "" Then GoTo Step1 'If userid is blank, bypass user creation for that set
      Print #1,
      Print #1, "dn: cn=" + userid1 + Format(i) + ",ou=" + OrgUnit11 + ",ou=" + OrgUnit1 + ",o=" + Org1
      Print #1, "changetype: " + cType1
      If cType1 = "delete" Then GoTo Step1
      'Print #1, "nsLicensedFor: mail"
      'Print #1, "mailHost: " + servername
      'Print #1, "mailDeliveryOption: mailbox"
      Print #1, "mail: " + userid1 + Format(i) + "@" + domainname
      Print #1, "uid: " + userid1 + Format(i)
      Print #1, "givenName: " + givenname1 + Format(i)
      'Print #1, "Language: " + Lang1
      Print #1, "sn: " + surname1 + Format(i)
      Print #1, "objectclass: inetOrgPerson"
      Print #1, "objectclass: organizationalPerson"
      Print #1, "objectclass: person"
      Print #1, "objectclass: ndsLoginProperties"
      Print #1, "objectclass: top"
      'Print #1, "objectclass: mailRecipient"
      'Print #1, "objectclass: nsLicenseUser"
      If passwordm = "Yes" Then Print #1, "userpassword: " + userid1 + Format(i) Else Print #1, "userpassword: " + password1
      
      If homedir = "cn=server1_Vol1,ou=OrganizationalUnit,o=Container#0#\Users_directory" Then GoTo Step1
      If ver = "NetWare 6" Then Print #1, "ndsHomeDirectory: " + homedir + "\" + userid1 + Format(i)
      If ver = "NetWare 5.1" Then Print #1, "homeDirectory: " + homedir + "\" + userid1 + Format(i)
      
'Step1:

'      If userid2 = "" Then GoTo Step2 'If userid is blank, bypass user creation for that set
'      Print #1,
'      Print #1, "dn: cn=" + userid2 + Format(i) + ",ou=" + OrgUnit22 + ",ou=" + OrgUnit2 + ",o=" + Org2
'      Print #1, "changetype: " + cType2
'      If cType2 = "delete" Then GoTo Step2
'      'Print #1, "nsLicensedFor: mail"
'      'Print #1, "mailHost: " + servername
'      'Print #1, "mailDeliveryOption: mailbox"
'      Print #1, "mail: " + userid2 + Format(i) + "@" + domainname
'      Print #1, "uid: " + userid2 + Format(i)
'      Print #1, "givenName: " + givenname2 + Format(i)
'      'Print #1, "Language: " + Lang2
'      Print #1, "sn: " + surname2 + Format(i)
'      Print #1, "objectclass: inetOrgPerson"
'      Print #1, "objectclass: organizationalPerson"
'      Print #1, "objectclass: person"
'      Print #1, "objectclass: ndsLoginProperties"
'      Print #1, "objectclass: top"
'      'Print #1, "objectclass: mailRecipient"
'      'Print #1, "objectclass: nsLicenseUser"
'      If passwordm = "Yes" Then Print #1, "userpassword: " + userid2 + Format(i) Else Print #1, "userpassword: " + password2
      
'      If homedir = "cn=server1_Vol1,ou=OrganizationalUnit,o=Container#0#\Users_directory" Then GoTo Step2
'      If ver = "NetWare 6" Then Print #1, "ndsHomeDirectory: " + homedir + "\" + userid2 + Format(i)
'      If ver = "NetWare 5.1" Then Print #1, "homeDirectory: " + homedir + "\" + userid2 + Format(i)
        
'Step2:
        
'      If userid3 = "" Then GoTo Step4 'If userid is blank, bypass user creation for that set
'      Print #1,
'      Print #1, "dn: cn=" + userid3 + Format(i) + ",ou=" + OrgUnit33 + ",ou=" + OrgUnit3 + ",o=" + Org3
'      Print #1, "changetype: " + cType3
'      If cType2 = "delete" Then GoTo Step4
'      'Print #1, "nsLicensedFor: mail"
'      'Print #1, "mailHost: " + servername
'      'Print #1, "mailDeliveryOption: mailbox"
'      Print #1, "mail: " + userid3 + Format(i) + "@" + domainname
'      Print #1, "uid: " + userid3 + Format(i)
'      Print #1, "givenName: " + givenname3 + Format(i)
'      'Print #1, "Language: " + Lang3
'      Print #1, "sn: " + surname3 + Format(i)
'      Print #1, "objectclass: inetOrgPerson"
'      Print #1, "objectclass: organizationalPerson"
'      Print #1, "objectclass: person"
'      Print #1, "objectclass: ndsLoginProperties"
'      Print #1, "objectclass: top"
'      'Print #1, "objectclass: mailRecipient"
'      'Print #1, "objectclass: nsLicenseUser"
'      If passwordm = "Yes" Then Print #1, "userpassword: " + userid3 + Format(i) Else Print #1, "userpassword: " + password3

'      If homedir = "cn=server1_Vol1,ou=OrganizationalUnit,o=Container#0#\Users_directory" Then GoTo Step4
'      If ver = "NetWare 6" Then Print #1, "ndsHomeDirectory: " + homedir + "\" + userid3 + Format(i)
'      If ver = "NetWare 5.1" Then Print #1, "homeDirectory: " + homedir + "\" + userid3 + Format(i)
      
'Step4:
     
'      If userid4 = "" Then GoTo Step6 'If userid is blank, bypass user creation for that set
'      Print #1,
'      Print #1, "dn: cn=" + userid4 + Format(i) + ",ou=" + OrgUnit44 + ",ou=" + OrgUnit4 + ",o=" + Org4
'      Print #1, "changetype: " + cType4
'      If cType2 = "delete" Then GoTo Step6
'      'Print #1, "nsLicensedFor: mail"
'      'Print #1, "mailHost: " + servername
'      'Print #1, "mailDeliveryOption: mailbox"
'      Print #1, "mail: " + userid4 + Format(i) + "@" + domainname
'      Print #1, "uid: " + userid4 + Format(i)
'      Print #1, "givenName: " + givenname4 + Format(i)
'      'Print #1, "Language: " + Lang4
'      Print #1, "sn: " + surname4 + Format(i)
'      Print #1, "objectclass: inetOrgPerson"
'      Print #1, "objectclass: organizationalPerson"
'      Print #1, "objectclass: person"
'      Print #1, "objectclass: ndsLoginProperties"
'      Print #1, "objectclass: top"
'      'Print #1, "objectclass: mailRecipient"
'      'Print #1, "objectclass: nsLicenseUser"
'      If passwordm = "Yes" Then Print #1, "userpassword: " + userid4 + Format(i) Else Print #1, "userpassword: " + password4
      
'      If homedir = "cn=server1_Vol1,ou=OrganizationalUnit,o=Container#0#\Users_directory" Then GoTo Step7
'      If ver = "NetWare 6" Then Print #1, "ndsHomeDirectory: " + homedir + "\" + userid4 + Format(i)
'      If ver = "NetWare 5.1" Then Print #1, "homeDirectory: " + homedir + "\" + userid4 + Format(i)
      
'Step6:
Step7:

'*** End User data

      If endval = "1" Then GoTo Step10
      
ProgressBar1.Visible = True

Next i

Step10:


If ContC1 = "Yes" And cType1 = "delete" Then Print #1,
If ContC1 = "Yes" And cType1 = "delete" Then Print #1, "dn: " + "ou=" + OrgUnit11 + ",ou=" + OrgUnit1 + ",o=" + Org1
If ContC1 = "Yes" And cType1 = "delete" Then Print #1, "changetype: " + cType1

If ContC1 = "Yes" And cType1 = "delete" Then Print #1,
If ContC1 = "Yes" And cType1 = "delete" Then Print #1, "dn: " + "ou=" + OrgUnit1 + ",o=" + Org1
If ContC1 = "Yes" And cType1 = "delete" Then Print #1, "changetype: " + cType1

'If ContC2 = "Yes" And cType2 = "delete" Then Print #1,
'If ContC2 = "Yes" And cType2 = "delete" Then Print #1, "dn: " + "ou=" + OrgUnit22 + ",ou=" + OrgUnit2 + ",o=" + Org2
'If ContC2 = "Yes" And cType2 = "delete" Then Print #1, "changetype: " + cType2

'If ContC2 = "Yes" And cType2 = "delete" Then Print #1,
'If ContC2 = "Yes" And cType2 = "delete" Then Print #1, "dn: " + "ou=" + OrgUnit2 + ",o=" + Org2
'If ContC2 = "Yes" And cType2 = "delete" Then Print #1, "changetype: " + cType2

'If ContC3 = "Yes" And cType3 = "delete" Then Print #1,
'If ContC3 = "Yes" And cType3 = "delete" Then Print #1, "dn: " + "ou=" + OrgUnit33 + ",ou=" + OrgUnit3 + ",o=" + Org3
'If ContC3 = "Yes" And cType3 = "delete" Then Print #1, "changetype: " + cType3

'If ContC3 = "Yes" And cType3 = "delete" Then Print #1,
'If ContC3 = "Yes" And cType3 = "delete" Then Print #1, "dn: " + "ou=" + OrgUnit3 + ",o=" + Org3
'If ContC4 = "Yes" And cType4 = "delete" Then Print #1, "changetype: " + cType3

'If ContC4 = "Yes" And cType4 = "delete" Then Print #1,
'If ContC4 = "Yes" And cType4 = "delete" Then Print #1, "dn: " + "ou=" + OrgUnit44 + ",ou=" + OrgUnit4 + ",o=" + Org4
'If ContC4 = "Yes" And cType4 = "delete" Then Print #1, "changetype: " + cType4

'If ContC4 = "Yes" And cType4 = "delete" Then Print #1,
'If ContC4 = "Yes" And cType4 = "delete" Then Print #1, "dn: " + "ou=" + OrgUnit4 + ",o=" + Org4
'If ContC4 = "Yes" And cType4 = "delete" Then Print #1, "changetype: " + cType4

   Close #1
   
'End Main LDIF

Step299:

   Open RICEBatch For Output As #2
      
      Print #2, "path %PATH%;C:\Program Files\UserSetBuilder"
      If GetLDAP = "Yes" Then GoTo Step300
      If Estop = "Yes" Then Estop = "" Else Estop = " -c"
      If Username = "" And Password = "" Then Port = "389"
      If Username = "" And Password = "" Then Print #2, "ice -S LDIF -f " + filename + Estop + " -D LDAP -s " + IPAddress + " -p " + Port
      If Username <> "" And Password <> "" Then Print #2, "ice -S LDIF -f " + filename + Estop + " -D LDAP -s " + IPAddress + " -p " + Port + " -d " + Username + " -w " + Password + " -L " + Rootcert
      
Step300:
      If GetLDAP = "No" Then GoTo Step301
      If Username = "" And Password = "" Then Port = "389"
      If Username = "" And Password = "" Then Print #2, "ice -S LDAP -s " + IPAddress + " -p " + Port + " -D LDIF -f " + filename
      If Username <> "" And Password <> "" Then Print #2, "ice -S LDAP -s " + IPAddress + " -p " + Port + " -d " + Username + " -w " + Password + " -L " + Rootcert + " -D LDIF -f " + filename

      Print #2,
      
Step301:

   Close #2

Call Shell(RICEBatch, 1)
        
End Sub
   
'Private Sub Command7()
'    LDAPUpdate = Text50.Text
'        Call Shell(LDAPUpdate, 1)
'End Sub


'Private Sub Command3_Click()
'   Dim IPAddress As String
'   Dim Port As String
'   Dim Username As String
'   Dim Password As String
'   Dim filename As String
'   Dim RICEBatch As String
         
'   IPAddress = Text70.Text
'   Port = Text71.Text
'   Username = Text72.Text
'   Password = Text73.Text
'   LDIFfile = Text75.Text
'   RootCert = Text76.Text
'   RICEBatch = Text74.Text
                     
'Start Main ICE Batch File Creation

'   Open RICEBatch For Output As #1
      
'      Print #1, "path %PATH%;C:\Program Files\UserSetBuilder"
'      Print #1, "ice -S LDIF -f " + LDIFfile + " -D LDAP -s " + IPAddress + " -p " + Port + " -d " + Username + " -w " + Password + " -L " + RootCert +
'" -t #Remove -t to stop on error."
'      Print #1,
      
'   Close #1
'End Sub
'Private Sub Command4_Click()
'Text75.Text = "C:\temp\ldif_file.ldi"
'CommonDialog1.ShowOpen
'Text74.Text = CommonDialog1.filename
'If CommonDialog1.filename = "" Then Text74.Text = "C:\temp\ldif_file.ldi"
'End Sub
Private Sub Command5_Click()
Text76.Text = "C:\Temp\Rootcert.der"
CommonDialog1.ShowOpen
Text76.Text = CommonDialog1.filename
If CommonDialog1.filename = "" Then Text76.Text = "C:\Temp\Rootcert.der"
End Sub
'Private Sub Command6_Click()
'Text77.Text = "C:\temp\RICE.BAT"
'CommonDialog1.ShowOpen
'Text77.Text = CommonDialog1.filename
'If CommonDialog1.filename = "" Then Text77.Text = "C:\temp\RICE.BAT"
'End Sub

'Private Sub Command7_Click()
'LDAPUpdate = Text100.Text
'Call Shell(LDAPUpdate, vbNormalFocus)
'End Sub

Private Sub Command8_Click()
Text51.Text = "C:\Temp\hdir.bat"
CommonDialog1.ShowOpen
Text51.Text = CommonDialog1.filename
If CommonDialog1.filename = "" Then Text51.Text = "C:\Temp\hdir.bat"
End Sub

Private Sub Command9_Click()

   Dim userid1 As String
'   Dim userid2 As String
'   Dim userid3 As String
'   Dim userid4 As String
   Dim bfilename As String
   Dim beginval As Long
   Dim endval As Long
   Dim retval As Long
   Dim path As String
      
   path = Text52.Text
   userid1 = Text1.Text
'   userid2 = Text6.Text
'   userid3 = Text11.Text
'   userid4 = Text16.Text
   beginval = 1
   endval = Val(Combo2.Text)
   bfilename = Text51.Text
   
'Begin User Home Directory Creation
   
   Open bfilename For Output As #2
   
If endval = "1" Then GoTo Step111
   
   ProgressBar2.Min = beginval
   ProgressBar2.Max = endval
   ProgressBar2.Value = ProgressBar2.Min
   ProgressBar2.Visible = False
   
Step111:
   
   For i = beginval To endval
   ProgressBar2.Value = i
      
      If userid1 = "" Then GoTo Step100
      Print #2, "md " + userid1 + Format(i)
      Print #2, "cd " + userid1 + Format(i)
      Print #2, "md public_html"
      Print #2, "cd public_html"
      Print #2, "ECHO " + userid1 + Format(i) + " > " + "index.html"
      Print #2, "cd ..\.."
      Print #2,
      
Step100:

'      If userid2 = "" Then GoTo Step101
'      Print #2, "md " + userid2 + Format(i)
'      Print #2, "cd " + userid2 + Format(i)
'      Print #2, "md public_html"
'      Print #2, "cd public_html"
'      Print #2, "ECHO " + userid2 + Format(i) + " > " + "index.html"
'      Print #2, "cd ..\.."
'      Print #2,
         
'Step101:

'      If userid3 = "" Then GoTo Step102
'      Print #2, "md " + userid3 + Format(i)
'      Print #2, "cd " + userid3 + Format(i)
'      Print #2, "md public_html"
'      Print #2, "cd public_html"
'      Print #2, "ECHO " + userid3 + Format(i) + " > " + "index.html"
'      Print #2, "cd ..\.."
'      Print #2,
      
'Step102:

'      If userid4 = "" Then GoTo Step103
'      Print #2, "md " + userid4 + Format(i)
'      Print #2, "cd " + userid4 + Format(i)
'      Print #2, "md public_html"
'      Print #2, "cd public_html"
'      Print #2, "ECHO " + userid4 + Format(i) + " > " + "index.html"
'      Print #2, "cd ..\.."
'      Print #2,
      
'Step103:
      
      
If endval = "1" Then GoTo Step110

ProgressBar2.Visible = True

   Next i
   
Step110:

   Close #2
   
        retval = ShellExecute(Form1.hwnd, "open", bfilename, "-fast", path, _
                SW_MAXIMIZE)
   
End Sub

'Sub Command10_Click()
'      Dim retval As Long
'      Dim bfilename As String
'      Dim path As String
      
'      bfilename = Text51.Text
'      path = Text52.Text
      
        ' Open a DOS Box
        'retval = ShellExecute(0, "open", "cmd", "", "C:\", SW_SHOW)
        '1. Run the program:
        'retval = ShellExecute(Form1.hwnd, "open", "C:\MyProg\startup.exe", "-fast", "C:\MyProg\", _
        '        SW_MAXIMIZE)
'        retval = ShellExecute(Form1.hwnd, "open", bfilename, "-fast", path, _
'                SW_MAXIMIZE)
        ' 2. Open the document:
        'retval = ShellExecute(Form1.hwnd, "open", "C:\Project\nucleus.doc", "", "C:\Project\", _
        '        SW_RESTORE)
        ' 3. Print the document (minimized in case a window opens):
        'retval = ShellExecute(Form1.hwnd, "print", "C:\Project\picture.bmp", "", "C:\Project\", _
        '        SW_MINIMIZE)
'End Sub

Private Sub Command11_Click()
Dim getdir As String
    getdir = Text52.Text
        getdir = BrowseForFolder(Me, "Select A Directory to Create User Home Directories", getdir)
    If Len(getdir) = 0 Then Exit Sub  'user selected cancel
    Text52.Text = getdir
   End Sub

Private Sub Command12_Click()
Text74.Text = "C:\temp\RICE.BAT"
CommonDialog1.ShowOpen
Text74.Text = CommonDialog1.filename
If CommonDialog1.filename = "" Then Text74.Text = "C:\temp\RICE.BAT"
End Sub
Sub Command13_Click()
      Dim retval As Long
      Dim bfilename As String
      Dim path As String
      
      'bfilename = Text51.Text
      'path = Text52.Text
      
        ' Open a DOS Box
        'retval = ShellExecute(0, "open", "cmd", "", "C:\", SW_SHOW)
        '1. Run the program:
        'retval = ShellExecute(Form1.hwnd, "open", "C:\MyProg\startup.exe", "-fast", "C:\MyProg\", _
        '        SW_MAXIMIZE)
        'retval = ShellExecute(Form1.hwnd, "open", bfilename, "-fast", path, _
        '        SW_MAXIMIZE)
        ' 2. Open the document:
        retval = ShellExecute(Form1.hwnd, "open", "ice.log", "", "", _
                SW_RESTORE)
        ' 3. Print the document (minimized in case a window opens):
        'retval = ShellExecute(Form1.hwnd, "print", "C:\Project\picture.bmp", "", "C:\Project\", _
        '        SW_MINIMIZE)
End Sub
