VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   8715
   ClientLeft      =   2565
   ClientTop       =   1500
   ClientWidth     =   9765
   Icon            =   "frmOptions.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8715
   ScaleWidth      =   9765
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   3
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample4 
         Caption         =   "Sample 4"
         Height          =   1785
         Left            =   2100
         TabIndex        =   11
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
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample3 
         Caption         =   "Sample 3"
         Height          =   1785
         Left            =   1545
         TabIndex        =   10
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
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample2 
         Caption         =   "Define Tree Layout"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6185
         Left            =   210
         TabIndex        =   5
         Top             =   255
         Width           =   7535
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   7740
      Index           =   0
      Left            =   240
      ScaleHeight     =   7740
      ScaleWidth      =   9405
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   480
      Width           =   9405
      Begin VB.Frame fraSample1 
         Caption         =   "Create Users"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   7185
         Left            =   210
         TabIndex        =   9
         Top             =   255
         Width           =   9015
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   240
            TabIndex        =   21
            Text            =   "111.222.333.444"
            Top             =   2040
            Width           =   1335
         End
         Begin VB.TextBox Text2 
            Height          =   285
            Left            =   1680
            TabIndex        =   20
            Text            =   "636"
            Top             =   2040
            Width           =   495
         End
         Begin VB.TextBox Text3 
            Height          =   285
            Left            =   2280
            TabIndex        =   19
            Text            =   "cn=admin,o=novell"
            Top             =   2040
            Width           =   1695
         End
         Begin VB.TextBox Text4 
            Height          =   285
            Left            =   4080
            TabIndex        =   18
            Text            =   "test"
            Top             =   2040
            Width           =   855
         End
         Begin VB.TextBox Text6 
            Height          =   285
            Left            =   2400
            TabIndex        =   17
            Text            =   "F:\Public\Rootcert.der"
            Top             =   2760
            Width           =   1815
         End
         Begin VB.TextBox Text5 
            Height          =   285
            Left            =   6240
            TabIndex        =   16
            Text            =   "C:\temp\RICE.BAT"
            Top             =   2040
            Width           =   1815
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
            Left            =   12000
            TabIndex        =   15
            Top             =   2040
            Width           =   975
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
            Left            =   1320
            TabIndex        =   14
            Top             =   2760
            Width           =   975
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            ItemData        =   "frmOptions.frx":000C
            Left            =   5040
            List            =   "frmOptions.frx":0016
            TabIndex        =   13
            Text            =   "No"
            Top             =   2040
            Width           =   1095
         End
         Begin VB.ComboBox Combo2 
            Height          =   315
            ItemData        =   "frmOptions.frx":0023
            Left            =   240
            List            =   "frmOptions.frx":002D
            TabIndex        =   12
            Text            =   "No"
            Top             =   2760
            Width           =   735
         End
         Begin VB.Label Label1 
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
            Left            =   240
            TabIndex        =   29
            Top             =   1800
            Width           =   1215
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
            Left            =   1680
            TabIndex        =   28
            Top             =   1800
            Width           =   495
         End
         Begin VB.Label Label3 
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
            Left            =   2280
            TabIndex        =   27
            Top             =   1800
            Width           =   1215
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
            Left            =   4080
            TabIndex        =   26
            Top             =   1800
            Width           =   975
         End
         Begin VB.Label Label8 
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
            Left            =   2400
            TabIndex        =   25
            Top             =   2520
            Width           =   2055
         End
         Begin VB.Label Label6 
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
            Left            =   6240
            TabIndex        =   24
            Top             =   1800
            Width           =   1935
         End
         Begin VB.Label Label5 
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
            Left            =   5040
            TabIndex        =   23
            Top             =   1800
            Width           =   1215
         End
         Begin VB.Label Label7 
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
            Left            =   240
            TabIndex        =   22
            Top             =   2520
            Width           =   975
         End
      End
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   375
      Left            =   8400
      TabIndex        =   3
      Top             =   8280
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   7200
      TabIndex        =   2
      Top             =   8280
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   6000
      TabIndex        =   1
      Top             =   8280
      Width           =   1095
   End
   Begin MSComctlLib.TabStrip tbsOptions 
      Height          =   8085
      Left            =   105
      TabIndex        =   0
      Top             =   120
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   14261
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   4
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Create"
            Key             =   "Group1"
            Object.ToolTipText     =   "Set Options for Group 1"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "User"
            Key             =   "Group2"
            Object.ToolTipText     =   "Set Options for Group 2"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Tree"
            Key             =   "Group3"
            Object.ToolTipText     =   "Set Options for Group 3"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "ICE"
            Key             =   "Group4"
            Object.ToolTipText     =   "Set Options for Group 4"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdApply_Click()
    MsgBox "Place code here to set options w/o closing dialog!"
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    MsgBox "Place code here to set options and close dialog!"
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    'handle ctrl+tab to move to the next tab
    If Shift = vbCtrlMask And KeyCode = vbKeyTab Then
        i = tbsOptions.SelectedItem.Index
        If i = tbsOptions.Tabs.Count Then
            'last tab so we need to wrap to tab 1
            Set tbsOptions.SelectedItem = tbsOptions.Tabs(1)
        Else
            'increment the tab
            Set tbsOptions.SelectedItem = tbsOptions.Tabs(i + 1)
        End If
    End If
End Sub

Private Sub Form_Load()
    'center the form
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
End Sub

Private Sub tbsOptions_Click()
    
    Dim i As Integer
    'show and enable the selected tab's controls
    'and hide and disable all others
    For i = 0 To tbsOptions.Tabs.Count - 1
        If i = tbsOptions.SelectedItem.Index - 1 Then
            picOptions(i).Left = 210
            picOptions(i).Enabled = True
        Else
            picOptions(i).Left = -20000
            picOptions(i).Enabled = False
        End If
    Next
    
End Sub
