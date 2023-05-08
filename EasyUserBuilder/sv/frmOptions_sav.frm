VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   8625
   ClientLeft      =   2565
   ClientTop       =   1500
   ClientWidth     =   11925
   Icon            =   "frmOptions.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8625
   ScaleWidth      =   11925
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   7380
      Index           =   3
      Left            =   210
      ScaleHeight     =   7380
      ScaleWidth      =   11685
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   480
      Width           =   11685
      Begin VB.Frame fraSample4 
         Height          =   6945
         Left            =   210
         TabIndex        =   5
         Top             =   255
         Width           =   10815
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   7380
      Index           =   2
      Left            =   210
      ScaleHeight     =   7380
      ScaleWidth      =   11685
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   480
      Width           =   11685
      Begin VB.Frame fraSample3 
         Caption         =   "Sample 3"
         Height          =   6945
         Left            =   210
         TabIndex        =   5
         Top             =   255
         Width           =   10815
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
         Caption         =   "Tree Information Sample 2"
         Height          =   6945
         Left            =   210
         TabIndex        =   5
         Top             =   255
         Width           =   10815
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   7380
      Index           =   0
      Left            =   210
      ScaleHeight     =   7380
      ScaleWidth      =   11685
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   480
      Width           =   11685
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   11040
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Frame fraSample1 
         Caption         =   "Sample 1"
         Height          =   6945
         Left            =   210
         TabIndex        =   5
         Top             =   255
         Width           =   10815
         Begin VB.TextBox Text70 
            Height          =   285
            Left            =   120
            TabIndex        =   24
            Text            =   "111.222.333.444"
            Top             =   3840
            Width           =   1335
         End
         Begin VB.TextBox Text71 
            Height          =   285
            Left            =   1560
            TabIndex        =   23
            Text            =   "636"
            Top             =   3840
            Width           =   495
         End
         Begin VB.TextBox Text72 
            Height          =   285
            Left            =   2160
            TabIndex        =   22
            Text            =   "cn=admin,o=novell"
            Top             =   3840
            Width           =   1695
         End
         Begin VB.TextBox Text73 
            Height          =   285
            Left            =   3960
            TabIndex        =   21
            Text            =   "test"
            Top             =   3840
            Width           =   855
         End
         Begin VB.TextBox Text76 
            Height          =   285
            Left            =   3120
            TabIndex        =   20
            Text            =   "F:\Public\Rootcert.der"
            Top             =   4560
            Width           =   1815
         End
         Begin VB.TextBox Text74 
            Height          =   285
            Left            =   120
            TabIndex        =   19
            Text            =   "C:\temp\RICE.BAT"
            Top             =   4560
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
            Left            =   11880
            TabIndex        =   18
            Top             =   3840
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
            Left            =   5040
            TabIndex        =   17
            Top             =   4440
            Width           =   975
         End
         Begin VB.ComboBox Combo12 
            Height          =   315
            ItemData        =   "frmOptions.frx":000C
            Left            =   4920
            List            =   "frmOptions.frx":0016
            TabIndex        =   16
            Text            =   "No"
            Top             =   3840
            Width           =   1095
         End
         Begin VB.ComboBox Combo13 
            Height          =   315
            ItemData        =   "frmOptions.frx":0023
            Left            =   2160
            List            =   "frmOptions.frx":002D
            TabIndex        =   15
            Text            =   "No"
            Top             =   4560
            Width           =   735
         End
         Begin VB.ComboBox Combo2 
            Height          =   315
            ItemData        =   "frmOptions.frx":003A
            Left            =   2400
            List            =   "frmOptions.frx":0056
            TabIndex        =   13
            Text            =   "1"
            Top             =   600
            Width           =   1215
         End
         Begin VB.CommandButton Command1 
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
            Height          =   615
            Left            =   8760
            TabIndex        =   12
            Top             =   5880
            Width           =   1455
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
            TabIndex        =   32
            Top             =   3600
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
            TabIndex        =   31
            Top             =   3600
            Width           =   495
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
            TabIndex        =   30
            Top             =   3600
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
            Left            =   3960
            TabIndex        =   29
            Top             =   3600
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
            Left            =   3120
            TabIndex        =   28
            Top             =   4320
            Width           =   2055
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
            Left            =   120
            TabIndex        =   27
            Top             =   4320
            Width           =   1935
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
            TabIndex        =   26
            Top             =   3600
            Width           =   1215
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
            Left            =   2160
            TabIndex        =   25
            Top             =   4320
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Number of Users to Create"
            Height          =   255
            Left            =   240
            TabIndex        =   14
            Top             =   720
            Width           =   2055
         End
      End
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   375
      Left            =   10440
      TabIndex        =   3
      Top             =   8040
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   9240
      TabIndex        =   2
      Top             =   8040
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   8040
      TabIndex        =   1
      Top             =   8040
      Width           =   1095
   End
   Begin MSComctlLib.TabStrip tbsOptions 
      Height          =   7725
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   13626
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   4
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Group 1"
            Key             =   "Group1"
            Object.ToolTipText     =   "Set Options for Group 1"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Group 2"
            Key             =   "Group2"
            Object.ToolTipText     =   "Set Options for Group 2"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Group 3"
            Key             =   "Group3"
            Object.ToolTipText     =   "Set Options for Group 3"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Group 4"
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

Private Sub Command1_Click()

   Dim userid1 As String
   Dim givenname1 As String
   Dim surname1 As String
   Dim password1 As String
   Dim Lang1 As String
   Dim Org1 As String
   Dim OrgUnit1 As String
   Dim OrgUnit11 As String
   Dim OrgUnit111 As String
   Dim OrgUnit1111 As String
   Dim cType1 As String
   Dim servername As String
   Dim domainname As String
   Dim filename As String
   Dim homedir As String
   Dim beginval As Long
   Dim endval As Long
   Dim ver As String
   Dim passwordm As String
   Dim context1 As String
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
   servername = Text61.Text
   domainname = Text62.Text
   filename = Text50.Text
   homedir = Text63.Text
   IPAddress = Text70.Text
   Port = Text71.Text
   Username = Text72.Text
   Password = Text73.Text
   Rootcert = Text76.Text
   RICEBatch = Text74.Text
   beginval = 1
   endval = Val(Combo2.Text)
   passwordm = Combo1.Text
   ContC1 = Combo8.Text
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
      Print #1, "mail: " + userid1 + Format(i) + "@" + domainname
      Print #1, "uid: " + userid1 + Format(i)
      Print #1, "givenName: " + givenname1 + Format(i)
      Print #1, "sn: " + surname1 + Format(i)
      Print #1, "objectclass: inetOrgPerson"
      Print #1, "objectclass: organizationalPerson"
      Print #1, "objectclass: person"
      Print #1, "objectclass: ndsLoginProperties"
      Print #1, "objectclass: top"
      If passwordm = "Yes" Then Print #1, "userpassword: " + userid1 + Format(i) Else Print #1, "userpassword: " + password1
      
      If homedir = "cn=server1_Vol1,ou=OrganizationalUnit,o=Container#0#\Users_directory" Then GoTo Step1
      If ver = "NetWare 6" Then Print #1, "ndsHomeDirectory: " + homedir + "\" + userid1 + Format(i)
      If ver = "NetWare 5.1" Then Print #1, "homeDirectory: " + homedir + "\" + userid1 + Format(i)
      
Step1:
     
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

