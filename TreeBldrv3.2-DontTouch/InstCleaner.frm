VERSION 5.00
Begin VB.Form InstCleaner 
   Caption         =   "Cleanup Installation Files"
   ClientHeight    =   2490
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3765
   Icon            =   "InstCleaner.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2490
   ScaleWidth      =   3765
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClean 
      Caption         =   "Delete All Installation Files"
      Height          =   375
      Left            =   840
      TabIndex        =   2
      ToolTipText     =   "Cleanup files created by Tree Builder."
      Top             =   1080
      Width           =   2295
   End
   Begin VB.CommandButton CmdBrowse5 
      Caption         =   "..."
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
      TabIndex        =   1
      ToolTipText     =   "Browse to a location of your choosing to create files using Tree Builder."
      Top             =   1080
      Width           =   375
   End
   Begin VB.TextBox DelInst 
      Height          =   285
      Left            =   360
      TabIndex        =   0
      Text            =   "C:\Temp\TreeBuilder\"
      ToolTipText     =   $"InstCleaner.frx":0CCA
      Top             =   600
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "CAUTION! This will delete ALL directories, sub-directories and their contents!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   360
      TabIndex        =   4
      ToolTipText     =   "This will delete ALL directories and sub-directories that you point to."
      Top             =   1680
      Width           =   2895
   End
   Begin VB.Label Label13 
      Caption         =   "Delete Insallation Files and Directories"
      Height          =   255
      Left            =   360
      TabIndex        =   3
      ToolTipText     =   "Delete Insallation Files and Directories by clicking the Delete Button."
      Top             =   240
      Width           =   2895
   End
End
Attribute VB_Name = "InstCleaner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdBrowse5_Click()
Dim getdir As String
    getdir = DelInst.Text
        getdir = BrowseForFolder(Me, "Select A Directory to Create User Home Directories in", getdir)
    If Len(getdir) = 0 Then Exit Sub  'user selected cancel
    DelInst.Text = getdir
End Sub
Private Sub cmdClean_Click()
Dim sFolder() As String
Dim sTmpFolder As String
Dim sFile As String
Dim sPath As String
Dim i As Integer
Dim MyVar As String

On Error GoTo ErrorHandler
sPath = DelInst.Text   'Set root path
DeleteFile (sPath + "Support" + "\" + "ASYCFILT.DLL")
DeleteFile (sPath + "Support" + "\" + "COMCAT.DLL")
DeleteFile (sPath + "Support" + "\" + "COMCT232.OCX")
DeleteFile (sPath + "Support" + "\" + "COMCTL32.OCX")
DeleteFile (sPath + "Support" + "\" + "comdlg32.OCX")
DeleteFile (sPath + "Support" + "\" + "ICE.cfg")
DeleteFile (sPath + "Support" + "\" + "ICE.exe")
DeleteFile (sPath + "Support" + "\" + "icenativ.DLL")
DeleteFile (sPath + "Support" + "\" + "ICE_help.txt")
DeleteFile (sPath + "Support" + "\" + "ldaphdlr.DLL")
DeleteFile (sPath + "Support" + "\" + "ldapsdk.DLL")
DeleteFile (sPath + "Support" + "\" + "ldapssl.DLL")
DeleteFile (sPath + "Support" + "\" + "ldapx.DLL")
DeleteFile (sPath + "Support" + "\" + "ldif.DLL")
DeleteFile (sPath + "Support" + "\" + "legacylogin.DLL")
DeleteFile (sPath + "Support" + "\" + "MSCOMCTL.OCX")
DeleteFile (sPath + "Support" + "\" + "MSCOMM32.OCX")
DeleteFile (sPath + "Support" + "\" + "MSVBVM60.DLL")
DeleteFile (sPath + "Support" + "\" + "OLEAUT32.DLL")
DeleteFile (sPath + "Support" + "\" + "OLEPRO32.DLL")
DeleteFile (sPath + "Support" + "\" + "Readme.txt")
DeleteFile (sPath + "Support" + "\" + "SETUP.exe")
DeleteFile (sPath + "Support" + "\" + "SETUP.LST")
DeleteFile (sPath + "Support" + "\" + "SETUP1.exe")
DeleteFile (sPath + "Support" + "\" + "ST6UNST.exe")
DeleteFile (sPath + "Support" + "\" + "STDOLE2.TLB")
DeleteFile (sPath + "Support" + "\" + "TABCTL32.OCX")
DeleteFile (sPath + "Support" + "\" + "TreeBldr3.DDF")
DeleteFile (sPath + "Support" + "\" + "TreeBldrv3.BAT")
DeleteFile (sPath + "Support" + "\" + "TreeBldrv3.exe")
DeleteFile (sPath + "Support" + "\" + "VB6STKIT.DLL")
RmDir (sPath + "Support")
DeleteFile (sPath + "SETUP.exe")
DeleteFile (sPath + "SETUP.LST")
DeleteFile (sPath + "TreeBldrv3.CAB")
ChDir ".."
RmDir (sPath)

On Error GoTo ErrorResume   'Resume next in case file is locked

ErrorResume:
'in case file is locked
Resume Next
ErrorHandler:
If Err.Description = "" Then Resume Next Else MsgBox Err.Description

'MyVar = MsgBox("Installation files have been deleted.", vbOK, "Done!")
MyVar = MsgBox("Installation files have been deleted.", 0, "Done!")
               If MyVar = vbOK Then
               Unload InstCleaner 'MsgBox ("OK Pressed")
               End If

End Sub
Private Sub cmdCleanORG_Click()
Dim sFolder() As String
Dim sTmpFolder As String
Dim sFile As String
Dim sPath As String
Dim i As Integer


On Error GoTo ErrorHandler
sPath = DelInst.Text   'Set root path
sTmpFolder = Dir(sPath, vbDirectory)   'Get all subfolders in directory and load into an array
Do While sTmpFolder <> ""
'Make sure not in root dir If sTmpFolder <> "." And sTmpFolder <> ".." Then 'Make sure it's a folder
If (GetAttr(sPath & sTmpFolder) And vbDirectory) = vbDirectory Then 'Add folder to array
ReDim Preserve sFolder(0 To i)
sFolder(i) = sTmpFolder & "\"
i = i + 1
'End If
End If
sTmpFolder = Dir   'Get next folder
Loop
On Error GoTo ErrorResume   'Resume next in case file is locked
'Delete all from Main Folder
Do Until Dir(sPath) = ""
sFile = Dir(sPath)
Kill sPath & sFile
RmDir sTmpFolder ' or sPath
Loop
'Delete from all Sub Folders
For i = 0 To UBound(sFolder)
sTmpFolder = sPath & sFolder(i)
Do Until Dir(sTmpFolder) = ""
sFile = Dir(sTmpFolder)
Kill sTmpFolder & sFile
RmDir sTmpFolder ' or sPath
Loop
Next
Exit Sub
ErrorResume:
'in case file is locked
Resume Next
ErrorHandler:
MsgBox Err.Description

MyVar = MsgBox("Installation files have been deleted.", vbOKCancel, "Done!")
               If MyVar = vbOK Then
               Unload InstCleaner 'MsgBox ("OK Pressed")
               End If

End Sub

