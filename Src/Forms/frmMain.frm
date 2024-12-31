VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H80000004&
   Caption         =   "RSC EZ-IMAGE/2000(tm) Version 2.0"
   ClientHeight    =   10005
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   13935
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10005
   ScaleWidth      =   13935
   StartUpPosition =   2  'CenterScreen
   Begin EZIMAGE.ctlImage UsrImage 
      Height          =   9345
      Left            =   -60
      TabIndex        =   8
      Top             =   360
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   16484
   End
   Begin MSComctlLib.Toolbar tlbMain 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   13935
      _ExtentX        =   24580
      _ExtentY        =   635
      ButtonWidth     =   609
      Appearance      =   1
      Style           =   1
      ImageList       =   "imgMain"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "New"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Open"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Save"
            ImageIndex      =   3
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar sbrMain 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   1
      Top             =   9690
      Width           =   13935
      _ExtentX        =   24580
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Text            =   "Idle."
            TextSave        =   "Idle."
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   21167
            MinWidth        =   21167
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab tabMain 
      Height          =   9225
      Left            =   2730
      TabIndex        =   0
      Top             =   420
      Width           =   11205
      _ExtentX        =   19764
      _ExtentY        =   16272
      _Version        =   393216
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "Source Information"
      TabPicture(0)   =   "frmMain.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "UsrData"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "imgApp"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "imgMain"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cdlDialog"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Templates"
      TabPicture(1)   =   "frmMain.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "UsrTemplate"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Composites"
      TabPicture(2)   =   "frmMain.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "UsrComposite"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Directories"
      TabPicture(3)   =   "frmMain.frx":035E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "UsrDirectory"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Process Images"
      TabPicture(4)   =   "frmMain.frx":037A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "UsrProcess"
      Tab(4).ControlCount=   1
      Begin MSComDlg.CommonDialog cdlDialog 
         Left            =   9450
         Top             =   8670
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         DefaultExt      =   "*.ezi"
         Filter          =   "EZ-IMAGE Files (*.ezi)|*.ezi"
      End
      Begin MSComctlLib.ImageList imgMain 
         Left            =   10560
         Top             =   8580
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   9
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":0396
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":0A10
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":108A
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":1704
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":1D7E
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":23F8
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":2A72
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":30EC
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":3766
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList imgApp 
         Left            =   9960
         Top             =   8580
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":3DE0
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":40FA
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin EZIMAGE.ctlData UsrData 
         Height          =   8715
         Left            =   60
         TabIndex        =   7
         Top             =   360
         Width           =   11055
         _ExtentX        =   19500
         _ExtentY        =   15372
      End
      Begin EZIMAGE.ctlComposite UsrComposite 
         Height          =   8715
         Left            =   -74940
         TabIndex        =   6
         Top             =   390
         Width           =   10995
         _ExtentX        =   19394
         _ExtentY        =   15372
      End
      Begin EZIMAGE.ctlProcess UsrProcess 
         Height          =   8715
         Left            =   -74940
         TabIndex        =   4
         Top             =   360
         Width           =   10995
         _ExtentX        =   19394
         _ExtentY        =   15372
      End
      Begin EZIMAGE.ctlTemplates UsrTemplate 
         Height          =   8775
         Left            =   -74910
         TabIndex        =   3
         Top             =   360
         Width           =   10995
         _ExtentX        =   19394
         _ExtentY        =   15478
      End
      Begin EZIMAGE.ctlDirectory UsrDirectory 
         Height          =   8715
         Left            =   -74940
         TabIndex        =   5
         Top             =   360
         Width           =   10965
         _ExtentX        =   19341
         _ExtentY        =   15372
      End
   End
   Begin VB.Menu mnuMain 
      Caption         =   "&File"
      Index           =   0
      Begin VB.Menu mnuFile 
         Caption         =   "&New"
         Index           =   0
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFile 
         Caption         =   "&Open"
         Index           =   1
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFile 
         Caption         =   "&Save"
         Index           =   2
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFile 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuFile 
         Caption         =   "E&xit"
         Index           =   4
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------
'
' Project......: RSC EZ-IMAGE(r)
'
' Component....: frmMain.frm
'
' Procedure....: (Declarations)
'
' Description..: The main form unly contains the main menu, toolbar, status bar,
'                and Tab control (SSTAB).
'                All other controls on this form are User Controls.
'
' Author.......: Ronald D. Redmer
'
' History......: 07-01-97 RDR Designed and Programmed
'
' (c) 1997-1999 Redmer Software Company, Inc.
' All Rights Reserved
'----------------------------------------------------------------------------
Option Explicit

Private Sub Form_Load()
    On Error Resume Next
    Me.Icon = imgApp.ListImages(1).Picture
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    End
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
    If MsgBox("Are you sure?", vbApplicationModal + vbYesNo + vbQuestion + vbDefaultButton2, "Exit Program") = vbNo Then Cancel = True
End Sub

Private Sub tlbMain_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error Resume Next
    mnuFile_Click Button.Index - 1
End Sub

Private Sub mnuFile_Click(Index As Integer)
    Dim iAns As Integer
    Dim s As String
    On Error Resume Next
    Select Case Index
    Case 0                  'New
        cdlDialog.FileName = "untitled.ezi"
        UsrData.NewFile
        Me.Refresh
    Case 1                  'Open
        cdlDialog.ShowOpen
        UsrData.ReadFile cdlDialog.FileName
        Me.Refresh
    Case 2                  'Save
        cdlDialog.ShowSave
        UsrData.WriteFile cdlDialog.FileName
        Me.Refresh
    Case 4
        Unload Me
    End Select
End Sub
