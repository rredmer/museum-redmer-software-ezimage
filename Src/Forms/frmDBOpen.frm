VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDBOpen 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Edit EZ-IMAGE Database"
   ClientHeight    =   6945
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12405
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6945
   ScaleWidth      =   12405
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSDataListLib.DataList dlstTables 
      Height          =   5715
      Left            =   60
      TabIndex        =   3
      Top             =   330
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   10081
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid dbgColumns 
      Height          =   5745
      Left            =   5760
      TabIndex        =   2
      Top             =   330
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   10134
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      AllowAddNew     =   -1  'True
      AllowDelete     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
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
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbDB 
      Height          =   840
      Left            =   60
      TabIndex        =   0
      Top             =   6090
      Width           =   1860
      _ExtentX        =   3281
      _ExtentY        =   1482
      ButtonWidth     =   1032
      ButtonHeight    =   1376
      ImageList       =   "imgDB"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Exit"
            Object.ToolTipText     =   "Exit"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Add"
            Object.ToolTipText     =   "Add Table"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Erase"
            Object.ToolTipText     =   "Erase Table"
            ImageIndex      =   3
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgDB 
      Left            =   11730
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDBOpen.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDBOpen.frx":0322
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDBOpen.frx":0644
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDBOpen.frx":0966
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lblWarning 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"frmDBOpen.frx":0C88
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   825
      Left            =   1920
      TabIndex        =   4
      Top             =   6090
      Width           =   10455
   End
   Begin VB.Label Label1 
      Caption         =   "Tables"
      Height          =   225
      Left            =   90
      TabIndex        =   1
      Top             =   90
      Width           =   675
   End
End
Attribute VB_Name = "frmDBOpen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private sCurTable As String

Private Sub Form_Load()
    On Error Resume Next
    With dlstTables
        Set .RowSource = frmMain.UsrData.rsTables       'Set the data source to the provider class
        .ListField = "TABLE"                            'Set the list field
        .BoundColumn = "TABLE"                          'Set the bound column
        .Refresh                                        'Refresh the control
        sCurTable = Trim(.BoundText)
    End With
    
    With dbgColumns
        Set .DataSource = frmMain.UsrData.rsColumnsEdit   'Set the data source to the provider class
        .Refresh                                        'Refresh the control
    End With
End Sub

Private Sub dlstTables_Click()
    On Error Resume Next
    frmMain.UsrData.UpdateColumns sCurTable             'Update the table definition
    frmMain.UsrData.GetColumns Trim(dlstTables.BoundText)
    With dbgColumns
        .ClearFields
        .ReBind
        .Refresh
    End With
    sCurTable = Trim(dlstTables.BoundText)
End Sub

Private Sub tlbDB_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error Resume Next
    Dim sTable As String
    Select Case Button.Index
        Case 1
            Unload Me
        Case 2
            '--- Add a new table
            sTable = Trim(InputBox("Table Name", "Add a new table"))
            If sTable <> "" Then
                frmMain.UsrData.AppendTable sTable
            End If
        Case 3
            '--- Erase the currently selected table
            If MsgBox("Delete table [" + Trim(dlstTables.BoundText) + "] ?", vbApplicationModal + vbQuestion + vbYesNo + vbDefaultButton2, "Confirm Delete") = vbYes Then
                If MsgBox("Clicking OK will PERMANENTLY delete the table." + vbCr + "If this is not what you want to do, click CANCEL now.", vbApplicationModal + vbOKCancel + vbInformation + vbDefaultButton2, "Are you sure?") = vbOK Then
                    frmMain.UsrData.DeleteTable Trim(dlstTables.BoundText)
                    dlstTables.Refresh
                End If
            End If
    End Select
    With dlstTables
        Set .RowSource = Nothing
        .Refresh
        Set .RowSource = frmMain.UsrData.rsTables       'Set the data source to the provider class
        .ListField = "TABLE"                            'Set the list field
        .BoundColumn = "TABLE"                          'Set the bound column
        .Refresh                                        'Refresh the control
        sCurTable = Trim(.BoundText)
    End With
End Sub

