VERSION 5.00
Object = "{00100003-B1BA-11CE-ABC6-F5B2E79D9E3F}#1.0#0"; "LTOCX10N.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl ctlData 
   ClientHeight    =   8700
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11130
   ScaleHeight     =   8700
   ScaleWidth      =   11130
   Begin VB.Frame fraDatabase 
      Caption         =   "Database Connection"
      ForeColor       =   &H00FF0000&
      Height          =   945
      Left            =   60
      TabIndex        =   20
      Top             =   90
      Width           =   10965
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "Edit"
         Height          =   285
         Left            =   1980
         TabIndex        =   43
         Top             =   210
         Width           =   855
      End
      Begin VB.OptionButton optDB 
         Caption         =   "ODBC Database"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   42
         Top             =   630
         Width           =   1485
      End
      Begin VB.OptionButton optDB 
         Caption         =   "EZ-IMAGE Database"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   41
         Top             =   270
         Value           =   -1  'True
         Width           =   1875
      End
      Begin VB.TextBox txtUID 
         Height          =   300
         Left            =   6315
         TabIndex        =   23
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox txtPWD 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   8835
         PasswordChar    =   "*"
         TabIndex        =   22
         Top             =   600
         Width           =   1635
      End
      Begin VB.ComboBox cboDSNList 
         Height          =   315
         ItemData        =   "ctlData.ctx":0000
         Left            =   1635
         List            =   "ctlData.ctx":0002
         Sorted          =   -1  'True
         TabIndex        =   21
         Top             =   570
         Width           =   4050
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Password"
         Height          =   195
         Index           =   2
         Left            =   8070
         TabIndex        =   36
         Top             =   645
         Width           =   690
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "User ID"
         Height          =   195
         Index           =   1
         Left            =   5730
         TabIndex        =   24
         Top             =   630
         Width           =   540
      End
   End
   Begin VB.Frame fraCriteria 
      Caption         =   "Record Selection"
      ForeColor       =   &H00FF0000&
      Height          =   6435
      Left            =   60
      TabIndex        =   10
      Top             =   1050
      Width           =   10965
      Begin VB.CheckBox chkSort 
         DownPicture     =   "ctlData.ctx":0004
         Height          =   315
         Index           =   2
         Left            =   10590
         Picture         =   "ctlData.ctx":0106
         Style           =   1  'Graphical
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   1410
         UseMaskColor    =   -1  'True
         Width           =   285
      End
      Begin VB.CheckBox chkSort 
         DownPicture     =   "ctlData.ctx":0208
         Height          =   300
         Index           =   1
         Left            =   10590
         Picture         =   "ctlData.ctx":030A
         Style           =   1  'Graphical
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   1080
         UseMaskColor    =   -1  'True
         Width           =   285
      End
      Begin VB.CheckBox chkSort 
         DownPicture     =   "ctlData.ctx":040C
         Height          =   300
         Index           =   0
         Left            =   10590
         Picture         =   "ctlData.ctx":050E
         Style           =   1  'Graphical
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   750
         UseMaskColor    =   -1  'True
         Width           =   285
      End
      Begin VB.TextBox txtCriteria 
         Height          =   315
         Index           =   2
         Left            =   3900
         TabIndex        =   13
         Top             =   1410
         Width           =   3495
      End
      Begin VB.TextBox txtCriteria 
         Height          =   315
         Index           =   0
         Left            =   3900
         TabIndex        =   12
         Top             =   750
         Width           =   3495
      End
      Begin VB.TextBox txtCriteria 
         Height          =   315
         Index           =   1
         Left            =   3900
         TabIndex        =   11
         Top             =   1080
         Width           =   3495
      End
      Begin MSDataListLib.DataCombo dbcCompare 
         Height          =   315
         Index           =   0
         Left            =   2670
         TabIndex        =   14
         Top             =   750
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dbcCriteria 
         Height          =   315
         Index           =   0
         Left            =   60
         TabIndex        =   15
         Top             =   750
         Width           =   2595
         _ExtentX        =   4577
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dbcCriteria 
         Height          =   315
         Index           =   1
         Left            =   60
         TabIndex        =   16
         Top             =   1080
         Width           =   2595
         _ExtentX        =   4577
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dbcCriteria 
         Height          =   315
         Index           =   2
         Left            =   60
         TabIndex        =   17
         Top             =   1410
         Width           =   2595
         _ExtentX        =   4577
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dbcCompare 
         Height          =   315
         Index           =   1
         Left            =   2670
         TabIndex        =   18
         Top             =   1080
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dbcCompare 
         Height          =   315
         Index           =   2
         Left            =   2670
         TabIndex        =   19
         Top             =   1410
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dbcSort 
         Height          =   315
         Index           =   0
         Left            =   7800
         TabIndex        =   25
         Top             =   750
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dbcSort 
         Height          =   315
         Index           =   1
         Left            =   7800
         TabIndex        =   26
         Top             =   1080
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dbcSort 
         Height          =   315
         Index           =   2
         Left            =   7800
         TabIndex        =   27
         Top             =   1410
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dbcTables 
         Height          =   315
         Left            =   630
         TabIndex        =   28
         Top             =   210
         Width           =   3270
         _ExtentX        =   5768
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
         Object.DataMember      =   ""
      End
      Begin MSDataGridLib.DataGrid dbgRecords 
         CausesValidation=   0   'False
         Height          =   4575
         Left            =   60
         TabIndex        =   29
         Top             =   1785
         Width           =   8355
         _ExtentX        =   14737
         _ExtentY        =   8070
         _Version        =   393216
         AllowUpdate     =   -1  'True
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
            RecordSelectors =   0   'False
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin LEADLib.LEAD ledTag 
         Height          =   3750
         Left            =   8460
         TabIndex        =   30
         Top             =   1770
         Width           =   2415
         _Version        =   65537
         _ExtentX        =   4260
         _ExtentY        =   6615
         _StockProps     =   229
         BackColor       =   12632256
         Appearance      =   1
         ScaleHeight     =   246
         ScaleWidth      =   157
         DataField       =   ""
         BitmapDataPath  =   ""
         AnnDataPath     =   ""
         PanWinTitle     =   "PanWindow"
         CLeadCtrl       =   0
      End
      Begin MSComctlLib.Toolbar tlbSource 
         Height          =   780
         Left            =   8460
         TabIndex        =   35
         Top             =   5580
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   1376
         ButtonWidth     =   1164
         ButtonHeight    =   1376
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Appearance      =   1
         Style           =   1
         ImageList       =   "imgSource"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Refresh"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Add"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Erase"
               ImageIndex      =   3
            EndProperty
         EndProperty
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Sort Order"
         Height          =   195
         Index           =   5
         Left            =   7830
         TabIndex        =   34
         Top             =   540
         Width           =   720
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Selection Criteria"
         Height          =   195
         Index           =   4
         Left            =   60
         TabIndex        =   33
         Top             =   540
         Width           =   1185
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Table"
         Height          =   195
         Index           =   3
         Left            =   90
         TabIndex        =   32
         Top             =   240
         Width           =   495
      End
      Begin VB.Label lblRecordsSelected 
         Caption         =   "0 Records Selected"
         Height          =   180
         Left            =   3945
         TabIndex        =   31
         Top             =   270
         Width           =   2715
      End
   End
   Begin VB.Frame fraReTag 
      Caption         =   "Image Identification"
      ForeColor       =   &H00FF0000&
      Height          =   1155
      Left            =   60
      TabIndex        =   0
      Top             =   7485
      Width           =   10965
      Begin MSComctlLib.Toolbar tlbTag 
         Height          =   780
         Left            =   10140
         TabIndex        =   9
         Top             =   225
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   1376
         ButtonWidth     =   1138
         ButtonHeight    =   1376
         Appearance      =   1
         Style           =   1
         ImageList       =   "imgSource"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Re-Tag"
               Object.ToolTipText     =   "Rename images according to the Re-Tag field."
               ImageIndex      =   4
            EndProperty
         EndProperty
      End
      Begin VB.CheckBox chkUnpad 
         Caption         =   "Remove leading 0's from fixed length of"
         Height          =   345
         Left            =   5475
         TabIndex        =   7
         Top             =   780
         Value           =   1  'Checked
         Width           =   3075
      End
      Begin VB.TextBox txtPadDigits 
         Height          =   285
         Left            =   8565
         TabIndex        =   6
         Text            =   "10"
         Top             =   810
         Width           =   375
      End
      Begin VB.CheckBox chkFieldCopy 
         Caption         =   "Copy information from current image tag field"
         Height          =   345
         Left            =   5475
         TabIndex        =   3
         Top             =   510
         Value           =   1  'Checked
         Width           =   3465
      End
      Begin MSDataListLib.DataCombo dbcImageTag 
         Height          =   315
         Left            =   1020
         TabIndex        =   1
         Top             =   180
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dbcTarget 
         Height          =   315
         Left            =   6330
         TabIndex        =   4
         Top             =   180
         Width           =   3465
         _ExtentX        =   6112
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin VB.Label lblStatus 
         Caption         =   "Status"
         Height          =   225
         Left            =   135
         TabIndex        =   40
         Top             =   795
         Visible         =   0   'False
         Width           =   5175
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "characters."
         Height          =   255
         Left            =   8985
         TabIndex        =   8
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label49 
         BackStyle       =   0  'Transparent
         Caption         =   "Re-Tag By"
         Height          =   255
         Left            =   5475
         TabIndex        =   5
         Top             =   210
         Width           =   840
      End
      Begin VB.Label lblCMPImageName 
         Caption         =   "Image Tag"
         Height          =   255
         Left            =   90
         TabIndex        =   2
         Top             =   240
         Width           =   945
      End
   End
   Begin MSComctlLib.ImageList imgSource 
      Left            =   10500
      Top             =   8085
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
            Picture         =   "ctlData.ctx":0610
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlData.ctx":0932
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlData.ctx":0C54
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlData.ctx":0F76
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "ctlData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'----------------------------------------------------------------------------
'
' Project......: RSC EZ-IMAGE(r)
'
' Component....: ctlData
'
' Procedure....: (Declarations)
'
' Description..: Data class user interface
'
' Author.......: Ronald D. Redmer
'
' History......: 07-01-97 RDR Designed and Programmed
'
' (c) 1997-2000 Redmer Software Company, Inc. All Rights Reserved.
'----------------------------------------------------------------------------
Option Explicit                                             'Require explicit variable declaration
Private Declare Function SQLDataSources Lib "ODBC32.DLL" (ByVal henv&, ByVal fDirection%, ByVal szDSN$, ByVal cbDSNMax%, pcbDSN%, ByVal szDescription$, ByVal cbDescriptionMax%, pcbDescription%) As Integer
Private Declare Function SQLAllocEnv% Lib "ODBC32.DLL" (env&)
Private Const SQL_SUCCESS As Long = 0                       'ODBC Success
Private Const SQL_FETCH_NEXT As Long = 1                    'ODBC Fetch next record
Private Const conCriteria = 3                               'The number of search criteria expressions
Private Const conSort = 2                                   'The number of sort expressions
Private cnnCon As ADODB.Connection                          'ADO Connection to ODBC Data Source
Private tblSrc As ADODB.Recordset                           'ADO Recordset for Imaging Source Table
Private catCon As ADOX.Catalog                              'ADO Extension:  Catalog of ODBC Source
Private m_sTable As String                                  'Text name of table selected for rsRecords
Private m_sCriteria As String                               'Text criteria (WHERE clause) of current rsRecords
Private sEzdata As String                                   'The name of the Built in EZ-IMAGE database
Private rsCompare As ADODB.Recordset                        'Bindable Recordset of boolean comparisons
Public rsTables As ADODB.Recordset                          'Bindable Recordset of tables in ODBC connection
Public rsColumns As ADODB.Recordset                         'Bindable Recordset of columns in ODBC table
Public rsColumnsEdit As ADODB.Recordset                     'Bindable Recordset of columns in ODBC table (Columns for Edit Screen)
Public rsRecords As ADODB.Recordset                         'Bindable Recordset of records matching a criteria

Private Sub cmdBrowse_Click()
    frmDBOpen.Show vbModal
End Sub

Private Sub optDB_Click(Index As Integer)
    On Error Resume Next
    Select Case Index
        Case 0
            cboDSNList.Enabled = False
            txtUID.Enabled = False
            txtPWD.Enabled = False
            cboDSNList.Visible = False
            txtUID.Visible = False
            txtPWD.Visible = False
            lblLabels(1).Visible = False
            lblLabels(2).Visible = False
            cmdBrowse.Enabled = True
            cmdBrowse.Visible = True
        Case 1
            cboDSNList.Enabled = True
            txtUID.Enabled = True
            txtPWD.Enabled = True
            cboDSNList.Visible = True
            txtUID.Visible = True
            txtPWD.Visible = True
            lblLabels(1).Visible = True
            lblLabels(2).Visible = True
            cmdBrowse.Enabled = False
            cmdBrowse.Visible = False
    End Select
    dbcTables.Text = ""                         'Clear the table combo box
    OpenDatabase                                'Open the database!
    
End Sub

Private Sub UserControl_Initialize()
    On Error Resume Next
    Set cnnCon = New ADODB.Connection                       'Create new connection object
    Set catCon = New ADOX.Catalog                           'Create new catalog object
    Set rsColumns = New ADODB.Recordset                     'Create new recordset object
    Set rsColumnsEdit = New ADODB.Recordset                 'Create new recordset object
    Set rsTables = New ADODB.Recordset                      'Create new recordset object
    Set rsRecords = New ADODB.Recordset                     'Create new recordset object
    Set rsCompare = New ADODB.Recordset                     'Create new recordset object
    With rsCompare
        .Fields.Append "COMPARE", adBSTR, 25                'Append a column to hold the comparison names from a given table
        .CursorType = adOpenDynamic                         'Set cursor type to dynamic to allow additions on the fly
        .LockType = adLockOptimistic                        'Set lock type to optimistic - very low chance of contention
        .Open                                               'Open the recordset
        .AddNew "COMPARE", "Like"
        .AddNew "COMPARE", "Contains"
        .AddNew "COMPARE", "="
        .AddNew "COMPARE", "<>"
        .AddNew "COMPARE", ">"
        .AddNew "COMPARE", ">="
        .AddNew "COMPARE", "<"
        .AddNew "COMPARE", "<="
        .Update
    End With
    
    With dbgRecords
        Set .DataSource = rsRecords
        .Refresh
    End With
    '--- Initialize ODBC Connection Information
    GetDSNsAndDrivers
    optDB_Click 0
    
End Sub

Private Sub UserControl_Terminate()
    On Error Resume Next
    CloseTables
End Sub

Private Sub OpenDatabase()
    On Error Resume Next
    Dim sConnect As String
    
    '--- Open the EZ-IMAGE database if that is selected, else open ODBC
    If optDB(0).Value = True Then
        sEzdata = IIf(Right(App.Path, 1) <> "\", Trim(App.Path) + "\EZIMAGE.MDB", "EZIMAGE.MDB")
        If Dir(sEzdata, vbNormal) <> "" Then
            sConnect = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + sEzdata + ";Persist Security Info=False"
            If cnnCon.State = adStateOpen Then
                cnnCon.Close
            End If
            cnnCon.Open sConnect                                    'Open the specified connection
            Set catCon.ActiveConnection = cnnCon                    'Set ADOX Catalog to connection
            GetTables
        Else
            MsgBox "EZIMAGE database not found.", vbApplicationModal + vbCritical + vbOKOnly, "Critical Error."
            End
        End If
    Else
        If Len(cboDSNList.Text) > 0 Then
            sConnect = "Provider=MSDASQL;DSN=" & cboDSNList.Text & ";"
            sConnect = sConnect & "UID=" & txtUID.Text & ";"
            sConnect = sConnect & "PWD=" & txtPWD.Text & ";"
            If cnnCon.State = adStateOpen Then
                cnnCon.Close
            End If
            cnnCon.Open sConnect                                    'Open the specified connection
            Set catCon.ActiveConnection = cnnCon                    'Set ADOX Catalog to connection
            GetTables
        End If
    End If
End Sub

Private Sub tlbSource_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error Resume Next
    If Len(Trim(dbcTables.BoundText)) = 0 Then
        MsgBox "Please select a table.", vbApplicationModal + vbInformation + vbOKOnly
        Exit Sub
    End If
    Select Case Button.Index
        Case 1
            GetRecords
        Case 2
            rsRecords.AddNew
        Case 3
            If MsgBox("Erase the currently selected record?", vbApplicationModal + vbQuestion + vbYesNo, "Are you sure?") = vbYes Then
                rsRecords.Delete adAffectCurrent
            End If
        Case 4
    End Select
End Sub

Private Sub tlbTag_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error Resume Next
    AlignImages
End Sub

Private Sub dbcTables_LostFocus()
    On Error GoTo ErrorHandler
    GetColumns dbcTables.BoundText                    'Refresh the columns recordsets to current table
    Exit Sub
ErrorHandler:
    MsgBox "An error occured reading the source table." & vbCr & "Please select another table.", vbExclamation + vbOKOnly + vbApplicationModal, EZ_CAPTION
    Resume Next
End Sub

Private Sub dbgRecords_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    With frmMain.UsrImage
        ledTag.Load .ImagePath & Trim$(GetFieldValue(dbcImageTag.Text)) & .ImageExtension, 0, 0, 1
    End With
    rsRecords.UpdateBatch adAffectAllChapters
End Sub

Public Property Get ImageTag() As String
    On Error Resume Next
    ImageTag = Trim$(dbcImageTag.BoundText)
End Property


Private Sub cboDSNList_LostFocus()
    On Error Resume Next
    OpenDatabase
End Sub

Sub GetDSNsAndDrivers()
  On Error Resume Next
  Dim i As Integer
  Dim sDSNItem As String * 1024
  Dim sDRVItem As String * 1024
  Dim sDSN As String
  Dim sDRV As String
  Dim iDSNLen As Integer
  Dim iDRVLen As Integer
  Dim lHenv As Long     'handle to the environment
  
  cboDSNList.AddItem "(None)"
  'get the DSNs
  If SQLAllocEnv(lHenv) <> -1 Then
    Do Until i <> SQL_SUCCESS
      sDSNItem = Space(1024)
      sDRVItem = Space(1024)
      i = SQLDataSources(lHenv, SQL_FETCH_NEXT, sDSNItem, 1024, iDSNLen, sDRVItem, 1024, iDRVLen)
      sDSN = VBA.Left(sDSNItem, iDSNLen)
      sDRV = VBA.Left(sDRVItem, iDRVLen)
        
      If sDSN <> Space(iDSNLen) Then
        cboDSNList.AddItem sDSN
      End If
    Loop
  End If
  cboDSNList.ListIndex = 0
End Sub


Private Sub CloseTables()
    On Error Resume Next
    If rsTables.State = adStateOpen Then                    'If the recordset is already open
        rsTables.Close                                      'Close the recordset to prevent Open failure
    End If
    Set rsTables = Nothing
    
    If rsColumns.State = adStateOpen Then                   'If the recordset is already open
        rsColumns.Close                                     'Close the recordset to prevent Open failure
    End If
    Set rsColumns = Nothing

    If rsColumnsEdit.State = adStateOpen Then               'If the recordset is already open
        rsColumnsEdit.Close                                 'Close the recordset to prevent Open failure
    End If
    Set rsColumnsEdit = Nothing

    If rsRecords.State = adStateOpen Then                   'If the recordset is already open
        rsRecords.Close                                     'Close the recordset to prevent Open failure
    End If
    Set rsRecords = Nothing
    
    If rsCompare.State = adStateOpen Then                   'If the recordset is already open
        rsCompare.Close                                     'Close the recordset to prevent Open failure
    End If
    Set rsCompare = Nothing
        
    If cnnCon.State = adStateOpen Then
        cnnCon.Close
    End If
    Set cnnCon = Nothing

End Sub

Private Sub GetTables()
    On Error Resume Next
    Dim tbl As ADOX.Table                               'ADOX Table Object (used for enumeration)
    
    With dbcTables
        Set .RowSource = Nothing
        .RowMember = ""
        .ListField = ""
        .Refresh
    End With
    
    If rsTables.State = adStateOpen Then
        rsTables.Close
    End If
    
    With rsTables                                           'Dynamically build tables recordset
        .Fields.Append "TABLE", adBSTR, 255                 'Append a column to hold the table name
        .CursorType = adOpenDynamic                         'Set cursor type to dynamic to allow additions on the fly
        .LockType = adLockOptimistic                        'Set lock type to optimistic - very low chance of contention
        .Open                                               'Open the recordset
    End With
    
    For Each tbl In catCon.Tables                           'Enumerate through all of the tables in the database
        If Trim$(tbl.Type) <> "VIEW" And Trim$(tbl.Type) <> "SYSTEM TABLE" And Trim$(tbl.Type) <> "ACCESS TABLE" Then   'Only add non-system tables
            With rsTables
                .AddNew
                .Fields("TABLE").Value = tbl.Name
                .Update
            End With
        End If
    Next

    With dbcTables
        Set .RowSource = rsTables
        .ListField = "TABLE"
        .Refresh
    End With

End Sub

Public Sub AppendTable(sTable As String)
    On Error GoTo ErrorHandler
    Dim tbl As ADOX.Table                             'ADOX Table Object (used for enumeration)
    Set tbl = New ADOX.Table
    With tbl
        .Name = sTable
        .Columns.Append "Sample_Field", adVarWChar, 25          'Append a sample column
    End With
    catCon.Tables.Append tbl
    Set tbl = Nothing
    GetTables
    Exit Sub
ErrorHandler:
    MsgBox "An error occured adding the table.", vbApplicationModal + vbInformation + vbOKOnly, "Warning."
End Sub

Public Sub DeleteTable(sTable As String)                  'Update column definitions from Edit Window
    On Error GoTo ErrorHandler
    If sTable = "" Then
        Exit Sub
    End If
    catCon.Tables.Delete sTable
    GetTables
    Exit Sub
ErrorHandler:
    MsgBox "An error occured deleting the table.", vbApplicationModal + vbInformation + vbOKOnly, "Warning."
End Sub

Public Sub GetColumns(sTable As String)                     'Retrieve column names into table for pick lists
    On Error Resume Next
    Dim tbl As ADOX.Table                                   'ADOX Table Object (used for enumeration)
    Dim col As ADOX.Column
    Dim i As Integer
    
    Set tbl = New ADOX.Table
    Set col = New ADOX.Column
    
    If rsColumns.State = adStateOpen Then
        rsColumns.Close
    End If
    
    If rsColumnsEdit.State = adStateOpen Then
        rsColumnsEdit.Close
    End If
    
    With rsColumns                                          'Dynamically build columns recordset
        .Fields.Append "COLUMN", adBSTR, 125                'Append a column to hold the column names from a given table
        .CursorType = adOpenDynamic                         'Set cursor type to dynamic to allow additions on the fly
        .LockType = adLockOptimistic                        'Set lock type to optimistic - very low chance of contention
        .Open                                               'Open the recordset
    End With
    
    With rsColumnsEdit                                      'Dynamically build columns recordset
        .Fields.Append "COLUMN", adBSTR, 125                'Append a column to hold the column names from a given table
        .Fields.Append "LENGTH", adBSTR, 10                 'Append a column to hold the column Length from a given table
        .CursorType = adOpenDynamic                         'Set cursor type to dynamic to allow additions on the fly
        .LockType = adLockOptimistic                        'Set lock type to optimistic - very low chance of contention
        .Open                                               'Open the recordset
    End With
    
    With rsColumns                                          'Add a blank row so that criteria can be cleared out!
        .AddNew
        .Fields("COLUMN").Value = ""
        .Update
    End With
   
    Set tbl = catCon.Tables.Item(sTable)                    'Point to specified table
    For Each col In tbl.Columns                             'For each column in the table
        With rsColumns                                      'Append a record to the columns table
            .AddNew
            .Fields("COLUMN").Value = col.Name
            .Update
        End With
        With rsColumnsEdit                                  'Append a record to the columns table
            .AddNew
            .Fields("COLUMN").Value = col.Name
            .Fields("LENGTH").Value = col.DefinedSize
            .Update
        End With
    Next
    
    For i = 0 To conCriteria - 1
        With dbcCriteria(i)
            Set .RowSource = rsColumns
            .ListField = "COLUMN"
            .Refresh
        End With
    Next
    For i = 0 To conCriteria - 1
        With dbcCompare(i)
            Set .RowSource = rsCompare
            .ListField = "COMPARE"
            .Refresh
        End With
    Next
    For i = 0 To conSort
        With dbcSort(i)
            Set .RowSource = rsColumns
            .ListField = "COLUMN"
            .Refresh
        End With
    Next
    
    With dbcImageTag
        Set .RowSource = rsColumns
        .ListField = "COLUMN"
        .Refresh
    End With
    
    With dbcTarget                                          'with the target column data-combo
        Set .RowSource = frmMain.UsrData.rsColumns          'Set the data source to the frmmain.usrdata provider class
        .ListField = "COLUMN"                               'Set the list field
        .BoundColumn = "COLUMN"                             'Set the bound column
        .Refresh                                            'Refresh the control
    End With

    Set tbl = Nothing
    Set col = Nothing
End Sub


Public Sub UpdateColumns(sTable As String)                  'Update column definitions from Edit Window
    On Error GoTo ErrorHandler
    Dim tbl As ADOX.Table                                   'ADOX Table Object (used for enumeration)
    Dim col As ADOX.Column
    Dim bColFound As Boolean
    Dim i As Integer
    Dim lSize As Long
    
    If sTable = "" Then
        Exit Sub
    End If

    Set tbl = catCon.Tables.Item(sTable)                    'Point to specified table
    
    '--- Process Adds and Changes
    If rsColumnsEdit.State <> adStateOpen Then
        Exit Sub
    End If
    
    rsColumnsEdit.MoveFirst                                 'Go to first record in edit table
    Do While Not rsColumnsEdit.EOF                          'Loop until no more records
        bColFound = False
        
        '--- Process changes
        For Each col In tbl.Columns                         'For each column in the original table
            If Trim(rsColumnsEdit("COLUMN").Value) = Trim(col.Name) Then
                bColFound = True                            'Column exists
                lSize = CLng(rsColumnsEdit("LENGTH").Value)
                If lSize <= &O0 Or lSize >= &O256 Then
                    lSize = col.DefinedSize
                End If
                If lSize <> col.DefinedSize Then            'User changed the size
                    tbl.Columns.Delete col.Name             'We have to delete the column (no modification allowed to existing columns)
                    bColFound = False                       'Set flag so column with new size is appended the table
                End If
            End If
        Next
        If Not bColFound Then                              'Add the column
            tbl.Columns.Append Trim(rsColumnsEdit("COLUMN").Value), adVarWChar, rsColumnsEdit("LENGTH").Value
        End If
        rsColumnsEdit.MoveNext                              'Move to next column
    Loop
    '--- Process Deletes
    For Each col In tbl.Columns                             'For each column in the table
        rsColumnsEdit.MoveFirst
        rsColumnsEdit.Find "COLUMN='" + col.Name + "'", , adSearchForward, 1
        If rsColumnsEdit.EOF Then                           'Column has been deleted
            tbl.Columns.Delete col.Name                     'Delete the column
        End If
    Next

    Exit Sub
ErrorHandler:
    MsgBox "An error occured updating the table definition.", vbApplicationModal + vbInformation + vbOKOnly, "Warning."
End Sub

Public Sub GetRecords()
    On Error Resume Next
    Dim sSource As String
    Dim cmd As ADODB.Command
    Dim sSQL As String                                      'SQL WHERE clause built from fields
    Dim sSort As String                                     'SQL ORDER BY clause build from fields
    Dim iCriteria As Integer                                'Criteria counter
       
    sSQL = ""
    sSort = ""                                              'Build sort criteria (ORDER BY)
    If Len(Trim$(dbcSort(0).Text)) > 0 Then
        sSort = sSort & " ORDER BY " & dbcSort(0).Text
        If Len(Trim$(dbcSort(1).Text)) > 0 Then
            sSort = sSort & "," & dbcSort(1).Text
        End If
        If Len(Trim$(dbcSort(2).Text)) > 0 Then
            sSort = sSort & "," & dbcSort(2).Text
        End If
    End If
    For iCriteria = 0 To conCriteria - 1                    'Loop for each criteria section
        If Len(Trim$(dbcCriteria(iCriteria).Text)) > 0 Then
            If iCriteria = 0 Then
                sSQL = " WHERE "
            Else
                sSQL = sSQL & " AND "
            End If
            txtCriteria(iCriteria).Text = Replace(txtCriteria(iCriteria).Text, "'", "")
            Select Case dbcCompare(iCriteria).Text
                Case "Like"                                                     'Append wildcard after criteria
                    sSQL = sSQL & "(" & dbcCriteria(iCriteria).Text & " LIKE '" & txtCriteria(iCriteria).Text & "*')"
                Case "Contains"                                                 'Append wildcard on both sides of criteria
                    sSQL = sSQL & "('" & txtCriteria(iCriteria).Text & "' $ " & dbcCriteria(iCriteria).Text & ")"
                Case Else                                                       'User criteria verbatim
                    sSQL = sSQL & "(" & dbcCriteria(iCriteria).Text & " " & dbcCompare(iCriteria).Text & " '" & txtCriteria(iCriteria).Text & "')"
            End Select
        End If
    Next
    sSQL = sSQL & sSort
    
    m_sTable = dbcTables.BoundText
    m_sCriteria = sSQL
    If rsRecords.State = adStateOpen Then                    'If the recordset is already open
        rsRecords.Close                                      'Close the recordset to prevent Open failure
    End If
    If Len(m_sTable) > 0 Then
        Set cmd = New ADODB.Command
        Set cmd.ActiveConnection = cnnCon
        cmd.CommandText = "SELECT * FROM [" & m_sTable & "] " & IIf(Len(Trim(m_sCriteria)) > 0, m_sCriteria, "")
        rsRecords.CursorLocation = adUseClient
        rsRecords.Open cmd, , adOpenDynamic, adLockBatchOptimistic
    End If
    
    With dbgRecords                                         'Toggle the grid datasource members to refresh
        Set .DataSource = Nothing                           'Close the datasource property of the grid
        .Refresh                                            'Refresh the grid
        Set .DataSource = rsRecords                         'Set the data source to the cData provider class
        .Refresh                                            'Refresh the grid
    End With
    dbgRecords.ReBind                                       'Rebind the grid control
    dbgRecords.Refresh                                      'Refresh the grid again
    lblRecordsSelected.Caption = Trim$(Str$(GetRecordsCount)) & " Records Selected."
End Sub

Public Function GetRecordsCount() As Long
    On Error Resume Next
    If rsRecords.State = adStateOpen Then                   'If the recordset is already open
        GetRecordsCount = rsRecords.RecordCount
    End If
End Function

Public Function GetFieldValue(sFieldName As String) As Variant
    On Error Resume Next
    If rsRecords.State = adStateOpen Then                   'If the recordset is already open
        GetFieldValue = rsRecords(Trim(sFieldName)).Value
    End If
End Function


Private Sub AlignImages()                                   'Remove leading 0's from scanned file names
    
    On Error GoTo ErrorHandler                              'Set up error handler.
    Dim sSrcName As String, sDstName As String, sExt As String, fRef As FileSystemObject, fl As TextStream
    
    Set fRef = New FileSystemObject
    Set fl = fRef.OpenTextFile(frmMain.UsrImage.ImagePath & "LOG.TXT", ForWriting, True, TristateFalse)
    
    sExt = frmMain.UsrImage.ImageExtension                   'Get the extension on the currently selected image file
    
    rsRecords.MoveLast
    rsRecords.MoveFirst
    If MsgBox("Align [" & Trim$(Str$(rsRecords.RecordCount)) & "] records?", vbYesNo + vbApplicationModal + vbQuestion, EZ_CAPTION) = vbNo Then
        Exit Sub                                            'Exit this routine
    End If
    fl.WriteLine "******* Align [" & Trim$(Str$(rsRecords.RecordCount)) & "] records."
        
    '---- Remove leading 0's from images
    lblStatus.Caption = "Removing leading 0's..."
    If chkUnpad.Value = 1 Then                                  'Remove leading 0's from images (based on image tag).
        rsRecords.MoveFirst                                     'Get to the first record
        Do While Not rsRecords.EOF                              'Loop until the last record
            sSrcName = frmMain.UsrImage.ImagePath & Trim(PADL(Trim(rsRecords(Trim(dbcImageTag.Text)).Value), Val(txtPadDigits.Text))) & sExt
            sDstName = frmMain.UsrImage.ImagePath & Trim$(rsRecords(Trim(dbcImageTag.Text)).Value) & sExt
            If Not fRef.FileExists(sDstName) Then               'If target does not exist
                If fRef.FileExists(sSrcName) Then               'If the source file exists, we can rename
                    Name sSrcName As sDstName                   'Good source, no target (simply rename)
                    fl.WriteLine "     Renaming [" & sSrcName & "] to [" & sDstName & "]: OK."
                Else                                            'No Source File To Rename!!
                    fl.WriteLine "     **ERROR Renaming [" & sSrcName & "] to [" & sDstName & "]: Source not found."
                End If
            Else                                                'Target already exists!
                fl.WriteLine "     **ERROR Renaming [" & sSrcName & "] to [" & sDstName & "]: Destination exists."
            End If
            rsRecords.MoveNext
        Loop
    End If
        
    '---- Re-tag images
    If chkFieldCopy.Value = 1 And Len(Trim(dbcTarget.Text)) > 0 Then
        lblStatus.Caption = "Re-tagging images..."              'Set status text to re-tagging images
        DoEvents
        fl.WriteLine "** RE-TAGGING [" & Trim(dbcImageTag.Text) & " TO " & Trim(dbcTarget.Text) & "] OK."
        rsRecords.MoveFirst                                     'Move to first record in data set
        Do While Not rsRecords.EOF                              'Loop for each record in data set
            If Len(Trim(dbcTarget.Text)) > 0 Then
                If Len(Trim(rsRecords(Trim(dbcTarget.Text)).Value)) = 0 Then
                    rsRecords(Trim(dbcTarget.Text)).Value = Left$(Trim(rsRecords(Trim(dbcImageTag.Text)).Value), 10)
                    rsRecords.Update
                End If
            Else
                MsgBox "Error:  Target field not specified.", vbOKOnly + vbApplicationModal + vbExclamation, "Warning"
                Exit Sub
            End If
            
            sSrcName = frmMain.UsrImage.ImagePath & Trim(rsRecords(Trim(dbcImageTag.Text)).Value) & sExt
            sDstName = frmMain.UsrImage.ImagePath & Trim(rsRecords(Trim(dbcTarget.Text)).Value) & sExt
            If fRef.FileExists(sSrcName) Then
                If Not fRef.FileExists(sDstName) Then
                    Name sSrcName As sDstName
                    fl.WriteLine "     ** RENAMING [" & sSrcName & " TO " & sDstName & "] OK."
                Else
                    fl.WriteLine "     **ERROR RENAMING [" & sSrcName & " TO " & sDstName & "] DESTINATION EXISTS."
                End If
            Else
                fl.WriteLine "     ** ERROR RENAMING [" & sSrcName & " TO " & sDstName & "] SOURCE NOT FOUND."
            End If
            rsRecords.MoveNext
        Loop
    End If
    
    fl.Close
    rsRecords.MoveFirst                               'Move to first record in data set
    MsgBox "Alignment complete.", vbOKOnly + vbApplicationModal + vbInformation, EZ_CAPTION
    Exit Sub
ErrorHandler:
    MsgBox "Error in alignment: #[" & Str(Err.Number) & "][" & Err.Description & "]"
    Exit Sub
End Sub
Private Function PADL(sSource As String, iLength As Integer) As String
    On Error Resume Next
    PADL = String(iLength - Len(sSource), "0") & sSource
End Function


Public Function WriteFile(sName As String)
    On Error Resume Next
    Dim rsOut As ADODB.Recordset
    
    Set rsOut = New ADODB.Recordset
    With rsOut.Fields
        .Append "FileType", adBSTR, 25
        .Append "FileVersion", adDouble
        .Append "FileDate", adDate
        .Append "Comment", adVarChar, 64
        
        '--- Data Control
        .Append "Database", adVarChar, 32
        .Append "UserName", adVarChar, 32
        .Append "Password", adVarChar, 32
        .Append "TableName", adVarChar, 128
        .Append "Criteria1", adVarChar, 64
        .Append "Operator1", adVarChar, 24
        .Append "Match1", adVarChar, 64
        .Append "Criteria2", adVarChar, 64
        .Append "Operator2", adVarChar, 24
        .Append "Match2", adVarChar, 64
        .Append "Criteria3", adVarChar, 64
        .Append "Operator3", adVarChar, 24
        .Append "Match3", adVarChar, 64
        .Append "Sort1", adVarChar, 64
        .Append "Sort2", adVarChar, 64
        .Append "Sort3", adVarChar, 64
        .Append "ImageTag", adVarChar, 64
        .Append "ImagePath", adVarChar, 256
        
        '--- Template Control
        .Append "TemplatePath", adVarChar, 256
        .Append "TemplateName", adVarChar, 256
        
        '--- Composite Control
        .Append "CompPageWidth", adDouble
        .Append "CompPageHeight", adDouble
        .Append "CompMarginTop", adDouble
        .Append "CompMarginBottom", adDouble
        .Append "CompMarginLeft", adDouble
        .Append "CompMarginRight", adDouble
        .Append "CompWhiteSpace", adDouble
        .Append "CompPageCols", adInteger
        .Append "CompPageRows", adInteger
        .Append "CompSource", adInteger
        .Append "CompRowShift", adInteger
        .Append "CompOvals", adInteger
        .Append "CompTitleRow1", adInteger
        .Append "CompTitleRow2", adInteger
        .Append "CompTitleRow3", adInteger
        .Append "CompTitleRow4", adInteger
        .Append "CompTitleCol1", adInteger
        .Append "CompTitleCol2", adInteger
        .Append "CompTitleCol3", adInteger
        .Append "CompTitleCol4", adInteger
        .Append "CompTitleColCount1", adInteger
        .Append "CompTitleColCount2", adInteger
        .Append "CompTitleColCount3", adInteger
        .Append "CompTitleColCount4", adInteger
        .Append "CompCaption", adInteger
        .Append "CompCaptionField1", adVarChar, 64
        .Append "CompCaptionField2", adVarChar, 64
        .Append "CompCaptionOffset", adDouble
        .Append "CompCaptionFontSize", adDouble
        
        '--- Directory Control
        .Append "DirPageWidth", adDouble
        .Append "DirPageHeight", adDouble
        .Append "DirMarginTop", adDouble
        .Append "DirMarginBottom", adDouble
        .Append "DirMarginLeft", adDouble
        .Append "DirMarginRight", adDouble
        .Append "DirWhiteSpace", adDouble
        .Append "DirPageCols", adInteger
        .Append "DirPageRows", adInteger
        .Append "DirSource", adInteger
        .Append "DirRowShift", adInteger
        .Append "DirOvals", adInteger
        .Append "DirCaption", adInteger
        .Append "DirCaptionField1", adVarChar, 64
        .Append "DirCaptionField2", adVarChar, 64
        .Append "DirCaptionOffset", adDouble
        .Append "DirCaptionFontSize", adDouble
        
        '--- Process Control
        .Append "PrcCropFactor", adDouble
        .Append "PrcRotationAngle", adDouble
        .Append "PrcSharpenFactor", adDouble
        .Append "PrcContrast", adInteger
        .Append "PrcGamma", adDouble
        .Append "PrcDeskew", adInteger
        .Append "PrcDespeckle", adInteger
        .Append "PrcFlip", adInteger
        .Append "PrcInvert", adInteger
        .Append "PrcStretchIntensity", adInteger
        .Append "PrcResize", adInteger
        .Append "PrcPath", adVarChar, 128
        .Append "PrcFileType", adVarChar, 128
    
    End With
    
    With rsOut
        .CursorType = adOpenDynamic                         'Set cursor type to dynamic to allow additions on the fly
        .CursorLocation = adUseClient
        .LockType = adLockOptimistic                        'Set lock type to optimistic - very low chance of contention
        .Open                                               'Open the recordset
        .AddNew
        .Fields("FileType").Value = EZ_CAPTION
        .Fields("FileVersion").Value = EZ_VERSION
        .Fields("FileDate").Value = Date
        .Fields("Comment").Value = ""
        
        '--- Data control
        .Fields("Database").Value = cboDSNList.Text
        .Fields("UserName").Value = txtUID.Text
        .Fields("Password").Value = txtPWD.Text
        .Fields("TableName").Value = dbcTables.BoundText
        .Fields("Criteria1").Value = dbcCriteria(0).BoundText
        .Fields("Criteria2").Value = dbcCriteria(1).BoundText
        .Fields("Criteria3").Value = dbcCriteria(2).BoundText
        .Fields("Operator1").Value = dbcCompare(0).BoundText
        .Fields("Operator2").Value = dbcCompare(1).BoundText
        .Fields("Operator3").Value = dbcCompare(2).BoundText
        .Fields("Match1").Value = txtCriteria(0).Text
        .Fields("Match2").Value = txtCriteria(1).Text
        .Fields("Match3").Value = txtCriteria(2).Text
        .Fields("Sort1").Value = dbcSort(0).BoundText
        .Fields("Sort2").Value = dbcSort(1).BoundText
        .Fields("Sort3").Value = dbcSort(2).BoundText
        .Fields("ImageTag").Value = dbcImageTag.BoundText
        .Fields("ImagePath").Value = frmMain.UsrImage.ImagePath
        
        '--- Template Control
        .Fields("TemplatePath").Value = frmMain.UsrTemplate.TemplatePath
        .Fields("TemplateName").Value = frmMain.UsrTemplate.TemplateName
        
        '--- Composite Control
        frmMain.UsrComposite.GetControls
        .Fields("CompPageWidth").Value = frmMain.UsrComposite.PageWidth
        .Fields("CompPageHeight").Value = frmMain.UsrComposite.PageHeight
        .Fields("CompMarginTop").Value = frmMain.UsrComposite.MarginTop
        .Fields("CompMarginBottom").Value = frmMain.UsrComposite.MarginBottom
        .Fields("CompMarginLeft").Value = frmMain.UsrComposite.MarginLeft
        .Fields("CompMarginRight").Value = frmMain.UsrComposite.MarginRight
        .Fields("CompWhiteSpace").Value = frmMain.UsrComposite.WhiteSpace
        .Fields("CompPageCols").Value = frmMain.UsrComposite.PageCols
        .Fields("CompPageRows").Value = frmMain.UsrComposite.PageRows
        .Fields("CompSource").Value = frmMain.UsrComposite.PageSource
        .Fields("CompRowShift").Value = frmMain.UsrComposite.RowShift
        .Fields("CompOvals").Value = frmMain.UsrComposite.ImageOval
        .Fields("CompTitleRow1").Value = frmMain.UsrComposite.TitleRow1
        .Fields("CompTitleRow2").Value = frmMain.UsrComposite.TitleRow2
        .Fields("CompTitleRow3").Value = frmMain.UsrComposite.TitleRow3
        .Fields("CompTitleRow4").Value = frmMain.UsrComposite.TitleRow4
        .Fields("CompTitleCol1").Value = frmMain.UsrComposite.TitleCol1
        .Fields("CompTitleCol2").Value = frmMain.UsrComposite.TitleCol2
        .Fields("CompTitleCol3").Value = frmMain.UsrComposite.TitleCol3
        .Fields("CompTitleCol4").Value = frmMain.UsrComposite.TitleCol4
        .Fields("CompTitleColCount1").Value = frmMain.UsrComposite.TitleColCount1
        .Fields("CompTitleColCount2").Value = frmMain.UsrComposite.TitleColCount2
        .Fields("CompTitleColCount3").Value = frmMain.UsrComposite.TitleColCount3
        .Fields("CompTitleColCount4").Value = frmMain.UsrComposite.TitleColCount4
        .Fields("CompCaptionField1").Value = frmMain.UsrComposite.CaptionField1
        .Fields("CompCaptionField2").Value = frmMain.UsrComposite.CaptionField2
        .Fields("CompCaptionOffset").Value = frmMain.UsrComposite.CaptionOffset
        .Fields("CompCaptionFontSize").Value = frmMain.UsrComposite.CaptionFontSize
        .Fields("CompCaption").Value = frmMain.UsrComposite.ImageCaption
        
        '--- Directory Control
        .Fields("DirPageWidth") = frmMain.UsrDirectory.PageWidth
        .Fields("DirPageHeight") = frmMain.UsrDirectory.PageHeight
        .Fields("DirMarginTop") = frmMain.UsrDirectory.MarginTop
        .Fields("DirMarginBottom") = frmMain.UsrDirectory.MarginBottom
        .Fields("DirMarginLeft") = frmMain.UsrDirectory.MarginLeft
        .Fields("DirMarginRight") = frmMain.UsrDirectory.MarginRight
        .Fields("DirWhiteSpace") = frmMain.UsrDirectory.WhiteSpace
        .Fields("DirPageCols") = frmMain.UsrDirectory.PageCols
        .Fields("DirPageRows") = frmMain.UsrDirectory.PageRows
        .Fields("DirSource") = frmMain.UsrDirectory.PageSource
        .Fields("DirRowShift") = frmMain.UsrDirectory.RowShift
        .Fields("DirOvals") = frmMain.UsrDirectory.ImageOval
        .Fields("DirCaption") = frmMain.UsrDirectory.ImageCaption
        .Fields("DirCaptionField1") = frmMain.UsrDirectory.CaptionField1
        .Fields("DirCaptionField2") = frmMain.UsrDirectory.CaptionField2
        .Fields("DirCaptionOffset") = frmMain.UsrDirectory.CaptionOffset
        .Fields("DirCaptionFontSize") = frmMain.UsrDirectory.CaptionFontSize
        
        '--- Process Control
        .Fields("PrcCropFactor").Value = frmMain.UsrProcess.CropFactor
        .Fields("PrcRotationAngle").Value = frmMain.UsrProcess.RotationAngle
        .Fields("PrcSharpenFactor").Value = frmMain.UsrProcess.SharpenFactor
        .Fields("PrcContrast").Value = frmMain.UsrProcess.Contrast
        .Fields("PrcGamma").Value = frmMain.UsrProcess.Gamma
        .Fields("PrcDeskew").Value = frmMain.UsrProcess.Deskew
        .Fields("PrcDespeckle").Value = frmMain.UsrProcess.Despeckle
        .Fields("PrcFlip").Value = frmMain.UsrProcess.Flip
        .Fields("PrcInvert").Value = frmMain.UsrProcess.Invert
        .Fields("PrcStretchIntensity").Value = frmMain.UsrProcess.Stretch_Intensity
        .Fields("PrcResize").Value = frmMain.UsrProcess.HR_Size
        .Fields("prcPath").Value = frmMain.UsrProcess.ProcessPath
        .Fields("PrcFileType").Value = frmMain.UsrProcess.FileType
        
        .Update
    End With

    If Len(sName) > 0 Then
        If Len(Dir$(sName, vbNormal)) > 0 Then
            Kill sName
        End If
        rsOut.Save sName, adPersistADTG
    End If
    rsOut.Close
    Set rsOut = Nothing
End Function

Public Function ReadFile(sName As String)
    On Error Resume Next
    Dim rsIn As ADODB.Recordset
    Set rsIn = New ADODB.Recordset
    
    If Len(sName) > 0 Then
        If Len(Dir$(sName, vbNormal)) > 0 Then
            With rsIn
                .Open sName, , , , adCmdFile
                .MoveFirst
                
                '--- Data Control
                cboDSNList.Text = .Fields("Database").Value
                txtUID.Text = .Fields("UserName").Value
                txtPWD.Text = .Fields("Password").Value
                dbcTables.BoundText = .Fields("TableName").Value
                dbcCriteria(0).BoundText = .Fields("Criteria1").Value
                dbcCriteria(1).BoundText = .Fields("Criteria2").Value
                dbcCriteria(2).BoundText = .Fields("Criteria3").Value
                dbcCompare(0).BoundText = .Fields("Operator1").Value
                dbcCompare(1).BoundText = .Fields("Operator2").Value
                dbcCompare(2).BoundText = .Fields("Operator3").Value
                txtCriteria(0).Text = .Fields("Match1").Value
                txtCriteria(1).Text = .Fields("Match2").Value
                txtCriteria(2).Text = .Fields("Match3").Value
                dbcSort(0).BoundText = .Fields("Sort1").Value
                dbcSort(1).BoundText = .Fields("Sort2").Value
                dbcSort(2).BoundText = .Fields("Sort3").Value
                dbcImageTag.BoundText = .Fields("ImageTag").Value
                
                '--- Image Control
                frmMain.UsrImage.ImagePath = .Fields("ImagePath").Value
                
                '--- Template Control
                frmMain.UsrTemplate.TemplatePath = .Fields("TemplatePath").Value
                frmMain.UsrTemplate.TemplateName = .Fields("TemplateName").Value
                
                '--- Composite Control
                frmMain.UsrComposite.PageWidth = .Fields("CompPageWidth").Value
                frmMain.UsrComposite.PageHeight = .Fields("CompPageHeight").Value
                frmMain.UsrComposite.MarginTop = .Fields("CompMarginTop").Value
                frmMain.UsrComposite.MarginBottom = .Fields("CompMarginBottom").Value
                frmMain.UsrComposite.MarginLeft = .Fields("CompMarginLeft").Value
                frmMain.UsrComposite.MarginRight = .Fields("CompMarginRight").Value
                frmMain.UsrComposite.WhiteSpace = .Fields("CompWhiteSpace").Value
                frmMain.UsrComposite.PageCols = .Fields("CompPageCols").Value
                frmMain.UsrComposite.PageRows = .Fields("CompPageRows").Value
                frmMain.UsrComposite.PageSource = .Fields("CompSource").Value
                frmMain.UsrComposite.RowShift = .Fields("CompRowShift").Value
                frmMain.UsrComposite.ImageOval = .Fields("CompOvals").Value
                frmMain.UsrComposite.TitleRow1 = .Fields("CompTitleRow1").Value
                frmMain.UsrComposite.TitleRow2 = .Fields("CompTitleRow2").Value
                frmMain.UsrComposite.TitleRow3 = .Fields("CompTitleRow3").Value
                frmMain.UsrComposite.TitleRow4 = .Fields("CompTitleRow4").Value
                frmMain.UsrComposite.TitleCol1 = .Fields("CompTitleCol1").Value
                frmMain.UsrComposite.TitleCol2 = .Fields("CompTitleCol2").Value
                frmMain.UsrComposite.TitleCol3 = .Fields("CompTitleCol3").Value
                frmMain.UsrComposite.TitleCol4 = .Fields("CompTitleCol4").Value
                frmMain.UsrComposite.TitleColCount1 = .Fields("CompTitleColCount1").Value
                frmMain.UsrComposite.TitleColCount2 = .Fields("CompTitleColCount2").Value
                frmMain.UsrComposite.TitleColCount3 = .Fields("CompTitleColCount3").Value
                frmMain.UsrComposite.TitleColCount4 = .Fields("CompTitleColCount4").Value
                frmMain.UsrComposite.CaptionField1 = .Fields("CompCaptionField1").Value
                frmMain.UsrComposite.CaptionField2 = .Fields("CompCaptionField2").Value
                frmMain.UsrComposite.CaptionOffset = .Fields("CompCaptionOffset").Value
                frmMain.UsrComposite.CaptionFontSize = .Fields("CompCaptionFontSize").Value
                frmMain.UsrComposite.ImageCaption = .Fields("CompCaption").Value
                frmMain.UsrComposite.PutControls
                
                '--- Directory Control
                frmMain.UsrDirectory.PageWidth = .Fields("DirPageWidth")
                frmMain.UsrDirectory.PageHeight = .Fields("DirPageHeight")
                frmMain.UsrDirectory.MarginTop = .Fields("DirMarginTop")
                frmMain.UsrDirectory.MarginBottom = .Fields("DirMarginBottom")
                frmMain.UsrDirectory.MarginLeft = .Fields("DirMarginLeft")
                frmMain.UsrDirectory.MarginRight = .Fields("DirMarginRight")
                frmMain.UsrDirectory.WhiteSpace = .Fields("DirWhiteSpace")
                frmMain.UsrDirectory.PageCols = .Fields("DirPageCols")
                frmMain.UsrDirectory.PageRows = .Fields("DirPageRows")
                frmMain.UsrDirectory.PageSource = .Fields("DirSource")
                frmMain.UsrDirectory.RowShift = .Fields("DirRowShift")
                frmMain.UsrDirectory.ImageOval = .Fields("DirOvals")
                frmMain.UsrDirectory.ImageCaption = .Fields("DirCaption")
                frmMain.UsrDirectory.CaptionField1 = .Fields("DirCaptionField1")
                frmMain.UsrDirectory.CaptionField2 = .Fields("DirCaptionField2")
                frmMain.UsrDirectory.CaptionOffset = .Fields("DirCaptionOffset")
                frmMain.UsrDirectory.CaptionFontSize = .Fields("DirCaptionFontSize")
                
                '--- Process Control
                frmMain.UsrProcess.CropFactor = .Fields("PrcCropFactor").Value
                frmMain.UsrProcess.RotationAngle = .Fields("PrcRotationAngle").Value
                frmMain.UsrProcess.SharpenFactor = .Fields("PrcSharpenFactor").Value
                frmMain.UsrProcess.Contrast = .Fields("PrcContrast").Value
                frmMain.UsrProcess.Gamma = .Fields("PrcGamma").Value
                frmMain.UsrProcess.Deskew = .Fields("PrcDeskew").Value
                frmMain.UsrProcess.Despeckle = .Fields("PrcDespeckle").Value
                frmMain.UsrProcess.Flip = .Fields("PrcFlip").Value
                frmMain.UsrProcess.Invert = .Fields("PrcInvert").Value
                frmMain.UsrProcess.Stretch_Intensity = .Fields("PrcStretchIntensity").Value
                frmMain.UsrProcess.HR_Size = .Fields("PrcResize").Value
                frmMain.UsrProcess.ProcessPath = .Fields("prcPath").Value
                frmMain.UsrProcess.FileType = .Fields("PrcFileType").Value
                
            End With
            OpenDatabase
            If Len(dbcTables.BoundText) > 0 Then
                GetColumns dbcTables.BoundText                    'Refresh the columns recordsets to current table
            End If
        End If
    End If
    Exit Function
    rsIn.Close
    Set rsIn = Nothing
End Function

Public Function NewFile()
    On Error Resume Next
    
    '--- Data Control
    cboDSNList.Text = ""
    txtUID.Text = ""
    txtPWD.Text = ""
    dbcTables.BoundText = ""
    dbcCriteria(0).BoundText = ""
    dbcCriteria(1).BoundText = ""
    dbcCriteria(2).BoundText = ""
    dbcCompare(0).BoundText = ""
    dbcCompare(1).BoundText = ""
    dbcCompare(2).BoundText = ""
    txtCriteria(0).Text = ""
    txtCriteria(1).Text = ""
    txtCriteria(2).Text = ""
    dbcSort(0).BoundText = ""
    dbcSort(1).BoundText = ""
    dbcSort(2).BoundText = ""
    dbcImageTag.BoundText = ""
    
    '--- Image Control
    frmMain.UsrImage.ImagePath = App.Path
    
    '--- Template Control
    frmMain.UsrTemplate.TemplatePath = App.Path
    
    '--- Composite Control
    frmMain.UsrComposite.SetDefaults
    
    '--- Directory Control
    frmMain.UsrDirectory.SetDefaults
    
    '--- Process Control
    frmMain.UsrProcess.CropFactor = 0
    frmMain.UsrProcess.RotationAngle = 0
    frmMain.UsrProcess.SharpenFactor = 0
    frmMain.UsrProcess.Contrast = 0
    frmMain.UsrProcess.Gamma = 1#
    frmMain.UsrProcess.Deskew = 0
    frmMain.UsrProcess.Despeckle = 0
    frmMain.UsrProcess.Flip = 0
    frmMain.UsrProcess.Invert = 0
    frmMain.UsrProcess.Stretch_Intensity = 0
    frmMain.UsrProcess.HR_Size = 0
    frmMain.UsrProcess.ProcessPath = App.Path
    
End Function

