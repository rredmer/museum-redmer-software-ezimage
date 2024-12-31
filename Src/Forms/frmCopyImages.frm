VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmCopyImages 
   BackColor       =   &H00C0E0FF&
   Caption         =   "Align Images With Database"
   ClientHeight    =   4455
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10035
   LinkTopic       =   "Form1"
   ScaleHeight     =   4455
   ScaleWidth      =   10035
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdTools 
      Height          =   705
      Index           =   1
      Left            =   855
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Align images to database"
      Top             =   3690
      Width           =   825
   End
   Begin VB.CommandButton cmdTools 
      Height          =   705
      Index           =   0
      Left            =   30
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Exit"
      Top             =   3690
      Width           =   825
   End
   Begin VB.Frame fraOptions 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Options"
      ForeColor       =   &H00C00000&
      Height          =   3645
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   9945
      Begin VB.CheckBox chkUnpad 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Remove leading 0's from scanned image files from fixed length of"
         Height          =   345
         Left            =   1830
         TabIndex        =   13
         Top             =   810
         Value           =   1  'Checked
         Width           =   4935
      End
      Begin VB.CheckBox chkFieldCopy 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Copy information from current image tag field"
         Height          =   345
         Left            =   1830
         TabIndex        =   12
         Top             =   510
         Value           =   1  'Checked
         Width           =   4065
      End
      Begin VB.DirListBox dirCopy 
         Height          =   1440
         Left            =   2100
         TabIndex        =   11
         Top             =   1770
         Width           =   2955
      End
      Begin VB.DriveListBox drvCopy 
         Height          =   315
         Left            =   2100
         TabIndex        =   10
         Top             =   1470
         Width           =   2970
      End
      Begin VB.TextBox txtNewCopyDir 
         Height          =   345
         Left            =   2100
         TabIndex        =   9
         Top             =   3210
         Width           =   2955
      End
      Begin VB.CheckBox chkCopy 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Copy Images To New Location"
         Height          =   315
         Left            =   1830
         TabIndex        =   6
         Top             =   1140
         Value           =   1  'Checked
         Width           =   3465
      End
      Begin VB.TextBox txtPadDigits 
         Height          =   315
         Left            =   6750
         TabIndex        =   5
         Text            =   "10"
         Top             =   810
         Width           =   405
      End
      Begin MSDataListLib.DataCombo dbcTarget 
         Height          =   315
         Left            =   1830
         TabIndex        =   8
         Top             =   180
         Width           =   3465
         _ExtentX        =   6112
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "characters."
         Height          =   255
         Left            =   7230
         TabIndex        =   7
         Top             =   870
         Width           =   855
      End
      Begin VB.Label Label49 
         BackStyle       =   0  'Transparent
         Caption         =   "Re-Tag Images Using"
         Height          =   255
         Left            =   135
         TabIndex        =   1
         Top             =   240
         Width           =   1740
      End
   End
   Begin VB.Label lblStatus 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Processing Status"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1740
      TabIndex        =   4
      Top             =   4050
      Visible         =   0   'False
      Width           =   8205
   End
End
Attribute VB_Name = "frmCopyImages"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------
'
' Project......: RSC EZ-IMAGE(r)
'
' Component....: frmCopyImages.frm
'
' Procedure....: (Declarations)
'
' Description..:  Form-level settings.
'
'----------------------------------------------------------------------------
Option Explicit                                             'Require explicit variable declaration
Private Const conExitButton = 0                             'Index of exit button
Private Const conProcessButton = 1                          'Index of processing images button
Private Const conRSCIcon = 101                              'Resource ID of RSC Icon
Private Const conStopButtonIcon = 110                       'Resource ID of Stop Icon
Private Const conProcessButtonIcon = 105                    'Resource ID of Process Icon
Private Const conExitButtonIcon = 112                       'Resource ID of Exit Button
Private bCancel As Boolean                                  'Cancel flag
Private bProcessing As Boolean                              'Set active when image processing loop running
Private Sub Form_Load()                                     'Process form load
    On Error GoTo ErrorHandler                              'Set error handler
    With dbcTarget                                          'with the target column data-combo
        Set .RowSource = cData                              'Set the data source to the cData provider class
        .RowMember = "COLUMNS"                              'Set the row member
        .ListField = "COLUMN"                               'Set the list field
        .BoundColumn = "COLUMN"                             'Set the bound column
        .Refresh                                            'Refresh the control
    End With
    Me.Icon = LoadResPicture(conRSCIcon, vbResIcon)         'Load the window icon
    cmdTools(conExitButton).Picture = LoadResPicture(conExitButtonIcon, vbResIcon)
    cmdTools(conProcessButton).Picture = LoadResPicture(conProcessButtonIcon, vbResBitmap)
    bCancel = False                                         'Initialize batch process cancel flag to false
    Exit Sub                                                'Exit this routine
ErrorHandler:                                               'Error handling code
    Resume Next                                             'Simply resume next line of code
End Sub
Private Sub cmdTools_Click(Index As Integer)
    Select Case Index
        Case 0                                              'Exit
            Me.Hide
        Case 1                                              'Align Images to Database
            lblStatus.Visible = True
            AlignImages
            lblStatus.Visible = False
    End Select
End Sub
Private Sub drvCopy_Change()
    dirCopy.Path = drvCopy.Drive
End Sub
Private Sub txtNewCopyDir_LostFocus()
    Dim sFolder As String
    On Error GoTo ErrorHandler
    If Trim$(txtNewCopyDir.Text) <> "" Then
        sFolder = dirCopy.Path & "\" & Trim$(txtNewCopyDir.Text)
        MkDir sFolder
        txtNewCopyDir.Text = ""
        dirCopy.Refresh
    End If
    Exit Sub
ErrorHandler:
    MsgBox "Could not create folder.", vbOKOnly + vbApplicationModal + vbInformation, EZ_CAPTION
    Resume Next
End Sub
Private Sub AlignImages()                                   'Remove leading 0's from scanned file names
    Dim sSrcName As String, sDstName As String, sExt As String
    Dim fRef As FileSystemObject, fl As TextStream
    
    On Error GoTo ErrorHandler                              'Set up error handler.
    Set fRef = New FileSystemObject
    Set fl = fRef.OpenTextFile(frmMain.dirSource.Path & "\LOG.TXT", ForWriting, True, TristateFalse)
    sExt = Right(frmMain.filSource.FileName, 4)             'Get the extension on the currently selected image file
    
    cData.rsRecords.MoveLast
    cData.rsRecords.MoveFirst
    If MsgBox("Align [" & Trim$(Str$(cData.rsRecords.RecordCount)) & "] records?", vbYesNo + vbApplicationModal + vbQuestion, EZ_CAPTION) = vbNo Then
        Exit Sub                                            'Exit this routine
    End If
    fl.WriteLine "******* Align [" & Trim$(Str$(cData.rsRecords.RecordCount)) & "] records."
        
    '---- Remove leading 0's from images
    lblStatus.Caption = "Removing leading 0's..."
    If chkUnpad.Value = 1 Then                              'Remove leading 0's from images (based on image tag).
        cData.rsRecords.MoveFirst                           'Get to the first record
        Do While Not cData.rsRecords.EOF                    'Loop until the last record
            sSrcName = Trim$(frmMain.dirSource.Path) & "\" & Trim(PADL(Trim(cData.rsRecords(frmMain.dbcImageTag.Text).Value), Val(txtPadDigits.Text))) & sExt
            sDstName = Trim$(frmMain.dirSource.Path) & "\" & Trim$(cData.rsRecords(frmMain.dbcImageTag.Text).Value) & sExt
            If Not fRef.FileExists(sDstName) Then             'If target does not exist
                If fRef.FileExists(sSrcName) Then           'If the source file exists, we can rename
                    Name sSrcName As sDstName               'Good source, no target (simply rename)
                    fl.WriteLine "     Renaming [" & sSrcName & "] to [" & sDstName & "]: OK."
                Else                                        'No Source File To Rename!!
                    fl.WriteLine "     **ERROR Renaming [" & sSrcName & "] to [" & sDstName & "]: Source not found."
                End If
            Else                                            'Target already exists!
                fl.WriteLine "     **ERROR Renaming [" & sSrcName & "] to [" & sDstName & "]: Destination exists."
            End If
            cData.rsRecords.MoveNext
        Loop
    End If
        
    '---- Re-tag images
    If chkFieldCopy.Value = 1 And Len(Trim(dbcTarget.Text)) > 0 Then
        lblStatus.Caption = "Re-tagging images..."          'Set status text to re-tagging images
        DoEvents
        fl.WriteLine "** RE-TAGGING [" & Trim(frmMain.dbcImageTag.Text) & " TO " & Trim(dbcTarget.Text) & "] OK."
        cData.rsRecords.MoveFirst                           'Move to first record in data set
        Do While Not cData.rsRecords.EOF                    'Loop for each record in data set
            If Len(Trim(dbcTarget.Text)) > 0 Then
                If Len(Trim(cData.rsRecords(Trim(dbcTarget.Text)).Value)) = 0 Then
                    cData.rsRecords(Trim(dbcTarget.Text)).Value = Left$(Trim(cData.rsRecords(Trim(frmMain.dbcImageTag.Text)).Value), 10)
                    cData.rsRecords.Update
                End If
            Else
                MsgBox "Error:  Target field not specified.", vbOKOnly + vbApplicationModal + vbExclamation, "Warning"
                Exit Sub
            End If
            
            sSrcName = Trim(frmMain.dirSource.Path) & "\" & Trim(cData.rsRecords(Trim(frmMain.dbcImageTag.Text)).Value) & sExt
            sDstName = Trim(frmMain.dirSource.Path) & "\" & Trim(cData.rsRecords(Trim(dbcTarget.Text)).Value) & sExt
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
            cData.rsRecords.MoveNext
        Loop
    End If
    
    fl.Close
    frmMain.filSource.Refresh
    cData.rsRecords.MoveFirst                               'Move to first record in data set
    MsgBox "Alignment complete.", vbOKOnly + vbApplicationModal + vbInformation, EZ_CAPTION
    Exit Sub
ErrorHandler:
    MsgBox "Error in alignment: #[" & Str(Err.Number) & "][" & Err.Description & "]"
    Exit Sub
End Sub
Private Function PADL(sSource As String, iLength As Integer) As String
    PADL = String(iLength - Len(sSource), "0") & sSource
End Function
