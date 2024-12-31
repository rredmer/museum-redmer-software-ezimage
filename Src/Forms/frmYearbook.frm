VERSION 5.00
Begin VB.Form frmYearbook 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Create PMA Standard Yearbook CD"
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
      TabIndex        =   2
      ToolTipText     =   "Align images to database"
      Top             =   3690
      Width           =   825
   End
   Begin VB.CommandButton cmdTools 
      Height          =   705
      Index           =   0
      Left            =   30
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Exit"
      Top             =   3690
      Width           =   825
   End
   Begin VB.Frame fraOptions 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Options"
      ForeColor       =   &H00C00000&
      Height          =   3645
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   9945
      Begin VB.CheckBox chkFieldCopy 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Create README.TXT file"
         Height          =   345
         Left            =   150
         TabIndex        =   8
         Top             =   240
         Value           =   1  'Checked
         Width           =   4065
      End
      Begin VB.DirListBox dirCopy 
         Height          =   1440
         Left            =   420
         TabIndex        =   7
         Top             =   1170
         Width           =   2955
      End
      Begin VB.DriveListBox drvCopy 
         Height          =   315
         Left            =   420
         TabIndex        =   6
         Top             =   870
         Width           =   2970
      End
      Begin VB.TextBox txtNewCopyDir 
         Height          =   345
         Left            =   420
         TabIndex        =   5
         Top             =   2610
         Width           =   2955
      End
      Begin VB.CheckBox chkCopy 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Create CD Direcotries in the following location"
         Height          =   315
         Left            =   150
         TabIndex        =   4
         Top             =   540
         Value           =   1  'Checked
         Width           =   3945
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
      TabIndex        =   3
      Top             =   4050
      Visible         =   0   'False
      Width           =   8205
   End
End
Attribute VB_Name = "frmYearbook"
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
        Case 1                                              'Output Images to CD Format
            lblStatus.Visible = True
            OutputImages
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
Private Sub OutputImages()                                   'Remove leading 0's from scanned file names
    Dim sSrcName As String
    Dim sTargetDir As String
    Dim sDstName As String
    Dim sExt As String
    Dim sTxt As String
    Dim fRef As FileSystemObject                            'Cross reference file system handle
    Dim fl As TextStream
    Dim fIndex As TextStream
    Dim iFolder As Integer
    Dim iNumFolders As Integer
    Dim iImage As Integer
    
    On Error GoTo ErrorHandler                              'Set up error handler.
    
    Set fRef = New FileSystemObject
    Set fl = fRef.OpenTextFile(frmYearbook.dirCopy.Path & "\LOG.TXT", ForWriting, True, TristateFalse)
    Set fIndex = fRef.OpenTextFile(frmYearbook.dirCopy.Path & "\INDEX.TXT", ForWriting, True, TristateFalse)
    
    
    sExt = Right(frmMain.filSource.FileName, 4)             'Get the extension on the currently selected image file
        
    If MsgBox("Produce CD folders from [" & Trim$(Str$(cData.rsRecords.RecordCount)) & "] images?", vbYesNo + vbApplicationModal + vbQuestion, EZ_CAPTION) = vbNo Then
        Exit Sub                                            'Exit this routine
    End If
        
    '---- Create folders for copy
    lblStatus.Caption = "Creating folders..."
    DoEvents
    iNumFolders = (cData.rsRecords.RecordCount / 200)       '200 images max per folder
    If iNumFolders < 1 Then
        iNumFolders = 1
    End If
    For iFolder = 1 To iNumFolders                          'Loop for each folder to be created
        sTargetDir = frmYearbook.dirCopy.Path & "\FOLDER" & Trim$(Str(iFolder))
        
       ' MsgBox sTargetDir
        
        If Not fRef.FolderExists(sTargetDir) Then
            fRef.CreateFolder sTargetDir
        End If
    Next
        
    '---- Copy images to folders
    lblStatus.Caption = "Copying images..."
    DoEvents
    iFolder = 1
    iImage = 1
    cData.rsRecords.MoveFirst                           'Move to first record in data set
    Do While Not cData.rsRecords.EOF                    'Loop for each record in data set
         
        If Len(Trim(cData.rsRecords(Trim(frmMain.dbcImageTag.Text)).Value)) > 0 Then
            sSrcName = Trim(frmMain.dirSource.Path) & "\" & Trim(cData.rsRecords(Trim(frmMain.dbcImageTag.Text)).Value) & sExt
            sDstName = Trim(dirCopy.Path) & "\FOLDER" & Trim$(Str(iFolder)) & "\" & Trim(cData.rsRecords(Trim(frmMain.dbcImageTag.Text)).Value) & sExt
            If fRef.FileExists(sSrcName) Then
                fRef.CopyFile sSrcName, sDstName, True
                If Not fRef.FileExists(sDstName) Then
                    fl.WriteLine "*** Error copying [" & sSrcName & "] to [" & sDstName & "] filesystem error."
                Else
                    fl.WriteLine "*** Copied [" & sSrcName & "] to [" & sDstName & "] OK."
                    
                    sTxt = "VOL1" & vbTab & "FOLDER" & Trim(Str(iFolder)) & vbTab
                    sTxt = sTxt & Trim(cData.rsRecords(Trim(frmMain.dbcImageTag.Text)).Value) & sExt & vbTab
                    sTxt = sTxt & Trim(cData.rsRecords("GRADE").Value & "") & vbTab
                    sTxt = sTxt & Trim(cData.rsRecords("LAST_NAME").Value & "") & vbTab
                    sTxt = sTxt & Trim(cData.rsRecords("FIRST_NAME").Value & "") & vbTab
                    sTxt = sTxt & Trim(cData.rsRecords("HOMEROOM").Value & "") & vbTab & vbTab
                    sTxt = sTxt & Trim(cData.rsRecords("TEACHER").Value & "") & vbTab
                    
                    fIndex.WriteLine sTxt
                    
                    iImage = iImage + 1
                    If iImage > 200 Then
                        iFolder = iFolder + 1
                        iImage = 1
                    End If
                    
                End If
            Else
                fl.WriteLine "*** Error copying [" & sSrcName & "] to [" & sDstName & "] Source not found."
            End If
        End If
        cData.rsRecords.MoveNext
    
    Loop
    
    fl.Close
    fIndex.Close
    cData.rsRecords.MoveFirst                               'Move to first record in data set
    MsgBox "CD image creation complete.", vbOKOnly + vbApplicationModal + vbInformation, EZ_CAPTION
    Exit Sub
ErrorHandler:
    MsgBox "Error in CD creation: #[" & Str(Err.Number) & "][" & Err.Description & "]"
    Resume Next
End Sub
