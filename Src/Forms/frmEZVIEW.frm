VERSION 5.00
Begin VB.Form frmEZVIEW 
   BackColor       =   &H00C0E0FF&
   Caption         =   "RSC EZ-VIEW CD-ROM Plug-In"
   ClientHeight    =   5760
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5580
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   5580
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraOptions 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Options"
      Height          =   1185
      Left            =   60
      TabIndex        =   11
      Top             =   3780
      Width           =   5475
      Begin VB.CheckBox chkOptionValidate 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Validate Image Files"
         Height          =   315
         Left            =   120
         TabIndex        =   14
         Top             =   840
         Width           =   3705
      End
      Begin VB.CheckBox chkOptionID 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Copy Subject ID to Student ID"
         Height          =   315
         Left            =   120
         TabIndex        =   13
         Top             =   510
         Value           =   1  'Checked
         Width           =   3705
      End
      Begin VB.CheckBox chkOptionReference 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Create Cross Reference Text File"
         Height          =   285
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Value           =   1  'Checked
         Width           =   2745
      End
   End
   Begin VB.Frame fraTarget 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Target Location"
      ForeColor       =   &H00C00000&
      Height          =   3705
      Left            =   2820
      TabIndex        =   7
      Top             =   30
      Width           =   2715
      Begin VB.DriveListBox drvTarget 
         Height          =   315
         Left            =   60
         TabIndex        =   10
         Top             =   210
         Width           =   2580
      End
      Begin VB.FileListBox filTarget 
         Height          =   1650
         Left            =   60
         TabIndex        =   9
         Top             =   1980
         Width           =   2595
      End
      Begin VB.DirListBox dirTarget 
         Height          =   1440
         Left            =   60
         TabIndex        =   8
         Top             =   540
         Width           =   2565
      End
   End
   Begin VB.Frame fraSource 
      BackColor       =   &H00C0E0FF&
      Caption         =   "EZ-VIEW Source Location"
      ForeColor       =   &H00C00000&
      Height          =   3705
      Left            =   60
      TabIndex        =   3
      Top             =   30
      Width           =   2715
      Begin VB.DirListBox dirSource 
         Height          =   1440
         Left            =   60
         TabIndex        =   6
         Top             =   540
         Width           =   2565
      End
      Begin VB.FileListBox filSource 
         Height          =   1650
         Left            =   60
         TabIndex        =   5
         Top             =   1980
         Width           =   2595
      End
      Begin VB.DriveListBox drvSource 
         Height          =   315
         Left            =   60
         TabIndex        =   4
         Top             =   210
         Width           =   2580
      End
   End
   Begin VB.CommandButton cmdEZVIEW 
      Height          =   705
      Index           =   0
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Exit"
      Top             =   5010
      Width           =   825
   End
   Begin VB.CommandButton cmdEZVIEW 
      Height          =   705
      Index           =   1
      Left            =   885
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Batch Process Images"
      Top             =   5010
      Width           =   825
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
      Left            =   1770
      TabIndex        =   2
      Top             =   5370
      Visible         =   0   'False
      Width           =   3705
   End
End
Attribute VB_Name = "frmEZVIEW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------
'
' Project......: RSC EZ-IMAGE(r)
'
' Component....: frmEZVIEW
'
' Procedure....: (Declarations)
'
' Description..: RSC EZ-VIEW Plug-In
'
' Author.......: Ronald D. Redmer
'
' History......: 07-01-97 RDR Designed and Programmed
'
' (c) 1997-1999 Redmer Software Company, Inc.
' All Rights Reserved
'----------------------------------------------------------------------------
Option Explicit                                             'Require explicit variable declaration
Private Const conExitButton = 0                             'Index of exit button
Private Const conProcessButton = 1                          'Index of processing images button
Private Const conRSCIcon = 101                              'Resource ID of RSC Icon
Private Const conStopButtonIcon = 110                       'Resource ID of Stop Icon
Private Const conProcessButtonIcon = 105                    'Resource ID of Process Icon
Private Const conExitButtonIcon = 112                       'Resource ID of Exit Button
Private Const conReportFolder = "\REPORTS\"                 'EZ-VIEW standard report folder
Private Const conDataFolder = "\DATA\"                      'EZ-VIEW standard report folder
Private Const conSetupFolder = "\SETUP\"                    'EZ-VIEW standard setup folder
Private Const conReferenceFolder = "\DATAMAC"               'EZ-VIEW standard reference folder for SASI
Private Const conSASIfile = "XREFPICT.TXT"                  'EZ-VIEW standard SASI file
Private Const conFoxProDSN = "DSN=Visual FoxPro Tables;UID=;PWD=;SourceType=DBF;Exclusive=No;BackgroundFetch=Yes;Collate=Machine;Null=No;Deleted=Yes;SourceDB="
Private bCancel As Boolean                                  'Cancel flag
Private bProcessing As Boolean                              'Set active when image processing loop running
Private sFileName As String                                 'Name of the currently selected image file
Private Sub Form_Load()                                     'Load the form and initialize controls
    On Error GoTo ErrorHandler                              'Set error handler
    Me.Icon = LoadResPicture(conRSCIcon, vbResIcon)         'Load the window icon
    cmdEZVIEW(conExitButton).Picture = LoadResPicture(conExitButtonIcon, vbResIcon)
    cmdEZVIEW(conProcessButton).Picture = LoadResPicture(conProcessButtonIcon, vbResBitmap)
    bCancel = False                                         'Initialize batch process cancel flag to false
    Exit Sub                                                'Exit this routine
ErrorHandler:                                               'Error handling code
    Resume Next                                             'Simply resume next line of code
End Sub
Private Sub cmdEZVIEW_Click(Index As Integer)               'Process command button clicks
    Select Case Index                                       'Select on the index of the button clicked
        Case conExitButton                                  'Exit button
            Unload Me                                       'Unload this form
        Case conProcessButton                               'Bacth process images button
            If Not bProcessing Then                         'Toggle Icons
                If MsgBox("Create EZ-VIEW Distribution?", vbApplicationModal + vbQuestion + vbYesNo, "RSC EZ-VIEW Plug-In") = vbYes Then
                    lblStatus.Visible = True                'Turn status text on
                    cmdEZVIEW(conProcessButton).Picture = LoadResPicture(conStopButtonIcon, vbResIcon)
                    CopySetup                               'Build the EZ-VIEW CD-ROM
                    CopyData                                'Copy data from source table to students table
                    cmdEZVIEW(conProcessButton).Picture = LoadResPicture(conProcessButtonIcon, vbResBitmap)
                    lblStatus.Visible = False               'Trun status text off
                End If
            Else
                bCancel = True
            End If
    End Select
End Sub
Private Sub CopySetup()                                     'Copy setup folders from source to target
    Dim fs As FileSystemObject                              'FileSystem object provided by "Microsoft Scripting Runtime"
    On Error GoTo ErrorHandler                              'Set error handler
    lblStatus.Caption = "Copying files..."                  'Update status
    Set fs = New FileSystemObject                           'Instantiate the object
    DoEvents
    fs.CopyFolder dirSource.Path, dirTarget.Path, True      'Copy folders (with overwrite set)
    DoEvents
    Set fs = Nothing                                        'Clear the filesystem object
    Exit Sub                                                'Exit this routine
ErrorHandler:                                               'Error handling code
    Resume Next                                             'Simply resume next line of code
End Sub
Private Sub CopyData()
    Dim cnnEZ As ADODB.Connection                           'Connection to EZ-VIEW Database
    Dim rsStudents As ADODB.Recordset                       'Students recordset
    Dim fRef As FileSystemObject                            'Cross reference file system handle
    Dim fl As TextStream
    On Error GoTo ErrorHandler                              'Set error handler
    lblStatus.Caption = "Appending Data..."                 'Set status text to appending data
    Set cnnEZ = New ADODB.Connection                        'Instantiate the connection object
    Set rsStudents = New ADODB.Recordset                    'Instantiate the recordset object
    Set fRef = New FileSystemObject
    cnnEZ.Open conFoxProDSN & dirTarget.Path & conDataFolder & ";"  'Open the target database using DSN-less ODBC connection
    rsStudents.Open "SELECT * FROM STUDENTS", cnnEZ, adOpenDynamic, adLockOptimistic
    cData.rsRecords.MoveFirst                               'Move to first record in records table
    If chkOptionReference.Value = 1 Then                    'If build reference file is checked
        fRef.CreateFolder dirTarget.Path & conReferenceFolder
        Set fl = fRef.OpenTextFile(dirTarget.Path & conReferenceFolder & "\" & conSASIfile, ForAppending, True, TristateFalse)
    End If
    Do While Not cData.rsRecords.EOF                        'Loop for each record in the recordset
        If chkOptionID.Value = 1 Then                       'If copy subject to student id is checked
            If Len(Trim$(cData.rsRecords("STUDENTID").Value)) = 0 Then      'Copy the subject field to the student id
                cData.rsRecords("STUDENTID").Value = Trim$(cData.rsRecords("SUBJECT").Value)
                cData.rsRecords.Update
            End If
        End If
        rsStudents.AddNew                                   'Add a new record to the EZ-VIEW Students table
        rsStudents("STUDENTID").Value = Trim$(cData.rsRecords("STUDENTID").Value)
        rsStudents("FIRST_NAME").Value = Trim$(cData.rsRecords("FIRST_NAME").Value)
        rsStudents("LAST_NAME").Value = Trim$(cData.rsRecords("LAST_NAME").Value)
        rsStudents("GRADE").Value = Trim$(cData.rsRecords("GRADE").Value)
        rsStudents("TEACHER").Value = Trim$(cData.rsRecords("TEACHER").Value)
        rsStudents("HOMEROOM").Value = Trim$(cData.rsRecords("HOMEROOM").Value)
        rsStudents("BOX").Value = Trim$(cData.rsRecords("BOX").Value)
        rsStudents("ADDRESS1").Value = Trim$(cData.rsRecords("ADDRESS1").Value)
        rsStudents("ADDRESS2").Value = Trim$(cData.rsRecords("ADDRESS2").Value)
        rsStudents("CITY").Value = Trim$(cData.rsRecords("CITY").Value)
        rsStudents("ZIP_CODE").Value = Trim$(cData.rsRecords("ZIP_CODE").Value)
        rsStudents("PHONE1").Value = Trim$(cData.rsRecords("HOME_PHONE").Value)
        rsStudents("GENDER").Value = Trim$(cData.rsRecords("GENDER").Value)
        rsStudents.Update                                   'Update the EZ-VIEW table with new values
        If chkOptionReference.Value = 1 Then                'If build reference file is checked
            fl.WriteLine Chr$(34) & Format$(rsStudents("STUDENTID").Value, "0000000000") & Chr$(34) & "," & Chr$(34) & Trim$(rsStudents("STUDENTID").Value) & ".PCT" & Chr$(34)
        End If
        cData.rsRecords.MoveNext                            'Move to the next record in the source recordset
        DoEvents
    Loop
    If chkOptionReference.Value = 1 Then                    'If build reference file is checked
        fl.Close                                            'Close the reference file
        Set fl = Nothing                                    'Release the reference file memory
        Set fRef = Nothing                                  'Release the filesystem memory
    End If
    rsStudents.Close                                        'Close the students table
    cnnEZ.Close                                             'Close the connection
    Set rsStudents = Nothing                                'Release students object
    Set cnnEZ = Nothing                                     'Release connection object
    Exit Sub                                                'Exit this routine
ErrorHandler:                                               'Error handling code
    Resume Next                                             'Simply resume next line of code
End Sub
Private Sub drvSource_change()
    On Error GoTo ErrorHandler
    dirSource.Path = drvSource.Drive
ErrorHandler:
    Resume Next
End Sub
Private Sub dirSource_Change()
    On Error GoTo ErrorHandler
    filSource.Path = dirSource.Path
ErrorHandler:
    Resume Next
End Sub
Private Sub drvTarget_Change()
    On Error GoTo ErrorHandler
    dirTarget.Path = drvTarget.Drive
ErrorHandler:
    Resume Next
End Sub
Private Sub dirTarget_Change()
    On Error GoTo ErrorHandler
    filTarget.Path = dirTarget.Path
ErrorHandler:
    Resume Next
End Sub

