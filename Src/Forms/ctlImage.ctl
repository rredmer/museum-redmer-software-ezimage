VERSION 5.00
Object = "{00100003-B1BA-11CE-ABC6-F5B2E79D9E3F}#1.0#0"; "LTOCX10N.OCX"
Begin VB.UserControl ctlImage 
   ClientHeight    =   9345
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2820
   ScaleHeight     =   9345
   ScaleWidth      =   2820
   Begin VB.Frame fraImage 
      Caption         =   "Image Location"
      ForeColor       =   &H00C00000&
      Height          =   9285
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   2715
      Begin VB.DirListBox dirSource 
         Height          =   1440
         Left            =   90
         TabIndex        =   2
         Top             =   540
         Width           =   2535
      End
      Begin VB.FileListBox filSource 
         Height          =   3210
         Left            =   60
         TabIndex        =   3
         Top             =   2010
         Width           =   2595
      End
      Begin VB.DriveListBox drvSource 
         Height          =   315
         Left            =   90
         TabIndex        =   1
         Top             =   210
         Width           =   2520
      End
      Begin LEADLib.LEAD ledSource 
         Height          =   3945
         Left            =   60
         TabIndex        =   4
         Top             =   5250
         Width           =   2565
         _Version        =   65537
         _ExtentX        =   4524
         _ExtentY        =   6959
         _StockProps     =   229
         BackColor       =   12632256
         Appearance      =   1
         ScaleHeight     =   259
         ScaleWidth      =   167
         DataField       =   ""
         BitmapDataPath  =   ""
         AnnDataPath     =   ""
         PanWinTitle     =   "PanWindow"
         CLeadCtrl       =   0
      End
   End
End
Attribute VB_Name = "ctlImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'----------------------------------------------------------------------------
'
' Project......: RSC EZ-IMAGE(r)
'
' Component....: ctlImage
'
' Procedure....: (Declarations)
'
' Description..: Image location
'
' Author.......: Ronald D. Redmer
'
' History......: 07-01-97 RDR Designed and Programmed
'
' (c) 1997-1999 Redmer Software Company, Inc.
' All Rights Reserved
'----------------------------------------------------------------------------
Option Explicit
Private sFileName As String

Private Sub drvSource_change()
    On Error Resume Next
    dirSource.Path = drvSource.Drive
End Sub
Private Sub dirSource_Change()
    On Error Resume Next
    filSource.Path = dirSource.Path
End Sub
Private Sub filSource_Click()
    On Error GoTo ErrorHandler
    sFileName = Trim$(IIf(Right$(dirSource.Path, 1) <> "\", dirSource.Path & "\", dirSource.Path) & Trim$(filSource.FileName))
    ledSource.Load sFileName, 0, 0, 1
    With frmMain
        If .tabMain.Tab = 4 Then
            .UsrProcess.FileName = sFileName
            .UsrProcess.EZ_Convert_Image
        End If
    End With
    Exit Sub
ErrorHandler:
    MsgBox "Error processing image.", vbApplicationModal + vbOKOnly + vbExclamation, EZ_CAPTION
    Resume Next
End Sub

Public Property Get ImagePath() As String
    On Error Resume Next
    ImagePath = IIf(Right$(Trim$(dirSource.Path), 1) <> "\", Trim$(dirSource.Path) & "\", Trim$(dirSource.Path))
End Property
Public Property Let ImagePath(sPath As String)
    Dim fso As Object
    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    drvSource.Drive = fso.GetDriveName(sPath)
    dirSource.Path = sPath
    filSource.Path = sPath
    Set fso = Nothing
End Property
Public Property Get ImageExtension() As String
    On Error Resume Next
    ImageExtension = Right$(filSource.List(1), 4)
End Property
Public Property Get ImageCount() As Integer
    On Error Resume Next
    ImageCount = filSource.ListCount
End Property
Public Property Get ImageAspectRatio() As Double
    On Error Resume Next
    ImageAspectRatio = (ledSource.SrcWidth / ledSource.SrcHeight)
End Property
Public Property Get ImageWidth() As Integer
    On Error Resume Next
    ImageWidth = ledSource.BitmapWidth
End Property
Public Property Get ImageHeight() As Integer
    On Error Resume Next
    ImageHeight = ledSource.BitmapHeight
End Property
Public Property Get ImageShortFileName() As String
    On Error Resume Next
    ImageShortFileName = filSource.FileName
End Property
Public Property Get ImageFileName() As String
    On Error Resume Next
    ImageFileName = sFileName
End Property
Public Property Let ImageFileName(sName As String)
    On Error Resume Next
    sFileName = ImageFileName
End Property
Public Property Get ImageNumber() As Integer
    On Error Resume Next
    ImageNumber = filSource.ListIndex
End Property
Public Property Let ImageNumber(iNum As Integer)
    On Error Resume Next
    filSource.ListIndex = iNum
End Property
