VERSION 5.00
Object = "{00100003-B1BA-11CE-ABC6-F5B2E79D9E3F}#1.0#0"; "LTOCX10N.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl ctlProcess 
   ClientHeight    =   8550
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10845
   ScaleHeight     =   8550
   ScaleWidth      =   10845
   Begin VB.Frame frmCVTOptions 
      Caption         =   "Image Processing Options"
      ForeColor       =   &H00C00000&
      Height          =   7605
      Left            =   60
      TabIndex        =   0
      Top             =   30
      Width           =   4215
      Begin VB.DirListBox dirCVT_HR_Folder 
         Height          =   2115
         Left            =   60
         TabIndex        =   26
         Top             =   4005
         Width           =   4035
      End
      Begin VB.DriveListBox drvCVT_HR_Drive 
         Height          =   315
         Left            =   60
         TabIndex        =   25
         Top             =   3675
         Width           =   4035
      End
      Begin VB.TextBox txtCVT_HR_Size 
         Height          =   300
         Left            =   1485
         TabIndex        =   23
         Text            =   "100"
         Top             =   225
         Width           =   540
      End
      Begin VB.TextBox txtCVT_Rotation_Angle 
         Height          =   300
         Left            =   1485
         TabIndex        =   10
         Text            =   "0"
         Top             =   840
         Width           =   540
      End
      Begin VB.TextBox txtCVT_Sharpen_Factor 
         Height          =   300
         Left            =   1485
         TabIndex        =   9
         Text            =   "0"
         Top             =   1155
         Width           =   540
      End
      Begin VB.TextBox txtCVT_Contrast 
         Height          =   285
         Left            =   1485
         TabIndex        =   8
         Text            =   "0"
         Top             =   1470
         Width           =   540
      End
      Begin VB.TextBox txtCVT_Gamma 
         Height          =   300
         Left            =   1485
         TabIndex        =   7
         Text            =   "1.0"
         Top             =   1785
         Width           =   540
      End
      Begin VB.CheckBox chkCVT_Deskew 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1485
         TabIndex        =   6
         Top             =   2085
         Width           =   240
      End
      Begin VB.CheckBox chkCVT_Despeckle 
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
         Left            =   1485
         TabIndex        =   5
         Top             =   2340
         Width           =   270
      End
      Begin VB.CheckBox chkCVT_Flip 
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
         Left            =   1485
         TabIndex        =   4
         Top             =   2610
         Width           =   270
      End
      Begin VB.CheckBox chkCVT_Invert 
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
         Left            =   1485
         TabIndex        =   3
         Top             =   2880
         Width           =   270
      End
      Begin VB.CheckBox chkCVT_Stretch_Intensity 
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
         Left            =   1485
         TabIndex        =   2
         Top             =   3150
         Width           =   270
      End
      Begin VB.TextBox txtCVT_Crop_Factor 
         Height          =   300
         Left            =   1485
         TabIndex        =   1
         Top             =   525
         Width           =   540
      End
      Begin MSDataListLib.DataCombo dbcHRTypes 
         Height          =   315
         Left            =   60
         TabIndex        =   27
         Top             =   6345
         Width           =   4035
         _ExtentX        =   7117
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Save As:"
         Height          =   180
         Left            =   60
         TabIndex        =   30
         Top             =   6150
         Width           =   855
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Save In:"
         Height          =   285
         Left            =   60
         TabIndex        =   29
         Top             =   3480
         Width           =   1440
      End
      Begin VB.Label Label25 
         BackStyle       =   0  'Transparent
         Caption         =   "Resize Image"
         Height          =   285
         Left            =   60
         TabIndex        =   24
         Top             =   255
         Width           =   1080
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Rotation Angle"
         Height          =   285
         Left            =   60
         TabIndex        =   20
         Top             =   870
         Width           =   1440
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Sharpen"
         Height          =   285
         Left            =   60
         TabIndex        =   19
         Top             =   1185
         Width           =   1440
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Contrast"
         Height          =   285
         Left            =   60
         TabIndex        =   18
         Top             =   1485
         Width           =   1440
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Deskew"
         Height          =   285
         Left            =   60
         TabIndex        =   17
         Top             =   2100
         Width           =   1440
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Despeckle"
         Height          =   285
         Left            =   60
         TabIndex        =   16
         Top             =   2385
         Width           =   1425
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Gamma Correct"
         Height          =   285
         Left            =   60
         TabIndex        =   15
         Top             =   1800
         Width           =   1350
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Flip Image"
         Height          =   285
         Left            =   60
         TabIndex        =   14
         Top             =   2655
         Width           =   1425
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Invert Image"
         Height          =   285
         Left            =   60
         TabIndex        =   13
         Top             =   2925
         Width           =   1425
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Stretch Intensity"
         Height          =   285
         Left            =   60
         TabIndex        =   12
         Top             =   3195
         Width           =   1440
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Crop Image"
         Height          =   285
         Left            =   60
         TabIndex        =   11
         Top             =   555
         Width           =   1440
      End
   End
   Begin MSComctlLib.Toolbar tlbProcess 
      Height          =   780
      Left            =   60
      TabIndex        =   21
      Top             =   7710
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   1376
      ButtonWidth     =   1191
      ButtonHeight    =   1376
      Appearance      =   1
      Style           =   1
      ImageList       =   "imgProcess"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Process"
            ImageIndex      =   1
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgProcess 
      Left            =   870
      Top             =   7920
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
            Picture         =   "ctlProcess.ctx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlProcess.ctx":031A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin LEADLib.LEAD ledHR 
      Height          =   7530
      Left            =   4320
      TabIndex        =   28
      Top             =   105
      Width           =   6435
      _Version        =   65537
      _ExtentX        =   11351
      _ExtentY        =   13282
      _StockProps     =   229
      Appearance      =   1
      ScaleHeight     =   498
      ScaleWidth      =   425
      DataField       =   ""
      BitmapDataPath  =   ""
      AnnDataPath     =   ""
      PanWinTitle     =   "PanWindow"
      CLeadCtrl       =   0
   End
   Begin VB.Label lblStatus 
      BorderStyle     =   1  'Fixed Single
      Height          =   405
      Left            =   840
      TabIndex        =   22
      Top             =   8100
      Visible         =   0   'False
      Width           =   9825
   End
End
Attribute VB_Name = "ctlProcess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'----------------------------------------------------------------------------
'
' Project......: RSC EZ-IMAGE(r)
'
' Component....: ctlProcess
'
' Procedure....: (Declarations)
'
' Description..: Process Images (LeadTools Primitives)
'
' Author.......: Ronald D. Redmer
'
' History......: 06-01-00 RDR Designed and Programmed
'
' (c) 1997-2000 Redmer Software Company, Inc.
' All Rights Reserved
'----------------------------------------------------------------------------
Option Explicit
Private sFileName As String
Private rsImageTypesHR As ADODB.Recordset                    'Bindable Recordset of LEAD image types (high resolution)
Private bProcessing As Boolean
Private bCancel As Boolean

Private Sub UserControl_Initialize()
    On Error Resume Next
    Set rsImageTypesHR = New ADODB.Recordset                'Create new recordset object
    BuildImageTypes
    With dbcHRTypes                                         'Setup the high resolution image type combo box
        Set .RowSource = rsImageTypesHR                         'The data comes from the frmmain.usrdata data source class
        .ListField = "DESCRIPTION"                          'Set the list field to description
        .BoundColumn = "TYPE_ID"                            'Bind the control to the image type id
        .DataField = "TYPE_ID"
        .Refresh                                            'Refresh the control
    End With
End Sub

Private Sub UserControl_Terminate()
    On Error Resume Next
    If rsImageTypesHR.State = adStateOpen Then              'If the recordset is already open
        rsImageTypesHR.Close                                'Close the recordset to prevent Open failure
    End If
    Set rsImageTypesHR = Nothing
End Sub

Private Sub tlbProcess_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error Resume Next
    If Not bProcessing Then
        Convert
    Else
        If MsgBox("Stop processing images?", vbQuestion + vbYesNo + vbApplicationModal + vbDefaultButton2, EZ_CAPTION) = vbYes Then
            bCancel = True
        End If
    End If
End Sub

Private Sub Convert()                                       'Batch process images
    Dim iCnt As Integer                                     'Simple image counter
    Dim iFileCnt As Integer                                 'Number of files to process
    On Error Resume Next
    
    iFileCnt = frmMain.UsrImage.ImageCount                  'Get the number of files to process
    If iFileCnt < 1 Then                                    'If no files in source folder
        MsgBox "No files found in source folder to convert.", vbOKOnly + vbInformation + vbApplicationModal, EZ_CAPTION
        Exit Sub                                            'Exit the routine
    End If
    If MsgBox("Convert " & Trim$(Str$(iFileCnt)) & " images?", vbYesNo + vbQuestion + vbApplicationModal, EZ_CAPTION) = vbNo Then
        Exit Sub                                            'Exit the routine
    End If
    bProcessing = True
    lblStatus.Visible = True                                'Show status text
    tlbProcess.Buttons(1).Image = 2
    bCancel = False                                         'Initialize cancel flag to false
    For iCnt = 0 To iFileCnt - 1                            'Loop for each file in source directory
        frmMain.UsrImage.ImageNumber = iCnt
        sFileName = frmMain.UsrImage.ImageFileName
        lblStatus.Caption = "Processing [" & Str(iCnt + 1) & "]/[" & Str(iFileCnt) & "] Name=[" & sFileName & "]."
        EZ_Convert_Image                                    'Convert the image
        ledHR.Save Trim$(dirCVT_HR_Folder.Path & "\" & Mid$(frmMain.UsrImage.ImageShortFileName, 1, Len(frmMain.UsrImage.ImageShortFileName) - 4)) & "." & Mid$(dbcHRTypes.Text, 2, 3), CInt(dbcHRTypes.BoundText), 24, QFACTOR_PQ1, SAVE_OVERWRITE
        DoEvents                                            'Process windows events
        If bCancel = True Then                              'If the user hit the cancel button
            Exit For                                        'Exit the image processing loop
        End If
    Next
    DoEvents
    bProcessing = False
    lblStatus.Visible = False
    tlbProcess.Buttons(1).Image = 1
    MsgBox "Finished Processing!", vbOKOnly + vbInformation + vbApplicationModal, EZ_CAPTION
End Sub

Public Sub EZ_Convert_Image()                               'Batch process images
    Dim dScaleHR As Double
    Dim dCropFactor As Double
    Dim iCropPosition As Integer
    On Error Resume Next
    ledHR.Load sFileName, 0, 0, 1                           'Load hi res control
    ledHR.ScaleMode = 3                                     'Set scalemode to pixels
    ledHR.AutoScroll = True                                 'Set scroll bars
    ledHR.AutoSetRects = True                               'Make sure the display rectangles are adjusted
    dScaleHR = Val(txtCVT_HR_Size.Text)                 'Convert scale for hi res
    dCropFactor = Val(txtCVT_Crop_Factor.Text)          'Convert scale for low res
    If dScaleHR > 0 Then
        ledHR.Size CInt((dScaleHR / 100) * frmMain.UsrImage.ImageWidth), CInt((dScaleHR / 100) * frmMain.UsrImage.ImageHeight), RESIZE_RESAMPLE
    End If
    If dCropFactor > 0 Then
        ledHR.SetDstClipRect 0, 0, CInt((dCropFactor / 100) * frmMain.UsrImage.ImageWidth), CInt((dCropFactor / 100) * frmMain.UsrImage.ImageHeight)
    End If
    ledHR.Rotate Val(txtCVT_Rotation_Angle.Text) * 100, True, RGB(0, 0, 0)
    ledHR.Sharpen Val(txtCVT_Sharpen_Factor.Text)       'Sharpen
    ledHR.Contrast Val(txtCVT_Contrast.Text)            'Contrast
    If Val(txtCVT_Gamma.Text) > 0 Then                  'Gamma
        ledHR.GammaCorrect Val(txtCVT_Gamma.Text) * 100
    End If
    If chkCVT_Invert.Value Then ledHR.Invert
    If chkCVT_Despeckle.Value Then ledHR.Despeckle
    If chkCVT_Deskew.Value Then ledHR.Deskew
    If chkCVT_Flip.Value Then ledHR.Flip
    If chkCVT_Stretch_Intensity.Value Then ledHR.StretchIntensity
    ledHR.ForceRepaint                                      'Repaint High res image
End Sub

Private Sub drvCVT_HR_Drive_Change()
    On Error Resume Next
    dirCVT_HR_Folder.Path = drvCVT_HR_Drive.Drive
End Sub

Private Sub txtCVT_Contrast_LostFocus()
    EZ_Convert_Image
End Sub
Private Sub txtCVT_Crop_Factor_LostFocus()
    EZ_Convert_Image
End Sub
Private Sub txtCVT_Gamma_LostFocus()
    EZ_Convert_Image
End Sub
Private Sub txtCVT_HR_Size_LostFocus()
    EZ_Convert_Image
End Sub
Private Sub txtCVT_Rotation_Angle_LostFocus()
    EZ_Convert_Image
End Sub
Private Sub txtCVT_Sharpen_Factor_LostFocus()
    EZ_Convert_Image
End Sub
Private Sub chkCVT_Deskew_Click()
    EZ_Convert_Image
End Sub
Private Sub chkCVT_Despeckle_Click()
    EZ_Convert_Image
End Sub
Private Sub chkCVT_Flip_Click()
    EZ_Convert_Image
End Sub
Private Sub chkCVT_Invert_Click()
    EZ_Convert_Image
End Sub
Private Sub chkCVT_Stretch_Intensity_Click()
    EZ_Convert_Image
End Sub

Private Sub BuildImageTypes()

    On Error Resume Next
    With rsImageTypesHR
        .Fields.Append "TYPE_NO", adInteger, 2              'Append a column to hold the image type number
        .Fields.Append "TYPE_ID", adInteger, 2              'Append a column to hold the image type number
        .Fields.Append "DESCRIPTION", adBSTR, 80            'Append a column to hold the type description
        .CursorLocation = adUseClient
        .CursorType = adOpenKeyset                          'Set cursor type to dynamic to allow additions on the fly
        .LockType = adLockOptimistic                        'Set lock type to optimistic - very low chance of contention
        .Open                                               'Open the recordset
        AddIt 1, FILE_CMP, "[CMP] LEAD Format"
        AddIt 2, FILE_JFIF, "[JPG] JPEG File Interchange Format with YUV 4:4:4"
        AddIt 3, FILE_LEAD2JFIF, "[JPG] JPEG File Interchange Format with YUV 4:2:2"
        AddIt 4, FILE_LEAD1JFIF, "[JPG] JPEG File Interchange Format with YUV 4:1:1"
        AddIt 5, FILE_JTIF, "[JPG] TIF with JPEG compression and YUV 4:4:4"
        AddIt 6, FILE_LEAD2JTIF, "[JPG] TIF with JPEG compression and YUV 4:2:2"
        AddIt 7, FILE_LEAD1JTIF, "[JPG] TIF with JPEG compression and YUV 4:1:1"
        AddIt 8, FILE_GIF, "[GIF] Compuserve GIF"
        AddIt 9, FILE_TIFLZW, "[TIF] TIF with LZW compression and RGB"
        AddIt 10, FILE_TIFLZW_CMYK, "[TIF] TIF with LZW compression and CMYK"
        AddIt 11, FILE_TIFLZW_YCC, "[TIF] TIF with LZW compression and YCbCr"
        AddIt 12, FILE_TIF, "[TIF] TIF uncompressed with RGB"
        AddIt 13, FILE_TIF_CMYK, "[TIF] TIF uncompressed with CMYK"
        AddIt 14, FILE_TIF_YCC, "[TIF] TIF uncompressed with YCbCr"
        AddIt 15, FILE_TIF_PACKBITS, "[TIF] TIF with PACKBITS compression RGB"
        AddIt 16, FILE_TIF_PACKBITS_CMYK, "[TIF] TIF with PACKBITS compression YCcbCr"
        AddIt 17, FILE_TIF_PACKBITS_YCC, "[TIF] TIF with PACKBITS compression CMYK"
        AddIt 18, FILE_BMP, "[BMP] Windows BMP uncompressed"
        AddIt 19, FILE_BMP_RLE, "[BMP] Windows BMP with RLE compression"
        AddIt 20, FILE_OS2, "[OS2] OS/2 BMP version 1.x"
        AddIt 21, FILE_OS2_2, "[OS2] OS/2 BMP version 2.x"
        AddIt 22, FILE_WIN_ICO, "[ICO] Windows ICO icon file"
        AddIt 23, FILE_WIN_CUR, "[CUR] Windows CUR cursor file"
        AddIt 24, FILE_FPX, "[FPX] Kodak FlashPix uncompressed"
        AddIt 25, FILE_FPX_SINGLE_COLOR, "[FPX] Kodak FlashPix single color compression"
        AddIt 26, FILE_FPX_JPEG, "[FPX] Kodak FlashPix compressed with JPEG (med quality)"
        AddIt 27, FILE_FPX_JPEG_QFACTOR, "[FPX] Kodak FlashPix compressed with JPEG (high quality)"
        AddIt 28, FILE_EXIF, "[EXF] Exif file containing TIF uncompressed and RGB"
        AddIt 29, FILE_EXIF_YCC, "[EXF] Exif file containing TIF uncompressed and YCbCr"
        AddIt 30, FILE_EXIF_JPEG, "[EXF] Exif file containing TIF JPEG compressed"
        AddIt 31, FILE_DICOM_GRAY, "[DCM] DICOM grayscale"
        AddIt 32, FILE_DICOM_COLOR, "[DCM] DICOM color (RGB)"
        AddIt 33, FILE_PCX, "[PCX] PCX Zsoft (Windows Paintbrush)"
        AddIt 34, FILE_WMF, "[WMF] WMF Windows Metafile"
        AddIt 35, FILE_PSD, "[PSD] PSD Adobe PhotoShop 3.x"
        AddIt 36, FILE_PNG, "[PNG] PNG Portable Network Graphics Format"
        AddIt 37, FILE_TGA, "[TGA] TGA Truevision TARGA"
        AddIt 38, FILE_EPS, "[EPS] EPS Encapsulated Postscript"
        AddIt 39, FILE_RAS, "[RAS] RAS Sun Microsystems Raster Format"
        AddIt 40, FILE_WPG, "[WPG] WPG Corel WordPerfect Graphics"
        AddIt 41, FILE_PCT, "[PCT] PCT Apple Macintosh PICT (JPEG compressed)"
        AddIt 42, FILE_CCITT, "[CCT] TIF compressed using CCITT"
        AddIt 43, FILE_CCITT_GROUP3_1DIM, "[CCT] TIF compressed using CCITT group 3 1 dimension"
        AddIt 44, FILE_CCITT_GROUP3_2DIM, "[CCT] TIF compressed using CCITT group 3 2 dimensions"
        AddIt 45, FILE_CCITT_GROUP4, "[CCT] TIF compressed using CCITT group 4"
        AddIt 46, FILE_FAX_G3_1D, "[FAX] Raw FAX compressed using CCITT group 3 1 dimension"
        AddIt 47, FILE_FAX_G3_2D, "[FAX] Raw FAX compressed using CCITT group 3 2 dimension"
        AddIt 48, FILE_FAX_G4, "[FAX] Raw FAX compressed using CCITT group 4"
        AddIt 49, FILE_WFX_G3_1D, "[FAX] Winfax compressed using CCITT group 3 1 dimension"
        AddIt 50, FILE_WFX_G4, "[FAX] Winfax compressed using CCITT group 3 2 dimensions"
        AddIt 51, FILE_ICA_G3_1D, "[ICA] IOCA compressed using CCITT group 3 1 dimension"
        AddIt 52, FILE_ICA_G3_2D, "[ICA] IOCA compressed using CCITT group 3 2 dimensions"
        AddIt 53, FILE_ICA_G4, "[ICA] IOCA compressed using CCITT group 4"
        AddIt 54, FILE_RAWICA_G3_1D, "[ICA] IOCA compressed using CCITT group 3 1 dim no-wrap"
        AddIt 55, FILE_RAWICA_G3_2D, "[ICA] IOCA compressed using CCITT group 3 2 dim no-wrap"
        AddIt 56, FILE_RAWICA_G4, "[ICA] IOCA compressed using CCITT group 4 no-wrap"
        AddIt 57, FILE_CALS, "[CAL] CALS Raster File"
        AddIt 58, FILE_MAC, "[MAC] MacPaint"
        AddIt 59, FILE_MSP, "[MSP] Microsoft Paint"
        AddIt 60, FILE_IMG, "[GEM] GEM Image"
        .Update
    End With

End Sub
Private Function AddIt(iIndex As Integer, iTypeID As Integer, sDescript As String)
    On Error Resume Next
    With rsImageTypesHR
        .AddNew
        .Fields("TYPE_NO").Value = iIndex
        .Fields("TYPE_ID").Value = iTypeID
        .Fields("DESCRIPTION").Value = sDescript
        .Update
    End With
End Function

Public Property Let FileName(ByVal sName As String)
    sFileName = sName
End Property

Public Property Get FileName() As String
    FileName = sFileName
End Property



Public Property Get CropFactor() As Double
    On Error Resume Next
    CropFactor = Val(txtCVT_Crop_Factor.Text)
End Property
Public Property Let CropFactor(dFactor As Double)
    On Error Resume Next
    txtCVT_Crop_Factor.Text = Trim$(Str$(dFactor))
End Property

Public Property Get RotationAngle() As Double
    On Error Resume Next
    RotationAngle = Val(txtCVT_Rotation_Angle.Text)
End Property
Public Property Let RotationAngle(dAngle As Double)
    On Error Resume Next
    txtCVT_Rotation_Angle.Text = Trim$(Str$(dAngle))
End Property

Public Property Get SharpenFactor() As Double
    On Error Resume Next
    SharpenFactor = Val(txtCVT_Sharpen_Factor.Text)
End Property
Public Property Let SharpenFactor(dFactor As Double)
    On Error Resume Next
    txtCVT_Sharpen_Factor.Text = Trim$(Str$(dFactor))
End Property

Public Property Get Contrast() As Integer
    On Error Resume Next
    Contrast = Val(txtCVT_Contrast.Text)
End Property
Public Property Let Contrast(iContrast As Integer)
    On Error Resume Next
    txtCVT_Contrast.Text = Trim$(Str$(iContrast))
End Property

Public Property Get Gamma() As Double
    On Error Resume Next
    Gamma = Val(txtCVT_Gamma.Text)
End Property
Public Property Let Gamma(iGamma As Double)
    On Error Resume Next
    txtCVT_Gamma.Text = Trim$(Str$(iGamma))
End Property

Public Property Get Deskew() As Integer
    On Error Resume Next
    Deskew = chkCVT_Deskew.Value
End Property
Public Property Let Deskew(iDeskew As Integer)
    On Error Resume Next
    chkCVT_Deskew.Value = iDeskew
End Property

Public Property Get Despeckle() As Integer
    On Error Resume Next
    Despeckle = chkCVT_Despeckle.Value
End Property
Public Property Let Despeckle(iDespeckle As Integer)
    On Error Resume Next
    chkCVT_Despeckle.Value = iDespeckle
End Property

Public Property Get Flip() As Integer
    On Error Resume Next
    Flip = chkCVT_Flip.Value
End Property
Public Property Let Flip(iFlip As Integer)
    On Error Resume Next
    chkCVT_Flip.Value = iFlip
End Property

Public Property Get Invert() As Integer
    On Error Resume Next
    Invert = chkCVT_Invert.Value
End Property
Public Property Let Invert(iInvert As Integer)
    On Error Resume Next
    chkCVT_Invert.Value = iInvert
End Property

Public Property Get Stretch_Intensity() As Integer
    On Error Resume Next
    Stretch_Intensity = chkCVT_Stretch_Intensity.Value
End Property
Public Property Let Stretch_Intensity(iStretch_Intensity As Integer)
    On Error Resume Next
    chkCVT_Stretch_Intensity.Value = iStretch_Intensity
End Property

Public Property Get HR_Size() As Integer
    On Error Resume Next
    HR_Size = Val(txtCVT_HR_Size.Text)
End Property
Public Property Let HR_Size(iHR_Size As Integer)
    On Error Resume Next
    txtCVT_HR_Size.Text = Trim$(Str$(iHR_Size))
End Property

Public Property Get ProcessPath() As String
    ProcessPath = IIf(Right(Trim$(dirCVT_HR_Folder.Path), 1) = "\", Trim$(dirCVT_HR_Folder.Path), Trim$(dirCVT_HR_Folder.Path) + "\")
End Property
Public Property Let ProcessPath(sPath As String)
    On Error Resume Next
    Dim fso As Scripting.FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    drvCVT_HR_Drive.Drive = fso.GetDriveName(sPath)
    dirCVT_HR_Folder.Path = sPath
    Set fso = Nothing
End Property

Public Property Get FileType() As String
    On Error Resume Next
    FileType = dbcHRTypes.BoundText
End Property
Public Property Let FileType(sType As String)
    On Error Resume Next
    dbcHRTypes.BoundText = sType
End Property

