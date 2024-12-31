VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl ctlComposite 
   ClientHeight    =   7515
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8085
   ScaleHeight     =   7515
   ScaleWidth      =   8085
   Begin VB.Frame fraTitle 
      Caption         =   "Title Text Blocks"
      Height          =   2085
      Left            =   5250
      TabIndex        =   32
      Top             =   30
      Width           =   2775
      Begin MSMask.MaskEdBox txtTitleRow 
         Height          =   315
         Index           =   0
         Left            =   570
         TabIndex        =   42
         Top             =   570
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   3
         Format          =   "000"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtTitleRow 
         Height          =   315
         Index           =   1
         Left            =   570
         TabIndex        =   43
         Top             =   900
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   3
         Format          =   "000"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtTitleRow 
         Height          =   315
         Index           =   2
         Left            =   570
         TabIndex        =   44
         Top             =   1230
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   3
         Format          =   "000"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtTitleRow 
         Height          =   315
         Index           =   3
         Left            =   570
         TabIndex        =   45
         Top             =   1560
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   3
         Format          =   "000"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtTitleCol 
         Height          =   315
         Index           =   0
         Left            =   1170
         TabIndex        =   46
         Top             =   570
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   3
         Format          =   "000"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtTitleCol 
         Height          =   315
         Index           =   1
         Left            =   1170
         TabIndex        =   47
         Top             =   900
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   3
         Format          =   "000"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtTitleCol 
         Height          =   315
         Index           =   2
         Left            =   1170
         TabIndex        =   48
         Top             =   1230
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   3
         Format          =   "000"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtTitleCol 
         Height          =   315
         Index           =   3
         Left            =   1170
         TabIndex        =   49
         Top             =   1560
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   3
         Format          =   "000"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtTitleColCount 
         Height          =   315
         Index           =   0
         Left            =   1830
         TabIndex        =   50
         Top             =   570
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   3
         Format          =   "000"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtTitleColCount 
         Height          =   315
         Index           =   1
         Left            =   1830
         TabIndex        =   51
         Top             =   900
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   3
         Format          =   "000"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtTitleColCount 
         Height          =   315
         Index           =   2
         Left            =   1830
         TabIndex        =   52
         Top             =   1230
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   3
         Format          =   "000"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtTitleColCount 
         Height          =   315
         Index           =   3
         Left            =   1830
         TabIndex        =   53
         Top             =   1560
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   3
         Format          =   "000"
         PromptChar      =   "_"
      End
      Begin VB.Label Label30 
         BackStyle       =   0  'Transparent
         Caption         =   "# Columns"
         Height          =   255
         Left            =   1830
         TabIndex        =   40
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label29 
         BackStyle       =   0  'Transparent
         Caption         =   "Row"
         Height          =   255
         Left            =   570
         TabIndex        =   39
         Top             =   360
         Width           =   585
      End
      Begin VB.Label Label28 
         BackStyle       =   0  'Transparent
         Caption         =   "Column"
         Height          =   255
         Left            =   1170
         TabIndex        =   38
         Top             =   360
         Width           =   525
      End
      Begin VB.Label Label31 
         BackStyle       =   0  'Transparent
         Caption         =   "Line"
         Height          =   195
         Left            =   150
         TabIndex        =   37
         Top             =   360
         Width           =   405
      End
      Begin VB.Label Label32 
         BackStyle       =   0  'Transparent
         Caption         =   "1."
         Height          =   195
         Left            =   150
         TabIndex        =   36
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label33 
         BackStyle       =   0  'Transparent
         Caption         =   "2."
         Height          =   195
         Left            =   150
         TabIndex        =   35
         Top             =   930
         Width           =   495
      End
      Begin VB.Label Label34 
         BackStyle       =   0  'Transparent
         Caption         =   "3."
         Height          =   195
         Left            =   150
         TabIndex        =   34
         Top             =   1260
         Width           =   495
      End
      Begin VB.Label Label35 
         BackStyle       =   0  'Transparent
         Caption         =   "4."
         Height          =   195
         Left            =   150
         TabIndex        =   33
         Top             =   1590
         Width           =   495
      End
   End
   Begin VB.Frame fraLAYCaption 
      Caption         =   "Image Caption Options"
      Height          =   1995
      Left            =   60
      TabIndex        =   24
      Top             =   4710
      Width           =   5115
      Begin VB.CheckBox chkCompCaption 
         Caption         =   "Caption Images"
         Height          =   225
         Left            =   150
         TabIndex        =   25
         Top             =   270
         Width           =   2085
      End
      Begin MSDataListLib.DataCombo dcbCompCaption 
         Height          =   315
         Index           =   0
         Left            =   1350
         TabIndex        =   26
         Top             =   540
         Width           =   2985
         _ExtentX        =   5265
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dcbCompCaption 
         Height          =   315
         Index           =   1
         Left            =   1350
         TabIndex        =   27
         Top             =   870
         Width           =   2985
         _ExtentX        =   5265
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin MSMask.MaskEdBox txtCompCaptionOffset 
         Height          =   315
         Left            =   1350
         TabIndex        =   41
         Top             =   1530
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   6
         Format          =   "00.000"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtCompCaptionFontSize 
         Height          =   315
         Left            =   1350
         TabIndex        =   54
         Top             =   1200
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   5
         Format          =   "000.0"
         PromptChar      =   "_"
      End
      Begin VB.Label Label36 
         BackColor       =   &H00C0E0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Caption field 1"
         Height          =   195
         Left            =   120
         TabIndex        =   31
         Top             =   600
         Width           =   1245
      End
      Begin VB.Label Label37 
         BackColor       =   &H00C0E0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Font Size"
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   1275
         Width           =   1155
      End
      Begin VB.Label Label53 
         BackColor       =   &H00C0E0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Caption field 2"
         Height          =   195
         Left            =   120
         TabIndex        =   29
         Top             =   930
         Width           =   1245
      End
      Begin VB.Label Label65 
         BackColor       =   &H00C0E0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Offset (inches)"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   1590
         Width           =   1155
      End
   End
   Begin VB.Frame fraCompPageLayout 
      Caption         =   "Page Layout"
      Height          =   4605
      Left            =   60
      TabIndex        =   1
      Top             =   30
      Width           =   5115
      Begin VB.OptionButton optCompPageSource 
         Caption         =   "Layout from Source Info"
         Height          =   285
         Index           =   0
         Left            =   60
         TabIndex        =   5
         Top             =   210
         Value           =   -1  'True
         Width           =   2235
      End
      Begin VB.OptionButton optCompPageSource 
         Caption         =   "Layout from Folder Info"
         Height          =   315
         Index           =   1
         Left            =   60
         TabIndex        =   4
         Top             =   495
         Width           =   2235
      End
      Begin VB.CheckBox chkCompOvals 
         Caption         =   "Oval Images"
         Height          =   225
         Left            =   90
         TabIndex        =   3
         Top             =   4170
         Width           =   2085
      End
      Begin VB.CheckBox chkCompRowShift 
         Caption         =   "Shift even rows"
         Height          =   225
         Left            =   90
         TabIndex        =   2
         Top             =   3930
         Width           =   2085
      End
      Begin MSMask.MaskEdBox txtCompPageWidth 
         Height          =   315
         Left            =   1560
         TabIndex        =   6
         Top             =   900
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   6
         Format          =   "00.000"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtCompPageHeight 
         Height          =   315
         Left            =   1560
         TabIndex        =   7
         Top             =   1230
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   6
         Format          =   "00.000"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtCompMarginTop 
         Height          =   315
         Left            =   1560
         TabIndex        =   8
         Top             =   1560
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   6
         Format          =   "00.000"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtCompMarginBottom 
         Height          =   315
         Left            =   1560
         TabIndex        =   9
         Top             =   1890
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   6
         Format          =   "00.000"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtCompMarginLeft 
         Height          =   315
         Left            =   1560
         TabIndex        =   10
         Top             =   2220
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   6
         Format          =   "00.000"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtCompMarginRight 
         Height          =   315
         Left            =   1560
         TabIndex        =   11
         Top             =   2550
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   6
         Format          =   "00.000"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtCompWhiteSpace 
         Height          =   315
         Left            =   1560
         TabIndex        =   12
         Top             =   2880
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   6
         Format          =   "00.000"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtCompColumns 
         Height          =   315
         Left            =   1560
         TabIndex        =   13
         Top             =   3210
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   3
         Format          =   "000"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtCompRows 
         Height          =   315
         Left            =   1560
         TabIndex        =   14
         Top             =   3540
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   3
         Format          =   "000"
         PromptChar      =   "_"
      End
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         Caption         =   "White Space %"
         Height          =   255
         Left            =   90
         TabIndex        =   23
         Top             =   2910
         Width           =   1515
      End
      Begin VB.Label Label26 
         BackStyle       =   0  'Transparent
         Caption         =   "Bottom Margin"
         Height          =   255
         Left            =   90
         TabIndex        =   22
         Top             =   1920
         Width           =   1515
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   "Top Margin"
         Height          =   255
         Left            =   90
         TabIndex        =   21
         Top             =   1575
         Width           =   1515
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "Page Height"
         Height          =   255
         Left            =   90
         TabIndex        =   20
         Top             =   1241
         Width           =   1515
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Page Width"
         Height          =   255
         Left            =   90
         TabIndex        =   19
         Top             =   930
         Width           =   1515
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Number of Rows"
         Height          =   255
         Left            =   90
         TabIndex        =   18
         Top             =   3570
         Width           =   1515
      End
      Begin VB.Label lblCompPropCols 
         BackStyle       =   0  'Transparent
         Caption         =   "Number of Columns"
         Height          =   255
         Left            =   90
         TabIndex        =   17
         Top             =   3225
         Width           =   1515
      End
      Begin VB.Label Label46 
         BackStyle       =   0  'Transparent
         Caption         =   "Left Margin"
         Height          =   255
         Left            =   90
         TabIndex        =   16
         Top             =   2250
         Width           =   1515
      End
      Begin VB.Label Label47 
         BackStyle       =   0  'Transparent
         Caption         =   "Right Margin"
         Height          =   255
         Left            =   90
         TabIndex        =   15
         Top             =   2580
         Width           =   1515
      End
   End
   Begin MSComctlLib.ImageList imgComposite 
      Left            =   2250
      Top             =   6780
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   23
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlComposite.ctx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlComposite.ctx":067A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlComposite.ctx":0994
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbComposite 
      Height          =   645
      Left            =   60
      TabIndex        =   0
      Top             =   6780
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   1138
      ButtonWidth     =   1191
      ButtonHeight    =   1138
      Appearance      =   1
      Style           =   1
      ImageList       =   "imgComposite"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Size"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Preview"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Build"
            ImageIndex      =   3
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "ctlComposite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'----------------------------------------------------------------------------
'
' Project......: RSC EZ-IMAGE(r)
'
' Component....: ctlComposite
'
' Procedure....: (Declarations)
'
' Description..: Process Adobe InDesign Composites
'
' Author.......: Ronald D. Redmer
'
' History......: 06-01-00 RDR Designed and Programmed
'
' (c) 1997-2000 Redmer Software Company, Inc.
' All Rights Reserved
'----------------------------------------------------------------------------
Option Explicit
Private sImagePath As String                                'Path to image files
Private sImageTag As String                                 'Field used to tag images
Private sImageExtension As String                           'Extension of images
Private dPageWidth As Double                                'The page width
Private dPageHeight As Double                               'The page height
Private dMarginTop As Double                                'The top margin
Private dMarginBottom As Double                             'The bottom margin
Private dMarginLeft As Double                               'The left margin
Private dMarginRight As Double                              'The right margin
Private dWhiteSpace As Double                               'The amount of white space on the page (%)
Private iPageCols As Integer                                'The number of columns on the page
Private iPageRows As Integer                                'The number of rows on the page
Private iTitleCount As Integer                              'Title block counter
Private iTitleRow(4) As Integer                             'Vertical row of title block
Private iTitleCol(4) As Integer                             'Horizontal column of title block (left)
Private iTitleColCount(4) As Integer                        'Number of columns to skip for title
Private sCaptionField1 As String                            'Caption field
Private sCaptionField2 As String                            'Caption field
Private dCaptionOffset As Double                            'The amount of white space to leave for the caption
Private dCaptionFontSize As Double                          'The font size to use for the caption
Private dAspectRatio As Double                              'The aspect ratio of the images

Private Sub UserControl_Initialize()
    On Error Resume Next
    SetDefaults                                             'Set control defaults
End Sub

Private Sub UserControl_EnterFocus()
    Dim iCaption As Integer                                 'Counter - caption columns
    On Error Resume Next
    '--- Set caption combo-box bindings to data source (app global)
    For iCaption = 0 To 1                                   'Loop for each caption column
        With dcbCompCaption(iCaption)                       'With the caption column data-combo
            Set .RowSource = frmMain.UsrData.rsColumns      'Set the data source to the provider class
            .ListField = "COLUMN"                           'Set the list field
            .BoundColumn = "COLUMN"                         'Set the bound column
            .Refresh                                        'Refresh the control
        End With
    Next
End Sub

Private Sub tlbComposite_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error Resume Next
    sImagePath = Trim$(frmMain.UsrImage.ImagePath)                         'Path to source images
    sImageTag = frmMain.UsrData.ImageTag                                   'Data field ised to tag images
    sImageExtension = frmMain.UsrImage.ImageExtension
    Select Case Button.Index                                            'Select on the index of the button clicked
        Case 1                                                          'Auto size
            AutoSize                                                    'Call auto sizing method
            PutControls                                                 'Update controls
        Case 2                                                          'Preview
            Build 1
        Case 3                                                          'Build
            Build 0
    End Select
End Sub

Public Sub Build(iMode As Integer)
    
    Dim oApp As InDesign.Application                        'Adobe InDesign Application Object
    Dim oDocument As InDesign.Document                      'Adobe InDesign Document Object
    Dim oTextFrame As InDesign.TextFrame                    'Adobe InDesign Text Frame Object
    Dim oWindow As InDesign.Window                          'Adobe InDesign Window Object
    Dim oSpreads As InDesign.Spreads                        'Adobe InDesign Spreads Collection Object
    Dim oSpread As InDesign.Spread                          'Adobe InDesign Spread Object
    Dim oPageItems As InDesign.PageItems                    'Adobe InDesign Page Items Collection Object
    Dim oImage As InDesign.Image                            'Adobe InDesign Image Object
    Dim oRect As InDesign.Rectangle                         'Adobe InDesign Rectangle Object
    Dim oOval As InDesign.Oval                              'Adobe InDesign Oval Object
    Dim oPolygon As InDesign.Polygon                        'Adobe InDesign Polygon Object
    Dim iImages As Integer                                  'The number of images on the composite
    Dim iRow As Integer                                     'The current row
    Dim iCol As Integer                                     'The current column
    Dim iRowOffset As Single                                'Amount to offset pictures on row
    Dim iWidth As Double                                    'The image width
    Dim iHeight As Double                                   'The image height
    Dim iOHeight As Integer                                 'The height of the image prior to scaling
    Dim dHSpace As Double                                   'Actual amount of white space between columns
    Dim dVSpace As Double                                   'Actual amount of white space between rows
    Dim iNumColsCurrentRow As Integer                       'Counter - Number of images on the current row
    Dim iImageNum As Integer                                'Counter - The current image number
    Dim sImageName As String                                'Temporary Image name place holder
    Dim sName As String                                     'Temporary full Image path place holder
    Dim bComplete As Boolean                                'Processing loop completion flag
    Dim sCaption1 As String                                 'Caption text 1
    Dim sCaption2 As String                                 'Caption text 2
    Dim iShiftFactor As Double                              'The shift-even-rows white space amount
    
    On Error GoTo ErrorHandler
    
    GetControls                                             'Retrieve the current form controls into locals
    
    '--- Start Adobe Indesign and setup the composite document
    Set oApp = CreateObject("InDesign.Application")         'Create an instance of the Adobe InDesign Application
    Set oDocument = oApp.Documents.Add                      'Add a new document to Indesign Publication
    Set oWindow = oApp.ActiveWindow                         'Point to the active window
    Set oSpreads = oDocument.Spreads                        'Point to the spreads in the active document
    Set oSpread = oSpreads(1)                               'Point to the first spread
    With oDocument.DocumentPreferences                      'Set document preferences
        .PageWidth = dPageWidth                             'Set the page width
        .PageHeight = dPageHeight                           'Set the page height
        .PageOrientation = idLandscape                      'Set the page orientation to landscape
    End With
    DoEvents                                                'Process Windows Events
    With oSpread.Pages(1).MarginPreferences                 'Set the page margins
        .MarginTop = dMarginTop                             'Top margin
        .MarginBottom = dMarginBottom                       'Bottom margin
        .MarginLeft = dMarginLeft                           'Left margin
        .MarginRight = dMarginRight                         'Right margin
    End With
    Set oRect = oSpread.Rectangles.Add                      'Add a border rectangle (slightly within margins)
    oRect.GeometricBounds = Array(dMarginTop + 0.02, dMarginLeft + 0.02, dPageHeight - dMarginBottom - 0.02, dPageWidth - dMarginRight - 0.02)
    
    '--- Set the image width (determined by page width minus margins minus white space divided by number of columns)
    iHeight = ((dPageHeight - (dMarginTop + dMarginBottom)) / iPageRows) * (1 - dWhiteSpace)
    iWidth = iHeight * dAspectRatio
    dHSpace = ((dPageWidth - (dMarginLeft + dMarginRight) - (iPageCols * iWidth))) / (iPageCols + 1)
    dVSpace = ((dPageHeight - (dMarginTop + dMarginBottom) - (iPageRows * iHeight))) / (iPageRows + 1)
    
    '--- Prepare pointers and counters for composite processing loop
    If optCompPageSource(0).Value = True Then               'If layout is source data
        frmMain.UsrData.rsRecords.MoveFirst                           'Move to first record in data set
    Else                                                    'Else layout is by folder
        iImageNum = 0                                       'Set the starting image to first in list
    End If
    iRow = 1                                                'Set current row pointer
    iCol = 1                                                'Set current column pointer
    bComplete = False                                       'Initialize completion status to false
    
    '--- COMPOSITE PROCESSING LOOP!
    Do While Not bComplete                                  'Loop for each record in data set
        
        '--- If even row shift is enabled then set row offset
        If chkCompRowShift.Value = 1 Then                   'If shift even rows is selected
            If iRow Mod 2 = 0 Then                          'If on an even row
                iNumColsCurrentRow = iPageCols - 1          'Reduce number of columns on the row
                iShiftFactor = (dWhiteSpace / 2) + (iWidth / 2) 'Set shift factor to 1/2 image
            Else                                            'Else processing odd numbered row
                iNumColsCurrentRow = iPageCols              'Set number of columns to normal
                iShiftFactor = 0                            'Eliminate shift factor
            End If
        Else
            iNumColsCurrentRow = iPageCols                  'Set number of columns to normal
            iShiftFactor = 0                                'Eliminate shift factor
        End If
        
        '--- Build file name from tag field or from list
        If optCompPageSource(0).Value = True Then
            sImageName = Trim$(frmMain.UsrData.rsRecords(sImageTag) & "")
            If Len(sImageName) <> 0 Then
                sName = sImagePath & "\" & sImageName & sImageExtension
            End If
        Else
            If iImageNum < frmMain.UsrImage.ImageCount Then
                frmMain.UsrImage.ImageNumber = iImageNum
            Else
                bComplete = True
                Exit Do
            End If
            sName = Trim$(frmMain.UsrImage.ImageFileName)
        End If
        
        If Dir$(sName, vbNormal) <> "" Then                 'If the image file exists!
                    
            '--- Image placement
            If (chkCompOvals.Value) Then
                Set oOval = oSpread.Ovals.Add
                oOval.GeometricBounds = Array( _
                    dMarginTop + ((iRow - 1) * iHeight) + (iRow * dVSpace), _
                    (iShiftFactor + dMarginLeft + ((iCol - 1) * (iWidth + dHSpace)) + dHSpace), _
                    dMarginTop + ((iRow - 1) * iHeight) + (iRow * dVSpace) + iHeight, _
                    (iShiftFactor + dMarginLeft + ((iCol - 1) * (iWidth + dHSpace)) + dHSpace) + iWidth)
                If iMode = 0 Then
                    oOval.Place sName, 0, 0
                    oOval.Fit idProportionally
                End If
            Else
                Set oRect = oSpread.Rectangles.Add
                oRect.GeometricBounds = Array( _
                    dMarginTop + ((iRow - 1) * iHeight) + (iRow * dVSpace), _
                    (iShiftFactor + dMarginLeft + ((iCol - 1) * (iWidth + dHSpace)) + dHSpace), _
                    dMarginTop + ((iRow - 1) * iHeight) + (iRow * dVSpace) + iHeight, _
                    (iShiftFactor + dMarginLeft + ((iCol - 1) * (iWidth + dHSpace)) + dHSpace) + iWidth)
                If iMode = 0 Then
                    oRect.Place sName, 0, 0
                    oRect.Fit idProportionally
                End If
            End If
            Set oOval = Nothing
            Set oRect = Nothing
            DoEvents
                    
            '--- Caption processing
            If chkCompCaption.Value Then                        'If captions are enabled, add them
                
                Set oTextFrame = oSpread.TextFrames.Add
                With oTextFrame
                    .GeometricBounds = Array( _
                        dMarginTop + ((iRow - 1) * iHeight) + (iRow * dVSpace) + iHeight + dCaptionOffset, _
                        (iShiftFactor + dMarginLeft + ((iCol - 1) * (iWidth + dHSpace)) + dHSpace), _
                        dMarginTop + ((iRow - 1) * iHeight) + (iRow * dVSpace) + iHeight + dVSpace, _
                        (iShiftFactor + dMarginLeft + ((iCol - 1) * (iWidth + dHSpace)) + dHSpace) + iWidth)
                        
                    If optCompPageSource(0).Value = True Then       'Process caption from fields
                         If dcbCompCaption(0).Text = "<None>" Then
                             sCaption1 = ""
                         Else
                             sCaption1 = Trim$(frmMain.UsrData.rsRecords(sCaptionField1))
                         End If
                         If dcbCompCaption(1).Text = "<None>" Then
                             sCaption2 = ""
                         Else
                             sCaption2 = Trim$(frmMain.UsrData.rsRecords(sCaptionField2))
                         End If
                         .TextContents = sCaption1 & " " & sCaption2
                    Else                                                'Process caption from file name
                        .TextContents = frmMain.UsrImage.ImageShortFileName
                    End If
                    .Paragraphs(1).Justification = idCenter
                    .Paragraphs(1).PointSize = 8#
                End With
                Set oTextFrame = Nothing
            End If
            
            '--- Increment the column number, if at specified number of composite columns then increment row.
            iCol = iCol + 1
            
            '--- Loop for each title row available
            For iTitleCount = 0 To 3
                '--- If currently processing a title row --------------------
                If Val(txtTitleRow(iTitleCount).Text) > 0 Then
                    If iRow = Val(txtTitleRow(iTitleCount).Text) Then
                        If iCol = Val(txtTitleCol(iTitleCount).Text) Then
                            iCol = Val(txtTitleCol(iTitleCount).Text) + Val(txtTitleColCount(iTitleCount).Text)
                        End If
                    End If
                End If
            Next
            
            '--- Row Processing
            If iCol = iPageCols + 1 Then
                iRow = iRow + 1
                iCol = 1
            End If
            
        End If
        
        '--- Retrieve next record/image
        If optCompPageSource(0).Value = True Then
            frmMain.UsrData.rsRecords.MoveNext
            If frmMain.UsrData.rsRecords.EOF = True Then
                bComplete = True
            End If
        Else
            iImageNum = iImageNum + 1
            bComplete = IIf(iImageNum > frmMain.UsrImage.ImageCount, True, False)
        End If
        DoEvents
    Loop
    
    Set oApp = Nothing
    Set oDocument = Nothing
    Set oSpreads = Nothing
    Set oSpread = Nothing
    Set oPageItems = Nothing
    Set oRect = Nothing
    Set oOval = Nothing
    
    MsgBox "Composite complete.", vbOKOnly + vbApplicationModal + vbInformation, EZ_CAPTION
    Exit Sub
ErrorHandler:
    Resume Next
End Sub

Public Sub AutoSize()                                       'Automatically size the composite
    Dim iImages As Integer                                  'Number of images for composite
    Dim iAns As Integer                                     'User response
    On Error GoTo ErrorHandler
    If optCompPageSource(0).Value = True Then               'Get the nubmer of images based on radio button settings
        iImages = frmMain.UsrData.GetRecordsCount()                   'Get the number of records on the main form source grid
    Else
        iImages = frmMain.UsrImage.ImageCount               'Get the number of files on the main form file list
    End If
    Select Case iImages                                     'Determine the number of columns based on the number of images
        Case 1 To 12                                        'COMP-12 standard
            iPageRows = 3
            iPageCols = 5
            iTitleRow(0) = 2
            iTitleCol(0) = 2
            iTitleColCount(0) = 3
        Case 13 To 25                                       'COMP-25 standard
            iPageRows = 4
            iPageCols = 7
            iTitleRow(0) = 2
            iTitleCol(0) = 3
            iTitleColCount(0) = 3
        Case 26 To 29                                       'COMP-29 standard
            iPageRows = 4
            iPageCols = 8
            iTitleRow(0) = 2
            iTitleCol(0) = 3
            iTitleColCount(0) = 4
        Case 30 To 42                                       'COMP-42 standard
            iPageRows = 5
            iPageCols = 9
            iTitleRow(0) = 2
            iTitleCol(0) = 4
            iTitleColCount(0) = 3
        Case Else                                           'Non-standard
            iAns = MsgBox("Too many images for standard composite type." & Chr(13) + Chr(10) & "Create non-standard?", vbYesNo + vbApplicationModal + vbQuestion, EZ_CAPTION)
            If iAns = vbNo Then
                Exit Sub
            End If
            iPageRows = 25
            iPageCols = 25
            iTitleRow(0) = 2
            iTitleCol(0) = 5
            iTitleColCount(0) = 5
    End Select
    dWhiteSpace = 0.15
    dPageWidth = 10
    dPageHeight = 7
    dMarginLeft = 0.375
    dMarginRight = 0.375
    dMarginTop = 0.375
    dMarginBottom = 0.375
    Exit Sub
ErrorHandler:
    MsgBox "Error calculating auto composite.  No source identified.", vbOKOnly + vbApplicationModal + vbInformation, EZ_CAPTION
    Resume Next
End Sub

Public Sub GetControls()
    On Error Resume Next
    
    dPageWidth = Val(txtCompPageWidth.Text)
    dPageHeight = Val(txtCompPageHeight.Text)
    dMarginTop = Val(txtCompMarginTop.Text)
    dMarginBottom = Val(txtCompMarginBottom.Text)
    dMarginLeft = Val(txtCompMarginLeft.Text)
    dMarginRight = Val(txtCompMarginRight.Text)
    dWhiteSpace = Val(txtCompWhiteSpace.Text)
    iPageCols = Val(txtCompColumns.Text)
    iPageRows = Val(txtCompRows.Text)
    
    sCaptionField1 = Trim$(dcbCompCaption(0).BoundText)
    sCaptionField2 = Trim$(dcbCompCaption(1).BoundText)
    dCaptionOffset = Val(txtCompCaptionOffset.Text)
    dCaptionFontSize = Val(txtCompCaptionFontSize.Text)
    
    dAspectRatio = frmMain.UsrImage.ImageAspectRatio
    
    For iTitleCount = 0 To 3
        iTitleRow(iTitleCount) = Val(txtTitleRow(iTitleCount).Text)
        iTitleCol(iTitleCount) = Val(txtTitleCol(iTitleCount).Text)
        iTitleColCount(iTitleCount) = Val(txtTitleColCount(iTitleCount).Text)
    Next
End Sub

Public Sub PutControls()
    On Error Resume Next
   
    txtCompPageWidth.Text = Format$(dPageWidth, txtCompPageWidth.Format)
    txtCompPageHeight.Text = Format$(dPageHeight, txtCompPageHeight.Format)
    txtCompMarginTop.Text = Format$(dMarginTop, txtCompMarginTop.Format)
    txtCompMarginBottom.Text = Format$(dMarginBottom, txtCompMarginBottom.Format)
    txtCompMarginLeft.Text = Format$(dMarginLeft, txtCompMarginLeft.Format)
    txtCompMarginRight.Text = Format$(dMarginRight, txtCompMarginRight.Format)
    txtCompWhiteSpace.Text = Format$(dWhiteSpace, txtCompWhiteSpace.Format)
    txtCompColumns.Text = Format$(iPageCols, txtCompColumns.Format)
    txtCompRows.Text = Format$(iPageRows, txtCompRows.Format)
    fraCompPageLayout.Refresh
    
    dcbCompCaption(0).BoundText = sCaptionField1
    dcbCompCaption(1).BoundText = sCaptionField2
    txtCompCaptionOffset.Text = Format$(dCaptionOffset, txtCompCaptionOffset.Format)
    txtCompCaptionFontSize.Text = Format$(dCaptionFontSize, txtCompCaptionFontSize.Format)
    
    For iTitleCount = 0 To 3
        txtTitleRow(iTitleCount).Text = Format$(iTitleRow(iTitleCount), txtTitleRow(iTitleCount).Format)
        txtTitleCol(iTitleCount).Text = Format$(iTitleCol(iTitleCount), txtTitleCol(iTitleCount).Format)
        txtTitleColCount(iTitleCount).Text = Format$(iTitleColCount(iTitleCount), txtTitleColCount(iTitleCount).Format)
    Next
End Sub

Public Sub SetDefaults()
    On Error Resume Next
    
    dPageWidth = 10#
    dPageHeight = 8#
    dMarginTop = 1#
    dMarginBottom = 1#
    dMarginLeft = 1#
    dMarginRight = 1#
    dWhiteSpace = 0.25
    iPageCols = 25#
    iPageRows = 25#
    sCaptionField1 = ""
    sCaptionField2 = ""
    dCaptionOffset = 0.02
    dCaptionFontSize = 12#
    
    For iTitleCount = 0 To 3
        iTitleRow(iTitleCount) = 0
        iTitleCol(iTitleCount) = 0
        iTitleColCount(iTitleCount) = 0
    Next
    PutControls
End Sub

Public Property Get PageWidth() As Double                              'The page width
    PageWidth = dPageWidth
End Property
Public Property Let PageWidth(dWidth As Double)
    dPageWidth = dWidth
End Property
Public Property Get PageHeight() As Double                             'The page height
    PageHeight = dPageHeight
End Property
Public Property Let PageHeight(dHeight As Double)
    dPageHeight = dHeight
End Property
Public Property Get MarginTop() As Double                              'The top margin
    MarginTop = dMarginTop
End Property
Public Property Let MarginTop(dTop As Double)
    dMarginTop = dTop
End Property
Public Property Get MarginBottom() As Double                           'The bottom margin
    MarginBottom = dMarginBottom
End Property
Public Property Let MarginBottom(dBottom As Double)
    dMarginBottom = dBottom
End Property
Public Property Get MarginLeft() As Double                             'The left margin
    MarginLeft = dMarginLeft
End Property
Public Property Let MarginLeft(dLeft As Double)
    dMarginLeft = dLeft
End Property
Public Property Get MarginRight() As Double                            'The right margin
    MarginRight = dMarginRight
End Property
Public Property Let MarginRight(dRight As Double)
    dMarginRight = dRight
End Property
Public Property Get WhiteSpace() As Double                             'The amount of white space on the page (%)
    WhiteSpace = dWhiteSpace
End Property
Public Property Let WhiteSpace(dSpace As Double)
    dWhiteSpace = dSpace
End Property
Public Property Get PageCols() As Integer                              'The number of columns on the page
    PageCols = iPageCols
End Property
Public Property Let PageCols(iCols As Integer)
    iPageCols = iCols
End Property
Public Property Get PageRows() As Integer                              'The number of rows on the page
    PageRows = iPageRows
End Property
Public Property Let PageRows(iRows As Integer)
    iPageRows = iRows
End Property
Public Property Get PageSource() As Integer                             'The page source
    PageSource = optCompPageSource(1).Value
End Property
Public Property Let PageSource(iSource As Integer)
    optCompPageSource(1).Value = iSource
End Property
Public Property Get RowShift() As Integer                               'Row shift indicator
    RowShift = chkCompRowShift.Value
End Property
Public Property Let RowShift(iShift As Integer)
    chkCompRowShift.Value = iShift
End Property
Public Property Get ImageOval() As Integer                              'Oval image indicator
    ImageOval = chkCompOvals.Value
End Property
Public Property Let ImageOval(iOval As Integer)
    chkCompOvals.Value = iOval
End Property
Public Property Get ImageCaption() As Integer                              'Oval image indicator
    ImageCaption = chkCompCaption.Value
End Property
Public Property Let ImageCaption(iCaption As Integer)
    chkCompCaption.Value = iCaption
End Property


Public Property Get CaptionField1() As String                          'Caption field
    CaptionField1 = sCaptionField1
End Property
Public Property Let CaptionField1(sField As String)
    sCaptionField1 = sField
End Property
Public Property Get CaptionField2() As String                          'Caption field
    CaptionField2 = sCaptionField2
End Property
Public Property Let CaptionField2(sField As String)
    sCaptionField2 = sField
End Property
Public Property Get CaptionOffset() As Double                          'The amount of white space to leave for the caption
    CaptionOffset = dCaptionOffset
End Property
Public Property Let CaptionOffset(dOffset As Double)
    dCaptionOffset = dOffset
End Property
Public Property Get CaptionFontSize() As Double                        'The font size to use for the caption
    CaptionFontSize = dCaptionFontSize
End Property
Public Property Let CaptionFontSize(dSize As Double)
    dCaptionFontSize = dSize
End Property
Public Property Get TitleRow1() As Integer
    TitleRow1 = iTitleRow(0)
End Property
Public Property Let TitleRow1(iValue As Integer)
   iTitleRow(0) = iValue
End Property
Public Property Get TitleRow2() As Integer
    TitleRow2 = iTitleRow(1)
End Property
Public Property Let TitleRow2(iValue As Integer)
   iTitleRow(1) = iValue
End Property
Public Property Get TitleRow3() As Integer
    TitleRow3 = iTitleRow(2)
End Property
Public Property Let TitleRow3(iValue As Integer)
   iTitleRow(2) = iValue
End Property
Public Property Get TitleRow4() As Integer
    TitleRow4 = iTitleRow(3)
End Property
Public Property Let TitleRow4(iValue As Integer)
   iTitleRow(3) = iValue
End Property
Public Property Get TitleCol1() As Integer
   TitleCol1 = iTitleCol(0)
End Property
Public Property Let TitleCol1(iValue As Integer)
   iTitleCol(0) = iValue
End Property
Public Property Get TitleCol2() As Integer
    TitleCol2 = iTitleCol(1)
End Property
Public Property Let TitleCol2(iValue As Integer)
   iTitleCol(1) = iValue
End Property
Public Property Get TitleCol3() As Integer
    TitleCol3 = iTitleCol(2)
End Property
Public Property Let TitleCol3(iValue As Integer)
   iTitleCol(2) = iValue
End Property
Public Property Get TitleCol4() As Integer
    TitleCol4 = iTitleCol(3)
End Property
Public Property Let TitleCol4(iValue As Integer)
   iTitleCol(3) = iValue
End Property
Public Property Get TitleColCount1() As Integer
   TitleColCount1 = iTitleColCount(0)
End Property
Public Property Let TitleColCount1(iValue As Integer)
   iTitleColCount(0) = iValue
End Property
Public Property Get TitleColCount2() As Integer
    TitleColCount2 = iTitleColCount(1)
End Property
Public Property Let TitleColCount2(iValue As Integer)
   iTitleColCount(1) = iValue
End Property
Public Property Get TitleColCount3() As Integer
    TitleColCount3 = iTitleColCount(2)
End Property
Public Property Let TitleColCount3(iValue As Integer)
   iTitleColCount(2) = iValue
End Property
Public Property Get TitleColCount4() As Integer
    TitleColCount4 = iTitleColCount(3)
End Property
Public Property Let TitleColCount4(iValue As Integer)
   iTitleColCount(3) = iValue
End Property
