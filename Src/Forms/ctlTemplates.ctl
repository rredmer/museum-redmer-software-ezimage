VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl ctlTemplates 
   ClientHeight    =   8565
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6765
   ScaleHeight     =   8565
   ScaleWidth      =   6765
   Begin VB.DriveListBox drvTemplate 
      Height          =   315
      Left            =   1140
      TabIndex        =   3
      Top             =   90
      Width           =   5505
   End
   Begin VB.DirListBox dirTemplate 
      Height          =   1890
      Left            =   1140
      TabIndex        =   2
      Top             =   420
      Width           =   5505
   End
   Begin VB.FileListBox filTemplate 
      Height          =   5550
      Left            =   1140
      Pattern         =   "*.INDD"
      TabIndex        =   1
      Top             =   2310
      Width           =   5535
   End
   Begin MSComctlLib.ImageList imgTemplate 
      Left            =   1590
      Top             =   7890
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   23
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlTemplates.ctx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlTemplates.ctx":067A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbTemplate 
      Height          =   645
      Left            =   90
      TabIndex        =   0
      Top             =   7890
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   1138
      ButtonWidth     =   1191
      ButtonHeight    =   1138
      Appearance      =   1
      Style           =   1
      ImageList       =   "imgTemplate"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Preview"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Build"
            ImageIndex      =   2
         EndProperty
      EndProperty
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Template"
      Height          =   225
      Left            =   60
      TabIndex        =   5
      Top             =   2370
      Width           =   765
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Location"
      Height          =   225
      Left            =   60
      TabIndex        =   4
      Top             =   120
      Width           =   765
   End
End
Attribute VB_Name = "ctlTemplates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'----------------------------------------------------------------------------
'
' Project......: RSC EZ-IMAGE(r)
'
' Component....: clsTemplate
'
' Procedure....: (Declarations)
'
' Description..: Process Adobe InDesign Templates
'
' Author.......: Ronald D. Redmer
'
' History......: 06-01-00 RDR Designed and Programmed
'
' (c) 1997-2000 Redmer Software Company, Inc.
' All Rights Reserved
'----------------------------------------------------------------------------
Option Explicit

'--- Dimension template properties
Private sTemplateName As String                                         'File name of template
Private sImagePath As String                                            'Path to image files
Private sImageTag As String                                             'Field used to tag images
Private sImageExtension As String                                       'Extension of images
Private bComplete As Boolean

Private Sub tlbTemplate_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error Resume Next
    Select Case Button.Index                                            'Select on the index of the button clicked
        Case 1
            sTemplateName = Trim$(dirTemplate.Path) & "\" & Trim$(filTemplate.FileName) 'Name of template front-side file
            Preview
        Case 2                                                          'Process template button
            sTemplateName = Trim$(dirTemplate.Path) & "\" & Trim$(filTemplate.FileName) 'Name of template front-side file
            sImagePath = Trim$(frmMain.UsrImage.ImagePath)              'Path to source images
            sImageTag = Trim$(frmMain.UsrData.ImageTag) & ""            'Data field ised to tag images
            sImageExtension = frmMain.UsrImage.ImageExtension
            Build
    End Select
End Sub

Public Sub Build()
    
    Dim oApp As InDesign.Application                                    'Adobe InDesign Application Object
    Dim oOriginal As InDesign.Document                                  'Adobe InDesign Document Object - Original
    Dim oDocument As InDesign.Document                                  'Adobe InDesign Document Object - Populated
    Dim oTextFrame As InDesign.TextFrame                                'Adobe InDesign Text Frame Object
    Dim oWindow As InDesign.Window                                      'Adobe InDesign Window Object
    Dim oSpreads As InDesign.Spreads
    Dim oSpread As InDesign.Spread                                      'Adobe InDesign Spread Object
    Dim oPageItems As InDesign.PageItems                                'Adobe InDesign Page Items Collection Object
    Dim oImage As InDesign.Image                                        'Adobe InDesign Image Object
    Dim iRecordCount As Integer                                         'The number of records in the recordset
    Dim iRec As Integer                                                 'Record counter
    Dim iSpread As Integer                                              'Spread counter
    Dim iPage As Integer                                                'Page counter (Counts number of spreads per Record)
    Dim iObject As Integer                                              'Object counter
    Dim sType As String                                                 'Object Name
    Dim oCurrent As Object                                              'The current object being templated
    
    On Error GoTo ErrorHandler
    
    Set oApp = CreateObject("InDesign.Application")                     'Open Adobe InDesign
    Set oOriginal = oApp.Open(sTemplateName)                            'Open the template file
    Set oWindow = oApp.ActiveWindow                                     'Point to the active window
    Set oSpreads = oOriginal.Spreads                                    'Point to the spreads in the active document
    DoEvents                                                            'Process windows events
    
    '--- Create a spreads for each record in the dataset
    iRecordCount = frmMain.UsrData.GetRecordsCount                      'Retrieve record count in active dataset
    For iRec = 1 To iRecordCount                                        'Loop for each record in the dataset
        Set oSpread = oOriginal.Spreads.Item(1)                         'Point to the active spread in the active window
        oSpread.Duplicate                                               'Duplicate first page of spread
        Set oSpread = oOriginal.Spreads.Item(2)                         'Point to the active spread in the active window
        oSpread.Duplicate                                               'Duplicate second page of spread
        DoEvents                                                        'Process windows events
    Next iRec
    
    '--- Loop through each record in the dataset and populate each spread
    iSpread = 1                                                         'Set pointer to first spread
    frmMain.UsrData.rsRecords.MoveFirst                                 'Move record pointer to first in set
    bComplete = False                                                   'Set completion flag
    Do While Not bComplete                                              'Loop until complete
        For iPage = 1 To 2                                              'Loop for each page in spread (currently on supports 2 page spreads)
            Set oSpread = oOriginal.Spreads.Item(iSpread)               'Point to the proper spread
            Set oPageItems = oSpread.PageItems                          'Return a collection of pageitems
            For iObject = 1 To oPageItems.Count                         'Loop for every object on the page
                Set oCurrent = oPageItems.Item(iObject)                 'Get a pointer to each object
                sType = TypeName(oCurrent)                              'Get the type of the object
                If sType = "TextFrame" Then                             'If the object type is a textframe, then perform field replacement
                    With frmMain.UsrData.rsColumns                      'Using the column information from the data set
                        .MoveFirst                                      'Point to the first column
                        Do While Not .EOF                               'Loop for each column in the dataset
                            If Trim$(UCase(oCurrent.TextContents)) = Trim$(UCase(.Fields("COLUMN").Value)) Then     'If the text contents match the column name
                                oCurrent.TextContents = frmMain.UsrData.rsRecords(.Fields("COLUMN").Value).Value    'Replace the text contents with the column data in the current record
                            End If
                            .MoveNext                                   'Move to next column record
                        Loop
                    End With
                End If
                If sType = "Rectangle" Then                             'If the object type is a rectangle, then load images based on layer name
                    If oCurrent.ItemLayer.Name = "Photograph" Then      'If the rectangle is in the Photograph Layer, place the appropriate image into it
                        oCurrent.Place sImagePath & Trim$(frmMain.UsrData.rsRecords(sImageTag).Value) & Trim$(sImageExtension), 0, 0
                        oCurrent.Fit idProportionally                   'Set the fitting to proportional (i.e. keep aspect ratio)
                    End If
                End If
            Next iObject                                                'Loop for next object on page
            iSpread = iSpread + 1                                       'Increment the spread number
        Next iPage                                                      'Loop for each page in the document
        frmMain.UsrData.rsRecords.MoveNext                              'Move to the next record in the dataset
        If frmMain.UsrData.rsRecords.EOF = True Then bComplete = True   'If at end of file break loop
    Loop
    
    '--- Release object pointers for exit
    Set oApp = Nothing
    Set oDocument = Nothing
    Set oWindow = Nothing
    Set oSpread = Nothing
    Set oPageItems = Nothing
    Set oCurrent = Nothing
    MsgBox "Template processing completed.", vbApplicationModal + vbOKOnly + vbInformation, EZ_CAPTION
    Exit Sub
ErrorHandler:
    Resume Next
End Sub

Public Sub Preview()
    Dim oPrv As Object
    On Error Resume Next
    Set oPrv = CreateObject("InDesign.Application")
    oPrv.Open sTemplateName                                 'Open the template file
    Set oPrv = Nothing
End Sub

Private Sub dirTemplate_Change()                            'Update file list on directory change
    On Error Resume Next
    filTemplate.Path = dirTemplate.Path
End Sub

Private Sub drvTemplate_Change()                            'Update Directory list on drive change
    On Error Resume Next
    dirTemplate.Path = drvTemplate.Drive
End Sub

Public Property Let TemplatePath(sPath As String)           'Set the Drive and Directory paths on property change
    On Error Resume Next
    Dim fso As Scripting.FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    drvTemplate.Drive = fso.GetDriveName(sPath)
    dirTemplate.Path = sPath
    filTemplate.Path = sPath
    Set fso = Nothing
End Property

Public Property Get TemplatePath() As String                'Retrieve template path (with trailing back-slash)
    On Error Resume Next
    TemplatePath = IIf(Right$(Trim$(dirTemplate.Path), 1) <> "\", Trim$(dirTemplate.Path) & "\", Trim$(dirTemplate.Path))
End Property

Public Property Let TemplateName(sName As String)
    On Error Resume Next
    filTemplate.FileName = sName
End Property

Public Property Get TemplateName() As String
    On Error Resume Next
    TemplateName = filTemplate.FileName
End Property
