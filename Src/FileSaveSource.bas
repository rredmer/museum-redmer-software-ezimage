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

