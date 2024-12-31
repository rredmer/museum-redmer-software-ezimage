Attribute VB_Name = "modMain"
'----------------------------------------------------------------------------
'
' Project......: EZ-IMAGE(r)
'
' Component....: modMain.bas
'
' Procedure....: (Declarations)
'
' Description..: Public and global variables
'
' Author.......: Ronald D. Redmer
'
' History......: 07-01-97 RDR Designed and Programmed
'
' (c) 1997-2000 Ronald D. Redmer
' All Rights Reserved
'----------------------------------------------------------------------------
Option Explicit                                             'Require explicit variable declaration
Public Const EZ_CAPTION As String = "EZ-IMAGE"
Public Const EZ_VERSION As Double = 1.01
Public Const EZ_DEBUG As Boolean = True
Public Const EZ_MSG_TECH_SUPPORT As String = "Please contact technical support."

'----------------------------------------------------------------------------
'
' Procedure....: Main
'
' Description..: Application Main Procedure: Show splash screen, open database,
'                and call the main form.
'
'----------------------------------------------------------------------------
Sub Main()                                                  'Application main procedure declaration
    On Error GoTo ErrorHandler                              'Local error handler
    frmSplash.Show                                          'Display the splash screen
    DoEvents                                                'Process windows events
    Load frmMain                                            'Load the main form
    frmSplash.Hide                                          'Hide the splash screen
    Unload frmSplash                                        'Remove splash screen from memory
    frmMain.Show                                            'Show the main application form
    Exit Sub                                                'Exit this routine
ErrorHandler:
    MsgBox "Error starting application." & vbCr & EZ_MSG_TECH_SUPPORT, vbOKOnly + vbApplicationModal + vbInformation, EZ_CAPTION
    End
End Sub
