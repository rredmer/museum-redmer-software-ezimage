VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSplash 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3105
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   7140
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   7140
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList imgSplash 
      Left            =   6510
      Top             =   2490
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   334
      ImageHeight     =   219
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSplash.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSplash.frx":35B36
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lblCompanyProduct 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "EZ-IMAGE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   555
      Left            =   3045
      TabIndex        =   7
      Top             =   165
      Width           =   2310
   End
   Begin VB.Label lblPlatform 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Windows 98/NT/2000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3015
      TabIndex        =   6
      Top             =   1350
      Width           =   3480
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      X1              =   3000
      X2              =   7080
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "(c) 1992-2000 Redmer Software Company"
      Height          =   255
      Left            =   3030
      TabIndex        =   5
      Top             =   900
      Width           =   4005
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "All Rights Reserved."
      Height          =   195
      Left            =   3030
      TabIndex        =   4
      Top             =   1140
      Width           =   4005
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FF0000&
      X1              =   90
      X2              =   7080
      Y1              =   2220
      Y2              =   2220
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Warning:  This program is protected by copyright law and international treaties.  Unauthorized"
      Height          =   285
      Left            =   90
      TabIndex        =   3
      Top             =   2280
      Width           =   6915
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "reproduction or distribution of this program, or any portion of it, may result in severe civil"
      Height          =   315
      Left            =   90
      TabIndex        =   2
      Top             =   2550
      Width           =   6855
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "and criminal penalties, and will be prosecuted to the maximum extent possible under law."
      Height          =   315
      Left            =   90
      TabIndex        =   1
      Top             =   2820
      Width           =   6855
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "â"
      BeginProperty Font 
         Name            =   "Symbol"
         Size            =   20.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   405
      Left            =   5460
      TabIndex        =   0
      Top             =   30
      Width           =   345
   End
   Begin VB.Image imgLogo 
      Height          =   1965
      Left            =   60
      Stretch         =   -1  'True
      Top             =   60
      Width           =   2895
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    On Error Resume Next
    imgLogo.Picture = imgSplash.ListImages(1).Picture
End Sub
