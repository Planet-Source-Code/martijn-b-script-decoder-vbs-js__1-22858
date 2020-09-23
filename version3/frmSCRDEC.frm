VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmSCRDEC 
   Caption         =   "ScriptDecoder"
   ClientHeight    =   4095
   ClientLeft      =   2610
   ClientTop       =   3210
   ClientWidth     =   6990
   Icon            =   "frmSCRDEC.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4095
   ScaleWidth      =   6990
   Begin VB.CommandButton cmdEncode 
      Caption         =   "&Encode"
      Height          =   375
      Left            =   4320
      TabIndex        =   3
      Top             =   3600
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1920
      Top             =   3600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "&Load"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   3600
      Width           =   1215
   End
   Begin VB.TextBox txtDec 
      Height          =   1575
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      ToolTipText     =   "Type Decoded Text Here"
      Top             =   1800
      Width           =   6735
   End
   Begin VB.TextBox txtEnc 
      Height          =   1575
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      ToolTipText     =   "Paste Encoded Text Here"
      Top             =   120
      Width           =   6735
   End
   Begin VB.CommandButton cmdDecode 
      Caption         =   "&Decode"
      Height          =   375
      Left            =   5640
      TabIndex        =   4
      Top             =   3600
      Width           =   1215
   End
End
Attribute VB_Name = "frmSCRDEC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const APPNAME = "ScriptDecoder"
Private oSDEC As New SCRDEC
Private oScrRun As New Scripting.Encoder

Private Sub cmdDecode_Click()
  Dim sTemp As String
  On Error Resume Next
  If txtEnc = vbNullString Then
    MsgBox "Nothing to Decode..", vbInformation, APPNAME
  Else
    oSDEC.DecodeScriptFile txtEnc, sTemp
  End If
  txtDec = sTemp
End Sub

Private Sub cmdEncode_Click()
  On Error Resume Next
  If txtDec = vbNullString Then
    MsgBox "Nothing to Encode..", vbInformation, APPNAME
  Else
    txtEnc = oScrRun.EncodeScriptFile(".asp", txtDec, 0&, "VBScript")
  End If
End Sub

Private Sub cmdLoad_Click()
  Dim arrFile() As Byte
  On Error Resume Next
  CommonDialog1.ShowOpen
  If CommonDialog1.FileName <> vbNullString Then
    Open CommonDialog1.FileName For Binary As #1
    ReDim arrFile(LOF(1))
    Get #1, 1, arrFile
    Close #1
    txtEnc.Text = StrConv(arrFile, vbUnicode)
  End If
End Sub

Private Sub Form_Load()
  On Error Resume Next
  txtEnc = "<%@ LANGUAGE = VBScript.Encode %><%'Copyright© 1998. XYZ Productions. All rights reserved.'**Start Encode**#@~^hwAAAA==P4+.PbYPb/c~1KPAlHPzG!PmmU@#@&tk9n~XKEMP8ELd,hrY4~?;I3H;R2p3c@#@&?n,Y4P1Vlk/,Wk^+~0KD~r" & vbTab & "0Wc@#@&@#@&rM~^tmxT+~Y4n,l8W7nPE?^.HwYGv@#@&DSgAAA==^#~@%>"
End Sub

Private Sub Form_Resize()
  On Error Resume Next
  With txtEnc
    .Top = 0
    .Left = 0
    .Width = Me.ScaleWidth
    .Height = (Me.ScaleHeight - cmdLoad.Height) \ 2
  End With

  With txtDec
    .Top = (Me.ScaleHeight - cmdLoad.Height) \ 2
    .Left = 0
    .Width = Me.ScaleWidth
    .Height = (Me.ScaleHeight - cmdLoad.Height) \ 2
  End With

  With cmdLoad
    .Top = (Me.ScaleHeight - cmdLoad.Height)
    .Left = 0
  End With

  With cmdEncode
    .Top = (Me.ScaleHeight - cmdLoad.Height)
    .Left = cmdLoad.Left + cmdLoad.Width
    .Height = cmdLoad.Height
  End With

  With cmdDecode
    .Top = (Me.ScaleHeight - cmdLoad.Height)
    .Left = cmdEncode.Left + cmdEncode.Width
    .Height = cmdLoad.Height
  End With

End Sub
