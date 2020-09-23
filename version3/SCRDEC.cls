VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SCRDEC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'  ===================================
'  Author and ©: Bambi@CrackDealer.Com
'  ===================================

'  Do anything you like with this code but,
'  remember: do not remove my email or blame me for anything that happens
'  before, during or after running this code. ;-)

'  MS SCRENC.EXE help:
' "Script Encoder is a simple command-line tool that enables script designers
'  to encode their final script so that Web hosts and Web clients cannot view
'  or modify their source. Note that this encoding only prevents casual
'  viewing of your code; it will not prevent the determined hacker from
'  seeing what youve done and how."

'  Public Methodes:

'  DecodeScript (CodedTag As String)
'    Coded:        any string containing the encoded text including the <script> tags (or more)
'    Offset:       Number of character skipped if Coded is not a complete string
'    returns:      decoded string

'  DecodeScriptFile (FileIn As String, FileOut As String) as long
'    FileIn:       a string file representation with multiple encoded script sections
'    FileOut:      a string with decoded script sections
'    returns:      number of replaced script sections
'    FIXED:         this function may hang if you feed it a malformed encoded file

Option Explicit
Option Base 1
Private Const sVBSE As String = "VBScript.Encode"
Private Const sJSE As String = "JScript.Encode"
Private Const sVBSD As String = "VBScript"
Private Const sJSD As String = "JScript"

Private Const sSTART As String = "#@~^"
Private Const sEND As String = "^#~@"
Private Const sDELIM As String = "=="
Private Const lCHECKLEN As Long = 6&
Private Const lESCAPE As Long = 64&
Private arrlP() As Long
Private arrlC(1& To 4&, 0& To 255&) As Long

Private Sub class_initialize()
  Const sORIGINAL As String = "@ !""#$%&'()*+,-./0123456789:;=?ABCDEFGHIJKLMNOPQRSTUVWXYZ[\]^_`abcdefghijklmnopqrstuvwxyz{|}~"
  Const sP1 As String = "7~ZE:yY'vcbMQSR &Tq+f*X{01lIx_)$/93o!u(BFJ\H6K}]jPi#@pet,wD?m|C8^[nWL4r%VshUG25.dO;-Aaz""`kN=g"
  Const sP2 As String = "i,er[fuLEv*C3B cJZ8y2Wl+FRO=p{Qz~;GAwV_&xndtg}h5""U:`jqoI\]'Y%|#m(196T4bN0^s@Ka$MkD!7SXH.P-)?/"
  Const sP3 As String = "dP""Ja^][B`#e_~ORz!F &c*vG%,)i'gbAZf2sMCq9|SH1rnpI?Kj.(5}$-T7{=l4mN+0otkL3V:xWw;D/YE\h6Xy@u8UQ"

  Dim lCountA As Long, lCountB As Long
  Dim arrbO() As Byte, arrbP1() As Byte, arrbP2() As Byte, arrbP3() As Byte
  Dim arrvP As Variant

  arrvP = Array(1&, 2&, 3&, 1&, 2&, 3&, 2&, 3&, 3&, 2&, 3&, 2&, 1&, 3&, 2&, 3&, 1&, 3&, 2&, 3&, 1&, 1&, 2&, 3&, 3&, 2&, 1&, 3&, 2&, 3&, 3&, 2&, 1&, 1&, 3&, 2&, 3&, 2&, 3&, 1&, 3&, 1&, 1&, 2&, 3&, 1&, 3&, 2&, 1&, 3&, 2&, 3&, 1&, 1&, 2&, 3&, 3&, 1&, 1&, 2&, 3&, 1&, 3&, 2&)
  ReDim arrlP(LBound(arrvP) To UBound(arrvP))
  For lCountA = LBound(arrvP) To UBound(arrvP)
    arrlP(lCountA) = arrvP(lCountA)
  Next

  For lCountA = 0& To 255&
    For lCountB = 1& To 4&
      arrlC(lCountB, lCountA) = lCountA
    Next
  Next

  arrbO = StrConv(Replace(sORIGINAL, "@", vbTab, 1&, -1, vbBinaryCompare), vbFromUnicode)
  arrbP1 = StrConv(Replace(sP1, "@", vbTab, 1&, -1, vbBinaryCompare), vbFromUnicode)
  arrbP2 = StrConv(Replace(sP2, "@", vbTab, 1&, -1, vbBinaryCompare), vbFromUnicode)
  arrbP3 = StrConv(Replace(sP3, "@", vbTab, 1&, -1, vbBinaryCompare), vbFromUnicode)

  For lCountA = 0& To 93&
    arrlC(1&, arrbP1(lCountA)) = arrbO(lCountA)
    arrlC(2&, arrbP2(lCountA)) = arrbO(lCountA)
    arrlC(3&, arrbP3(lCountA)) = arrbO(lCountA)
  Next

  arrlC(4&, Asc("&")) = 10&
  arrlC(4&, Asc("#")) = 13&
  arrlC(4&, Asc("!")) = 60&
  arrlC(4&, Asc("*")) = 62&
  arrlC(4&, Asc("$")) = 64&

End Sub

Public Function DecodeScript(Coded As String, Optional ByVal Offset As Long = 0&) As String
  On Error Resume Next
  Dim arrCoded() As Byte
  Dim arrDecoded() As Byte
  Dim lPosI As Long, lPosO As Long, lPosIC As Long, lLen As Long

  lLen = Len(Coded) - 1&
  arrCoded() = StrConv(Coded, vbFromUnicode)
  ReDim arrDecoded(0& To lLen)
  Offset = Offset Mod 64

  For lPosI = 0& To lLen
    If arrCoded(lPosI) = lESCAPE Then
      lPosI = lPosI + 1&
      arrDecoded(lPosO) = arrlC(4&, arrCoded(lPosI))
    Else
      arrDecoded(lPosO) = arrlC(arrlP((lPosIC + Offset) Mod 64& + 1&), arrCoded(lPosI))
    End If
    lPosO = lPosO + 1&
    If arrCoded(lPosI) < 128& Then
      lPosIC = lPosIC + 1&
    End If
  Next

  ReDim Preserve arrDecoded(0& To lPosO - 1&)
  DecodeScript = StrConv(arrDecoded(), vbUnicode)
End Function

Public Function DecodeScriptFile(FileIn As String, FileOut As String) As Long
  On Error Resume Next
  Dim lS As Long, lE As Long, lSc As Long, lEc As Long
  Dim sScript As String, lFileIn As Long, lFileOut As Long
  Dim lFound As Long, lProtect As Long
  
  '***  lProtect is used to prevent endless loops
  '***  lProtect is the max number of loops
  lProtect = 100&
  FileOut = Replace(FileIn, sVBSE, sVBSD, 1&, , vbTextCompare)
  FileOut = Replace(FileOut, sJSE, sJSD, 1&, , vbTextCompare)

  lS = InStr(1&, FileOut, sSTART, vbBinaryCompare)
  lE = InStr(lS + 1&, FileOut, sEND, vbBinaryCompare)

  Do While lS > 0& And lE > 0& And lProtect > 0&
    lSc = lS + Len(sSTART) + lCHECKLEN + 2&
    lEc = lE - lCHECKLEN - 2&
    If Mid$(FileOut, lSc - 2&, 2&) = sDELIM Then
      If Mid$(FileOut, lEc + lCHECKLEN, 2&) = sDELIM Then
        sScript = DecodeScript(Mid$(FileOut, lSc, lEc - lSc))
        FileOut = Replace(FileOut, Mid$(FileOut, lS, lE - lS + Len(sEND)), sScript, 1&, , vbBinaryCompare)
        lFound = lFound + 1&
      End If
    End If
    lS = InStr(1&, FileOut, sSTART, vbBinaryCompare)
    lE = InStr(lS + 1&, FileOut, sEND, vbBinaryCompare)
    lProtect = lProtect - 1&
  Loop

  DecodeScriptFile = lFound
End Function
