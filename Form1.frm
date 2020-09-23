VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form Form1 
   BackColor       =   &H00404040&
   Caption         =   "Fast HTML/XML Highlight using regExp"
   ClientHeight    =   5520
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10050
   LinkTopic       =   "Form1"
   ScaleHeight     =   5520
   ScaleWidth      =   10050
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Un-Highlight"
      Height          =   375
      Left            =   3000
      TabIndex        =   3
      Top             =   3840
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Highlight All"
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Top             =   3840
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Highlight Selected"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   3840
      Width           =   1575
   End
   Begin RichTextLib.RichTextBox rtf1 
      CausesValidation=   0   'False
      Height          =   3735
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   6588
      _Version        =   393217
      ScrollBars      =   3
      TextRTF         =   $"Form1.frx":0000
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Fast HTML Highlight
'--------------------------------------------------
'Copyright 2001 DGS http://www.2dgs.com
'Written by Gary Varnell
'You may use this code freely as long as the above
'copyright info remains intact
'==================================================
' Needs reference to Microsoft VBscript Regular Expressions.
' Get it at http://msdn.microsoft.com/downloads/default.asp?URL=/downloads/sample.asp?url=/msdn-files/027/001/733/msdncompositedoc.xml
Option Explicit
Dim apppath As String
Dim starttime As Date
Dim tmpchr As String * 1
Dim tmpint As Long
Dim varColorText, varColorTag, varColorProp, varColorPropVal, varColorComment As OLE_COLOR
Private Sub Form_Load()
apppath = App.Path & IIf(Right(App.Path, 1) = "\", "", "\")
rtf1.LoadFile apppath & "testxml.html"
' Define Colors
varColorText = vbBlack
varColorTag = &HC00000
varColorProp = &HC000C0
varColorPropVal = &HC000&
varColorComment = &H808080
End Sub
Private Sub Form_Resize()
rtf1.Left = 0
rtf1.Width = Form1.ScaleWidth
rtf1.Height = Form1.ScaleHeight - 500
Command1.Top = Form1.ScaleHeight - 400
Command2.Top = Command1.Top
Command3.Top = Command1.Top
End Sub
Function colorhtml()
If rtf1.SelLength < 1 Then
    MsgBox "No text selected"
Exit Function
End If
Dim SS As Long
Dim SL As Long
Dim strBSL As String
Dim strESL As String
Dim header As String
Dim colortbl As String
Dim footer As String
' define fonts/ colors and create rtf header
Dim rtfcolor(4) As String
rtfcolor(0) = fcnGetRTFColor(varColorText)
rtfcolor(1) = fcnGetRTFColor(varColorTag)
rtfcolor(2) = fcnGetRTFColor(varColorProp)
rtfcolor(3) = fcnGetRTFColor(varColorPropVal)
rtfcolor(4) = fcnGetRTFColor(varColorComment)
colortbl = Join(rtfcolor, ";") & ";"

header = "{\rtf1\ansi\ansicpg1252\deff0\deftab720{\fonttbl{\f0\fswiss MS Sans Serif;}{\f1\froman\fcharset2 Symbol;}{\f2\fswiss MS Sans Serif;}}"
colortbl = "{\colortbl" & colortbl & "}"
header = header & vbCrLf & colortbl & vbCrLf & "\deflang1033\pard\plain\f0\fs17 "
footer = "\par \plain\f2\fs17\cf0" & vbCrLf & "\par }"
strBSL = ""
strESL = ""
If rtf1.SelLength > 0 Then
    SS = rtf1.SelStart
    SL = rtf1.SelLength
    rtf1.SelStart = 0
    rtf1.SelLength = SS
    strBSL = rtf1.SelRTF
    rtf1.SelStart = SS + SL
    rtf1.SelLength = Len(rtf1.Text) - rtf1.SelStart
    strESL = rtf1.SelRTF
    rtf1.SelStart = SS
    rtf1.SelLength = SL
End If

Dim tmpstr As String
tmpstr = rtf1.SelText
'regEx to escape RTF
tmpstr = ReplaceText("([{}\\])", "\$1", tmpstr)
tmpstr = ReplaceText("(\r)", "\par \r", tmpstr)
'regEx to color tags and prop/value pairs
tmpstr = ReplaceText("(<[^>]+>)", "\plain\f2\fs17\cf1 $1\plain\f2\fs17\cf0 ", tmpstr)
tmpstr = ReplaceText("( \w[\w\d\s:_\-\.]* *= *)(""[^""]+""|'[^']+'|\d+)", "\plain\f2\fs17\cf2 $1\plain\f2\fs17\cf3 $2\plain\f2\fs17\cf1 ", tmpstr)
' no prop just =value only
'tmpstr = ReplaceText("=(\d+|""[^""]+"")", "\plain\f2\fs17\cf2 =$1\plain\f2\fs17\cf1 ", tmpstr)
rtf1.TextRTF = header & strBSL & tmpstr & "\plain\f2\fs17\cf0 " & strESL & footer
    rtf1.SelStart = SS
    rtf1.SelLength = SL
' now fix comments and text
'regx to select all nontags with name value pairs
    Dim TagregEx, Match, Matches   ' Create variable.
    Set TagregEx = New RegExp      ' Create a regular expression.
    TagregEx.Pattern = ">[^<]*=[^>]*<"   ' Set pattern.
    TagregEx.IgnoreCase = False    ' Set case insensitivity.
    TagregEx.Global = True         ' Set global applicability.
Set Matches = TagregEx.Execute(rtf1.SelText)    ' Execute search.
For Each Match In Matches
'Debug.Print Match.Value
    rtf1.SelStart = Match.FirstIndex + SS + 1
    rtf1.SelLength = Match.Length - 2
    rtf1.SelColor = vbBlack
Next
'regx to fix comments
    Set TagregEx = New RegExp      ' Create a regular expression.
    TagregEx.Pattern = "<!--[\w\W]+?-->"   ' Set pattern.
    TagregEx.IgnoreCase = False    ' Set case insensitivity.
    TagregEx.Global = True         ' Set global applicability.
Set Matches = TagregEx.Execute(rtf1.SelText)    ' Execute search.
For Each Match In Matches
'Debug.Print Match.Value
    rtf1.SelStart = Match.FirstIndex + SS
    rtf1.SelLength = Match.Length
    rtf1.SelColor = &H808080
Next

End Function
Function ReplaceText(patrn, replStr, textStr)
  Dim regEx, str1               ' Create variables.
  Set regEx = New RegExp            ' Create regular expression.
  regEx.Pattern = patrn            ' Set pattern.
  regEx.IgnoreCase = True            ' Make case insensitive.
  regEx.Global = True
  ReplaceText = regEx.Replace(textStr, replStr)   ' Make replacement.
End Function


Private Sub Command1_Click()
rtf1.Visible = False
colorhtml
rtf1.SelLength = 0
rtf1.Visible = True
End Sub

Private Sub Command2_Click()
MousePointer = vbHourglass
starttime = Time
rtf1.Visible = False
rtf1.SelStart = 0
rtf1.SelLength = Len(rtf1.Text)
colorhtml
rtf1.SelStart = 1
rtf1.Visible = True
MousePointer = vbNormal
MsgBox "Colorize took " & Second(Time - starttime) & " Seconds" & vbCrLf & "(" & Time - starttime & ")", vbInformation, "Colorize Benchmark"
End Sub

Private Sub Command3_Click()
rtf1.TextRTF = rtf1.Text
End Sub

Private Sub rtf1_KeyPress(KeyAscii As Integer)
If Chr(KeyAscii) = "<" Then
    rtf1.SelColor = varColorTag
End If
If INcomment = True Then Exit Sub
If INtag = True Then
    If Chr(KeyAscii) = "-" Then
        ' check if we are in a comment
        rtf1.SelStart = rtf1.SelStart - 3
        rtf1.SelLength = 3
        Debug.Print rtf1.SelText
        If rtf1.SelText = "<!-" Then ' comment
          rtf1.SelColor = varColorComment
        End If
        rtf1.SelStart = rtf1.SelStart + 4
    End If
    If Chr(KeyAscii) = " " Then
        If INpropval Then
            rtf1.SelColor = varColorPropVal
        Else
            rtf1.SelColor = varColorProp
        End If
    ElseIf Chr(KeyAscii) = "=" Then
            rtf1.SelText = "="
            rtf1.SelColor = varColorPropVal
            KeyAscii = 0
    ElseIf Chr(KeyAscii) = ">" Then
            rtf1.SelColor = varColorTag
            rtf1.SelText = ">"
            KeyAscii = 0
            rtf1.SelColor = varColorText
    End If
End If
End Sub

Private Sub rtf1_KeyUp(KeyCode As Integer, Shift As Integer)
'If KeyCode & Shift = "1901" Then ' user pressed >
'    rtf1.SelColor = vbBlack
'End If
End Sub
Private Function INtag() As Boolean
If rtf1.SelStart > 0 Then
    If InStrRev(rtf1.Text, "<", rtf1.SelStart, vbTextCompare) > InStrRev(rtf1.Text, ">", rtf1.SelStart, vbTextCompare) Then INtag = True
End If
End Function
Private Function INcomment() As Boolean
If rtf1.SelStart > 0 Then
    If InStrRev(rtf1.Text, "<!--", rtf1.SelStart, vbTextCompare) > InStrRev(rtf1.Text, "-->", rtf1.SelStart, vbTextCompare) Then INcomment = True
End If
End Function
Private Function INpropval() As Boolean
Dim x, y As Long
x = InStrRev(rtf1.Text, """", rtf1.SelStart, vbTextCompare)
y = InStrRev(rtf1.Text, "=", rtf1.SelStart, vbTextCompare)
If x > y Then
If InStrRev(rtf1.Text, """", x - 1, vbTextCompare) < InStrRev(rtf1.Text, "=", x - 1, vbTextCompare) Then INpropval = True
End If
End Function


