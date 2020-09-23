VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FormNeon 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000004&
   Caption         =   "Neon Ver 1.0"
   ClientHeight    =   4275
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   5700
   Icon            =   "FormNeon.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   285
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   380
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   795
      Left            =   180
      ScaleHeight     =   49
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   223
      TabIndex        =   0
      Top             =   3150
      Width           =   3405
   End
   Begin VB.Frame Frame4 
      Caption         =   "Options: "
      Height          =   1275
      Left            =   210
      TabIndex        =   16
      Top             =   1560
      Width           =   2505
      Begin VB.TextBox TextHeightObj 
         Height          =   315
         Left            =   660
         TabIndex        =   26
         Text            =   "50"
         Top             =   810
         Width           =   405
      End
      Begin VB.TextBox TextWidthObj 
         Height          =   315
         Left            =   660
         TabIndex        =   24
         Text            =   "300"
         Top             =   300
         Width           =   405
      End
      Begin VB.CheckBox CheckLoop 
         Alignment       =   1  'Right Justify
         Caption         =   "Loop"
         Height          =   285
         Left            =   1560
         TabIndex        =   19
         Top             =   780
         Value           =   1  'Checked
         Width           =   675
      End
      Begin VB.TextBox TextStep 
         Height          =   315
         Left            =   1890
         TabIndex        =   17
         Text            =   "3"
         Top             =   300
         Width           =   405
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         Caption         =   "height:"
         Height          =   195
         Left            =   150
         TabIndex        =   27
         Top             =   870
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         Caption         =   "width:"
         Height          =   195
         Left            =   150
         TabIndex        =   25
         Top             =   360
         Width           =   420
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         Caption         =   "Step:"
         Height          =   195
         Left            =   1470
         TabIndex        =   18
         Top             =   360
         Width           =   375
      End
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   5160
      Top             =   2850
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      Filter          =   "*.BMP ; *.JPG; *.GIF|*.BMP;*.JPG;*.GIF|All Files |*.*"
      Flags           =   1
   End
   Begin VB.Frame Frame3 
      Caption         =   "Main Object: "
      Height          =   1275
      Left            =   2850
      TabIndex        =   9
      Top             =   150
      Width           =   2505
      Begin VB.CommandButton CommandMask 
         Caption         =   "Mask"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1650
         TabIndex        =   15
         Top             =   705
         Width           =   645
      End
      Begin VB.CommandButton CommandText 
         Caption         =   "Text"
         Height          =   375
         Left            =   990
         TabIndex        =   14
         Top             =   270
         Width           =   645
      End
      Begin VB.OptionButton OptionTextImage 
         Caption         =   "Image"
         Height          =   285
         Index           =   1
         Left            =   150
         TabIndex        =   13
         Top             =   750
         Width           =   765
      End
      Begin VB.OptionButton OptionTextImage 
         Caption         =   "Text"
         Height          =   285
         Index           =   0
         Left            =   150
         TabIndex        =   12
         Top             =   330
         Value           =   -1  'True
         Width           =   675
      End
      Begin VB.CommandButton CommandImage 
         Caption         =   "Image"
         Enabled         =   0   'False
         Height          =   375
         Left            =   990
         TabIndex        =   11
         Top             =   705
         Width           =   645
      End
      Begin VB.CommandButton CommandFont 
         Caption         =   "Font"
         BeginProperty Font 
            Name            =   "Century"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1650
         TabIndex        =   10
         Top             =   270
         Width           =   645
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Pixel picture: "
      Height          =   1275
      Left            =   2850
      TabIndex        =   6
      Top             =   1560
      Width           =   2505
      Begin VB.PictureBox Picture3 
         Height          =   345
         Left            =   1320
         ScaleHeight     =   19
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   61
         TabIndex        =   22
         Top             =   720
         Width           =   975
         Begin VB.PictureBox PictureOFFp 
            AutoRedraw      =   -1  'True
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            Picture         =   "FormNeon.frx":27A2
            ScaleHeight     =   17
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   58
            TabIndex        =   23
            Top             =   0
            Width           =   870
         End
      End
      Begin VB.PictureBox Picture2 
         Height          =   345
         Left            =   180
         ScaleHeight     =   19
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   61
         TabIndex        =   20
         Top             =   720
         Width           =   975
         Begin VB.PictureBox PictureONp 
            AutoRedraw      =   -1  'True
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            Picture         =   "FormNeon.frx":2808
            ScaleHeight     =   255
            ScaleWidth      =   870
            TabIndex        =   21
            Top             =   0
            Width           =   870
         End
      End
      Begin VB.CommandButton CommandPixelPic 
         Caption         =   "OFF pixels"
         Height          =   375
         Index           =   1
         Left            =   1320
         TabIndex        =   8
         Top             =   300
         Width           =   975
      End
      Begin VB.CommandButton CommandPixelPic 
         Caption         =   "ON pixels"
         Height          =   375
         Index           =   0
         Left            =   180
         TabIndex        =   7
         Top             =   300
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Move: "
      Height          =   1275
      Left            =   210
      TabIndex        =   3
      Top             =   150
      Width           =   2505
      Begin VB.CommandButton CommandPlay 
         Caption         =   "Play"
         Height          =   405
         Left            =   180
         TabIndex        =   29
         Top             =   300
         Width           =   855
      End
      Begin VB.CommandButton CommandStop 
         Caption         =   "Stop"
         Height          =   405
         Left            =   1440
         TabIndex        =   28
         Top             =   300
         Width           =   855
      End
      Begin VB.OptionButton OptionDir 
         Caption         =   "Left"
         Height          =   240
         Index           =   0
         Left            =   180
         TabIndex        =   5
         Top             =   870
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton OptionDir 
         Caption         =   "Right"
         Height          =   240
         Index           =   1
         Left            =   1590
         TabIndex        =   4
         Top             =   870
         Width           =   765
      End
   End
   Begin VB.PictureBox PictureLable 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   4500
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   73
      TabIndex        =   2
      Top             =   3300
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.TextBox TextNeon 
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   4290
      TabIndex        =   1
      Text            =   "Visual Basic Is The Best Language!         * * *     «Neon»   by:Saeed Serpooshan     (c) 2001  iran"
      Top             =   3840
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   4230
      Top             =   2880
   End
   Begin VB.PictureBox PictureTemp 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   2640
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   94
      TabIndex        =   30
      Top             =   3750
      Visible         =   0   'False
      Width           =   1410
   End
End
Attribute VB_Name = "FormNeon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------------------------------------------------------------
'      «In The Name Of The Most High »
'
' Neon Ver 1.0
' This program draw your text in a neon board
'
' by: Saeed Serpooshan - Iran - 2001 (1380)
' EMail: SSerpooshan@Yahoo.com , Admin@JamAcademic.Com
' WebPage: http://www.JamAcademic.com/vb
'------------------------------------------------------------------------------------------------------------------------------------------

Dim xBegin As Integer, ShouldRedraw As Integer, dxText As Long, dxPerViewWindow As Long
Dim pstep As Integer, qstep As Integer
Dim xbStep As Integer
Dim RightDirection As Integer


Private Sub CommandFont_Click()
On Error Resume Next
CD1.FontName = PictureLable.Font.Name
CD1.FontBold = PictureLable.Font.Bold
CD1.FontSize = PictureLable.Font.Size
CD1.FontItalic = PictureLable.Font.Italic
CD1.FontUnderline = PictureLable.Font.Underline
CD1.FontStrikethru = PictureLable.Font.Strikethrough
Err.Clear

CD1.ShowFont
If Err Then Exit Sub

PictureLable.Font.Name = CD1.FontName
PictureLable.Font.Bold = CD1.FontBold
PictureLable.Font.Size = CD1.FontSize
PictureLable.Font.Italic = CD1.FontItalic
PictureLable.Font.Underline = CD1.FontUnderline
PictureLable.Font.Strikethrough = CD1.FontStrikethru
Set TextNeon.Font = PictureLable.Font
ResetParams

End Sub

Private Sub CommandImage_Click()
MsgBox "this will add in future versions..."
End Sub

Private Sub CommandPixelPic_Click(Index As Integer)
On Error Resume Next
CD1.ShowOpen
If Err Then Exit Sub
F = CD1.FileName
If Index = 0 Then Set PictureONp.Picture = LoadPicture(F) Else Set PictureOFFp.Picture = LoadPicture(F)
ResetParams
End Sub

Private Sub CommandStop_Click()
Timer1.Enabled = False
End Sub

Private Sub CommandPlay_Click()
Timer1.Enabled = True
End Sub

Private Sub CommandText_Click()
On Error Resume Next
a = InputBox("Enter your text:", "Set Text", TextNeon.Text)
If a <> "" Then TextNeon.Text = a

End Sub

Private Sub Form_Load()
PictureONp.AutoSize = True
PictureOFFp.AutoSize = True
Call ResetParams
xb1 = xBegin: DrawNeon: xBegin = xb1
End Sub

Sub ResetParams()
RightDirection = OptionDir(1).Value = True
PictureONp.AutoSize = True
PictureOFFp.AutoSize = True

dxText = PictureLable.TextWidth(TextNeon.Text)

PictureLable.Width = dxText + 4
PictureLable.Height = PictureLable.TextHeight(TextNeon.Text) + 4
TextNeon.Height = PictureLable.Height

PictureLable.Cls
PictureLable.Print TextNeon.Text

xbStep = Val(TextStep.Text)

pstep = PictureOFFp.Width: qstep = PictureOFFp.Height

Picture1.Width = Val(TextWidthObj.Text) + 4
Picture1.Height = Val(TextHeightObj.Text) + 4

'Picture1.Width = TextNeon.Width * pstep + 4
'Picture1.Height = PictureLable.Height * qstep + 4
dxPerViewWindow = Picture1.ScaleWidth / pstep
PictureTemp.Width = Picture1.Width: PictureTemp.Height = Picture1.Height

xBegin = IIf(RightDirection, dxText, -dxPerViewWindow)
ShouldRedraw = True

End Sub

Private Sub Form_Resize()
If Picture1.Width < FormNeon.ScaleWidth Then
 Picture1.Left = (FormNeon.ScaleWidth - Picture1.Width) \ 2
Else
 Picture1.Left = 0
End If
End Sub

Private Sub OptionDir_Click(Index As Integer)
RightDirection = OptionDir(1).Value = True
If RightDirection Then
  If xBegin <= -dxPerViewWindow Then xBegin = dxText
Else 'LeftDirection:
  If xBegin >= dxText Then xBegin = -dxPerViewWindow
End If

End Sub

Private Sub OptionTextImage_Click(Index As Integer)
 Dim a As Boolean
 a = Index = 0
 CommandText.Enabled = a: CommandFont.Enabled = a
 CommandImage.Enabled = Not a: CommandMask.Enabled = Not a
End Sub

Private Sub TextHeightObj_Change()
a = Val(TextHeightObj.Text)
If a > 800 Or a < 0 Then Exit Sub
Picture1.Height = a + 4
PictureTemp.Height = Picture1.Height
ShouldRedraw = True
xb1 = xBegin: DrawNeon: xBegin = xb1
End Sub

Private Sub TextWidthObj_Change()
a = Val(TextWidthObj.Text)
If a > 1600 Or a < 0 Then Exit Sub
Picture1.Width = a + 4
PictureTemp.Width = Picture1.Width
dxPerViewWindow = Picture1.ScaleWidth / pstep
Form_Resize
ShouldRedraw = True
xb1 = xBegin: DrawNeon: xBegin = xb1
End Sub

Private Sub TextNeon_Change()
ResetParams
End Sub

Private Sub TextStep_Change()
xbStep = Val(TextStep.Text)
End Sub

Private Sub Timer1_Timer()
DrawNeon
End Sub

Sub DrawNeon()
Static LastxBegin As Long
Dim Col1 As Long, Col2 As Long, Col As Long, hdc As Long
Dim xx As Long, yy As Long
Dim x As Integer, y As Integer, dx As Long, dy As Long
Dim xxMax As Long, yyMax As Long, xAdd As Integer

Dim imgON, imgOff 'As IPictureDisp
Set imgON = PictureONp.Image
Set imgOff = PictureOFFp.Image

Col1 = PictureLable.BackColor
Col2 = PictureLable.ForeColor
xxMax = Picture1.ScaleWidth: yyMax = Picture1.ScaleHeight

dx = 2000: dy = 2000

xx = 0
xa = xBegin: xb = xBegin + dxPerViewWindow

xAdd = Abs(xBegin - LastxBegin)

If ShouldRedraw Or xAdd >= dxPerViewWindow Then
 Picture1.Cls: PictureTemp.Cls 'Cls cause to clear extra buffer and so faster draw!! (if you change height size from 50 to 250 then speed will be lower, now if you return to 50 the speed dont come back to it's previous value if you don't clear buffer by cls method!
 '(another way to clear buffer is: set Picture1.Picture = Loadpicture("")
 ShouldRedraw = False
Else
 wi = dxPerViewWindow - xAdd
 w = wi * pstep: h = yyMax
 Select Case xBegin - LastxBegin
 Case Is > 0
    If w <> 0 And h <> 0 Then
     PictureTemp.PaintPicture Picture1.Image, 0, 0, w, h, xAdd * pstep, 0, w, h
     Picture1.PaintPicture PictureTemp.Image, 0, 0, w, h, 0, 0, w, h
    End If
    xa = xa + wi: xx = xx + w
 Case 0
    xa = xBegin + dxPerViewWindow + 1 'for don't draw anything in for-next
 Case Is < 0
    If w <> 0 And h <> 0 Then
     PictureTemp.PaintPicture Picture1.Image, 0, 0, w, h, 0, 0, w, h
     Picture1.PaintPicture PictureTemp.Image, xAdd * pstep, 0, w, h, 0, 0, w, h
    End If
    xb = xb - wi
 End Select
End If

yb = 0

For x = xa To xb
yy = 0
For y = yb To yb + dy
  If x > dxText Or x < 0 Then Col = Col1 Else Col = PictureLable.Point(x, y)
  If Col = Col2 Then
   Picture1.PaintPicture imgON, xx, yy
  Else
   Picture1.PaintPicture imgOff, xx, yy
End If
  yy = yy + qstep: If yy > yyMax Then Exit For
Next
'DoEvents
xx = xx + pstep: If xx > xxMax Then Exit For
Next

LastxBegin = xBegin

If RightDirection Then
  If xBegin <= -dxPerViewWindow Then
   xBegin = dxText
  Else
   xBegin = xBegin - xbStep
  End If
Else 'LeftDirection:
  If xBegin >= dxText Then
   xBegin = -dxPerViewWindow
  Else
   xBegin = xBegin + xbStep
  End If
End If



End Sub
