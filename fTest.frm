VERSION 5.00
Begin VB.Form frmVBJPEG 
   Caption         =   "Save Picture To JPEG"
   ClientHeight    =   7845
   ClientLeft      =   420
   ClientTop       =   345
   ClientWidth     =   9060
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "fTest.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7845
   ScaleWidth      =   9060
   Begin VB.PictureBox picTest 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   0
      Picture         =   "fTest.frx":1272
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   1
      Top             =   2880
      Width           =   480
   End
   Begin VB.CommandButton cmdSaveStdPic 
      Caption         =   "&Save..."
      Height          =   315
      Left            =   240
      TabIndex        =   0
      Top             =   2160
      Width           =   1395
   End
End
Attribute VB_Name = "frmVBJPEG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim tJ As JPEG_CORE_PROPERTIES_VB
Dim filename As String

Private Sub cmdSaveStdPic_Click()
camheight = 288
camwidth = 352
bmpjpgzetten
For i = 1 To 20
BitBlt m_hDC, 0, 0, camwidth, camheight, picTest.hdc, 0, 0, vbSrcCopy
a = ijlInit(tJ)
tJ.DIBWidth = camwidth
tJ.DIBHeight = -camheight
tJ.DIBBytes = m_lPtr
tJ.DIBPadBytes = 0
tJ.JPGWidth = camwidth
tJ.JPGHeight = camheight
tJ.jquality = 51
filename = App.Path & "\Tempv" & i & ".jpg"
CopyMemory tJ.JPGFile, filename, 4
a = ijlWrite(tJ, 8)
Next i
bmpjpgopnul
End Sub

