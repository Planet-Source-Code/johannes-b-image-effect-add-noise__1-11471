VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Add noise (Made by Johannes.B    Email: JB_Rulez_54@hotmail.com)"
   ClientHeight    =   5790
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7005
   LinkTopic       =   "Form1"
   ScaleHeight     =   386
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   467
   StartUpPosition =   3  'Windows Default
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      LargeChange     =   2
      Left            =   600
      Max             =   10
      Min             =   1
      TabIndex        =   5
      Top             =   600
      Value           =   5
      Width           =   2655
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Save picture..."
      Height          =   375
      Left            =   4320
      TabIndex        =   4
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Open picture..."
      Height          =   375
      Left            =   2760
      TabIndex        =   3
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "CLS"
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add noise"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1815
   End
   Begin VB.PictureBox image1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   4260
      Left            =   120
      Picture         =   "Noise.frx":0000
      ScaleHeight     =   4200
      ScaleWidth      =   5610
      TabIndex        =   0
      Top             =   960
      Width           =   5670
      Begin MSComDlg.CommonDialog CM 
         Left            =   4560
         Top             =   2760
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Noise"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   600
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim BC
Dim AA, BB, AAA, BBB As Integer
Dim JB
'API calls
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long


Private Sub Command1_Click()
JB = 0
Command1.Caption = "Please wait..."
Do
'Random values
JB = JB + 1
AA = Rnd * image1.ScaleWidth
BB = Rnd * image1.ScaleHeight


AAA = Rnd * image1.Width
BBB = Rnd * image1.Height

'Get color



BC = image1.Point(AA, BB)

'Draw
SetPixel image1.hdc, AAA, BBB, BC

Loop Until JB > Val(image1.ScaleHeight) + Val(image1.ScaleWidth) * HScroll1.Value
Command1.Caption = "Add noise"
End Sub

Private Sub Command2_Click()
image1.Cls
End Sub


Private Sub Command3_Click()
CM.CancelError = True
On Error GoTo err

CM.Filter = "All supported formats ()|*.BMP;*.JPG;*.GIF;*.WMF;*.EMF;*.DIB;*.ICO;*.CUR|Bitmap (*.BMP)|*.Bmp|Bitmap (*.DIB)|*.Dib|Gif Images (*.GIF)|*.Gif|Jpeg Images (*.JPG)|*.Jpg|Metafiles (*.WMF)|*.Wmf|Metafiles (*.EMF)|*.Emf|Icons (*.ICO)|*.Ico|Icons (*.CUR)|*.Cur"

CM.ShowOpen

image1.Picture = LoadPicture(CM.FileName)

Exit Sub
err:
Exit Sub
End Sub


Private Sub Command4_Click()
CM.CancelError = True
On Error GoTo err

CM.Filter = "Bitmap (*.BMP)|*.bmp"

CM.ShowSave

SavePicture image1.Image, CM.FileName

Exit Sub
err:
Exit Sub
End Sub


