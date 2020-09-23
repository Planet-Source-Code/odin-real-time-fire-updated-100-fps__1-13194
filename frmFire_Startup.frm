VERSION 5.00
Begin VB.Form frmFire_Startup 
   BorderStyle     =   0  'None
   Caption         =   "Check Colormode"
   ClientHeight    =   540
   ClientLeft      =   1575
   ClientTop       =   1545
   ClientWidth     =   1560
   LinkTopic       =   "Form1"
   ScaleHeight     =   540
   ScaleWidth      =   1560
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Height          =   375
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   435
      TabIndex        =   0
      Top             =   0
      Width           =   495
   End
End
Attribute VB_Name = "frmFire_Startup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'type used to determine the size of the picturebox
Private Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type
'used to get the bitmap information from picturebox
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
'used to get the color format (32, 24, 16)
Dim PicInfo As BITMAP

Private Sub Form_Load()
GetObject Picture1.Image, Len(PicInfo), PicInfo
If PicInfo.bmBitsPixel < 16 Then MsgBox "This program must run at a minimum of 16-bit color (65535)", vbInformation + vbOKOnly, "Invalid Colormode": End
If PicInfo.bmBitsPixel = 16 Then frmFire_16.Show: Me.Hide
If PicInfo.bmBitsPixel = 24 Then frmFire_New.Show: Me.Hide
If PicInfo.bmBitsPixel = 32 Then frmFire_32.Show: Me.Hide
End Sub
