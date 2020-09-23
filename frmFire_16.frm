VERSION 5.00
Begin VB.Form frmFire_16 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fire"
   ClientHeight    =   1815
   ClientLeft      =   1575
   ClientTop       =   1545
   ClientWidth     =   1560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   121
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   104
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   1560
      Width           =   615
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   1500
      Left            =   0
      ScaleHeight     =   100
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   100
      TabIndex        =   0
      Top             =   0
      Width           =   1500
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "0 FPS"
      Height          =   195
      Left            =   720
      TabIndex        =   2
      Top             =   1560
      Width           =   435
   End
End
Attribute VB_Name = "frmFire_16"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'this will be used to get FPS
Private Declare Function GetTickCount Lib "kernel32" () As Long
'width of the fire area
Const fWidth = 100
'height of the fire area
Const fHeight = 100
'holds the luminance of each pixel
Dim Buffer1(1, 1 To 10000) As Byte
'holds the cooling amount of each pixel
Dim CoolingMap(1 To 10000) As Byte
'a buffer to hold the previous cooling amount
Dim NCoolingmap(1 To 10000) As Byte
'holds entire color
Dim FireCol1(255) As Byte
Dim FireCol2(255) As Byte
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
'sets the pixel colors in the picturebox
Private Declare Function SetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, lpBits As Any) As Long
'used in the loops
Dim I As Long
'the maximum loop count
Dim MaxInf As Long
'the minimum loop count
Dim MinInf As Long
'how many total pixels there are
Dim TotInf As Long
'what buffer is need currently
Dim CurBuf As Byte
'what is the newer buffer
Dim NewBuf As Byte
'determines whether the fire loop is running
Dim Running As Boolean
'determines whether the fire loop should stop
Dim StopIt As Boolean
'holds the pictures pixel information
Dim PicBits() As Byte
'holds the picturebox information
Dim PicInfo As BITMAP

Public Sub DoFire_16bit()
'this sub is used if the user is in 16-bit color
'holds the starting time (for FPS)
Dim ST As Long
'holds the ending time (for FPS)
Dim ET As Long
'holds the luminance of pixel to the right
Dim N1 As Long
'holds the luminance of pixel to the left
Dim N2 As Long
'holds the luminance of pixel underneath
Dim N3 As Long
'holds the luminance of pixel above
Dim N4 As Long
'holds a value used in use with the picture
Dim Counter As Long
'holds how many frames have been done
Dim Frames As Long
'holds the value of the current buffer (see later)
Dim OldBuf As Byte
'holds the new luminance of the pixel
Dim P As Integer
'holds the cooling value of the pixel
Dim Col As Integer
'gets the current time
ST = GetTickCount
'sets the frames to 0 cuz we just started
Frames = 0
'start the loop
Do
'set the counter to 1
Counter = 1
'start loop to calculate the fire
For I = MinInf To MaxInf
'gets the luminance of the pixel to the right
N1 = Buffer1(CurBuf, I + 1)
'gets the luminance of the pixel to the left
N2 = Buffer1(CurBuf, I - 1)
'gets the luminance of the pixel underneath
N3 = Buffer1(CurBuf, I + fWidth)
'gets the luminance of the pixel above
N4 = Buffer1(CurBuf, I - fWidth)
'gets the cooling amount
Col = CoolingMap(I)
'finds the average of surrounding pixels - cooling amount
P = CByte((N1 + N2 + N3 + N4) / 4) - Col
'if value is less than 0 make it 0
If P < 0 Then P = 0
'sets the new color into the buffer
Buffer1(NewBuf, I - fWidth) = P
'color is a 16bit color now only 2 bytes (5 bits per color)
PicBits(Counter) = FireCol1(Buffer1(NewBuf, I - fWidth)) '* 4
PicBits(Counter + 1) = FireCol2(Buffer1(NewBuf, I - fWidth)) '* 4
'add two to the counter so we get to the next set of color
Counter = Counter + 2
'end of loop
Next I
'we need to swap the buffers
'this holds the current newbuf value
OldBuf = NewBuf
'sets the newbuf to the curbuf value
NewBuf = CurBuf
'sets the curbuf to the newbuf value (held in OldBuf)
CurBuf = OldBuf
'adds some hotspots
AddHotspots (100)
'adds some coldspots
AddColdSpots (100)
'draws the new image
SetBitmapBits Picture1.Image, UBound(PicBits), PicBits(1)
'updates the picturebox
Picture1.Refresh
'allows the loop to see changes in the StopIt variable
DoEvents
'adds one to frames
Frames = Frames + 1
'continue loop until StopIt doesn't equal false
Loop While StopIt = False
'gets the current time
ET = GetTickCount()
'calculates the frames per second and displays them
Label1.Caption = Format(Frames / ((ET - ST) / 1000), "0.00") & " FPS"
End Sub

Public Sub SetColorArrays()
FireCol1(0) = 0
FireCol2(0) = 0
FireCol1(1) = 0
FireCol2(1) = 0
FireCol1(2) = 0
FireCol2(2) = 0
FireCol1(3) = 0
FireCol2(3) = 0
FireCol1(4) = 0
FireCol2(4) = 0
FireCol1(5) = 0
FireCol2(5) = 0
FireCol1(6) = 0
FireCol2(6) = 0
FireCol1(7) = 0
FireCol2(7) = 0
FireCol1(8) = 0
FireCol2(8) = 0
FireCol1(9) = 0
FireCol2(9) = 0
FireCol1(10) = 0
FireCol2(10) = 0
FireCol1(11) = 0
FireCol2(11) = 0
FireCol1(12) = 0
FireCol2(12) = 0
FireCol1(13) = 0
FireCol2(13) = 0
FireCol1(14) = 0
FireCol2(14) = 0
FireCol1(15) = 0
FireCol2(15) = 0
FireCol1(16) = 0
FireCol2(16) = 0
FireCol1(17) = 0
FireCol2(17) = 0
FireCol1(18) = 0
FireCol2(18) = 0
FireCol1(19) = 0
FireCol2(19) = 0
FireCol1(20) = 0
FireCol2(20) = 0
FireCol1(21) = 0
FireCol2(21) = 0
FireCol1(22) = 0
FireCol2(22) = 0
FireCol1(23) = 0
FireCol2(23) = 0
FireCol1(24) = 0
FireCol2(24) = 0
FireCol1(25) = 0
FireCol2(25) = 0
FireCol1(26) = 0
FireCol2(26) = 0
FireCol1(27) = 0
FireCol2(27) = 0
FireCol1(28) = 0
FireCol2(28) = 0
FireCol1(29) = 0
FireCol2(29) = 0
FireCol1(30) = 0
FireCol2(30) = 0
FireCol1(31) = 0
FireCol2(31) = 0
FireCol1(32) = 0
FireCol2(32) = 0
FireCol1(33) = 0
FireCol2(33) = 0
FireCol1(34) = 0
FireCol2(34) = 0
FireCol1(35) = 0
FireCol2(35) = 0
FireCol1(36) = 0
FireCol2(36) = 0
FireCol1(37) = 0
FireCol2(37) = 0
FireCol1(38) = 0
FireCol2(38) = 0
FireCol1(39) = 0
FireCol2(39) = 0
FireCol1(40) = 0
FireCol2(40) = 8
FireCol1(41) = 0
FireCol2(41) = 8
FireCol1(42) = 0
FireCol2(42) = 8
FireCol1(43) = 0
FireCol2(43) = 8
FireCol1(44) = 0
FireCol2(44) = 8
FireCol1(45) = 0
FireCol2(45) = 8
FireCol1(46) = 0
FireCol2(46) = 8
FireCol1(47) = 0
FireCol2(47) = 8
FireCol1(48) = 0
FireCol2(48) = 16
FireCol1(49) = 0
FireCol2(49) = 16
FireCol1(50) = 0
FireCol2(50) = 16
FireCol1(51) = 0
FireCol2(51) = 16
FireCol1(52) = 0
FireCol2(52) = 16
FireCol1(53) = 0
FireCol2(53) = 16
FireCol1(54) = 0
FireCol2(54) = 24
FireCol1(55) = 64
FireCol2(55) = 24
FireCol1(56) = 64
FireCol2(56) = 24
FireCol1(57) = 64
FireCol2(57) = 24
FireCol1(58) = 64
FireCol2(58) = 24
FireCol1(59) = 64
FireCol2(59) = 32
FireCol1(60) = 64
FireCol2(60) = 32
FireCol1(61) = 64
FireCol2(61) = 32
FireCol1(62) = 64
FireCol2(62) = 32
FireCol1(63) = 64
FireCol2(63) = 40
FireCol1(64) = 64
FireCol2(64) = 40
FireCol1(65) = 64
FireCol2(65) = 40
FireCol1(66) = 64
FireCol2(66) = 40
FireCol1(67) = 64
FireCol2(67) = 48
FireCol1(68) = 64
FireCol2(68) = 48
FireCol1(69) = 64
FireCol2(69) = 48
FireCol1(70) = 128
FireCol2(70) = 48
FireCol1(71) = 128
FireCol2(71) = 56
FireCol1(72) = 128
FireCol2(72) = 56
FireCol1(73) = 128
FireCol2(73) = 56
FireCol1(74) = 128
FireCol2(74) = 64
FireCol1(75) = 128
FireCol2(75) = 64
FireCol1(76) = 192
FireCol2(76) = 64
FireCol1(77) = 192
FireCol2(77) = 64
FireCol1(78) = 192
FireCol2(78) = 72
FireCol1(79) = 192
FireCol2(79) = 72
FireCol1(80) = 192
FireCol2(80) = 72
FireCol1(81) = 192
FireCol2(81) = 80
FireCol1(82) = 192
FireCol2(82) = 80
FireCol1(83) = 192
FireCol2(83) = 80
FireCol1(84) = 192
FireCol2(84) = 88
FireCol1(85) = 0
FireCol2(85) = 89
FireCol1(86) = 0
FireCol2(86) = 89
FireCol1(87) = 0
FireCol2(87) = 97
FireCol1(88) = 0
FireCol2(88) = 97
FireCol1(89) = 0
FireCol2(89) = 97
FireCol1(90) = 0
FireCol2(90) = 105
FireCol1(91) = 0
FireCol2(91) = 105
FireCol1(92) = 0
FireCol2(92) = 105
FireCol1(93) = 64
FireCol2(93) = 113
FireCol1(94) = 64
FireCol2(94) = 113
FireCol1(95) = 64
FireCol2(95) = 113
FireCol1(96) = 64
FireCol2(96) = 121
FireCol1(97) = 128
FireCol2(97) = 121
FireCol1(98) = 128
FireCol2(98) = 121
FireCol1(99) = 128
FireCol2(99) = 129
FireCol1(100) = 128
FireCol2(100) = 129
FireCol1(101) = 128
FireCol2(101) = 129
FireCol1(102) = 128
FireCol2(102) = 137
FireCol1(103) = 192
FireCol2(103) = 137
FireCol1(104) = 192
FireCol2(104) = 137
FireCol1(105) = 192
FireCol2(105) = 137
FireCol1(106) = 192
FireCol2(106) = 145
FireCol1(107) = 192
FireCol2(107) = 145
FireCol1(108) = 192
FireCol2(108) = 145
FireCol1(109) = 0
FireCol2(109) = 154
FireCol1(110) = 0
FireCol2(110) = 154
FireCol1(111) = 0
FireCol2(111) = 154
FireCol1(112) = 64
FireCol2(112) = 154
FireCol1(113) = 64
FireCol2(113) = 162
FireCol1(114) = 64
FireCol2(114) = 162
FireCol1(115) = 64
FireCol2(115) = 162
FireCol1(116) = 64
FireCol2(116) = 170
FireCol1(117) = 64
FireCol2(117) = 170
FireCol1(118) = 128
FireCol2(118) = 170
FireCol1(119) = 128
FireCol2(119) = 178
FireCol1(120) = 128
FireCol2(120) = 178
FireCol1(121) = 128
FireCol2(121) = 178
FireCol1(122) = 128
FireCol2(122) = 178
FireCol1(123) = 192
FireCol2(123) = 186
FireCol1(124) = 192
FireCol2(124) = 186
FireCol1(125) = 192
FireCol2(125) = 186
FireCol1(126) = 0
FireCol2(126) = 187
FireCol1(127) = 0
FireCol2(127) = 195
FireCol1(128) = 0
FireCol2(128) = 195
FireCol1(129) = 0
FireCol2(129) = 195
FireCol1(130) = 0
FireCol2(130) = 195
FireCol1(131) = 64
FireCol2(131) = 203
FireCol1(132) = 64
FireCol2(132) = 203
FireCol1(133) = 64
FireCol2(133) = 203
FireCol1(134) = 64
FireCol2(134) = 203
FireCol1(135) = 64
FireCol2(135) = 203
FireCol1(136) = 128
FireCol2(136) = 211
FireCol1(137) = 128
FireCol2(137) = 211
FireCol1(138) = 192
FireCol2(138) = 211
FireCol1(139) = 194
FireCol2(139) = 211
FireCol1(140) = 194
FireCol2(140) = 211
FireCol1(141) = 194
FireCol2(141) = 211
FireCol1(142) = 194
FireCol2(142) = 219
FireCol1(143) = 2
FireCol2(143) = 220
FireCol1(144) = 2
FireCol2(144) = 220
FireCol1(145) = 66
FireCol2(145) = 220
FireCol1(146) = 66
FireCol2(146) = 220
FireCol1(147) = 66
FireCol2(147) = 220
FireCol1(148) = 66
FireCol2(148) = 220
FireCol1(149) = 66
FireCol2(149) = 228
FireCol1(150) = 130
FireCol2(150) = 228
FireCol1(151) = 130
FireCol2(151) = 228
FireCol1(152) = 130
FireCol2(152) = 228
FireCol1(153) = 130
FireCol2(153) = 228
FireCol1(154) = 194
FireCol2(154) = 228
FireCol1(155) = 194
FireCol2(155) = 228
FireCol1(156) = 194
FireCol2(156) = 228
FireCol1(157) = 2
FireCol2(157) = 229
FireCol1(158) = 2
FireCol2(158) = 237
FireCol1(159) = 4
FireCol2(159) = 237
FireCol1(160) = 4
FireCol2(160) = 237
FireCol1(161) = 68
FireCol2(161) = 237
FireCol1(162) = 68
FireCol2(162) = 237
FireCol1(163) = 68
FireCol2(163) = 237
FireCol1(164) = 68
FireCol2(164) = 237
FireCol1(165) = 68
FireCol2(165) = 237
FireCol1(166) = 132
FireCol2(166) = 237
FireCol1(167) = 132
FireCol2(167) = 237
FireCol1(168) = 132
FireCol2(168) = 237
FireCol1(169) = 196
FireCol2(169) = 237
FireCol1(170) = 196
FireCol2(170) = 237
FireCol1(171) = 196
FireCol2(171) = 237
FireCol1(172) = 196
FireCol2(172) = 237
FireCol1(173) = 4
FireCol2(173) = 238
FireCol1(174) = 4
FireCol2(174) = 238
FireCol1(175) = 4
FireCol2(175) = 246
FireCol1(176) = 4
FireCol2(176) = 246
FireCol1(177) = 4
FireCol2(177) = 246
FireCol1(178) = 68
FireCol2(178) = 246
FireCol1(179) = 68
FireCol2(179) = 246
FireCol1(180) = 68
FireCol2(180) = 246
FireCol1(181) = 132
FireCol2(181) = 246
FireCol1(182) = 132
FireCol2(182) = 246
FireCol1(183) = 132
FireCol2(183) = 246
FireCol1(184) = 132
FireCol2(184) = 246
FireCol1(185) = 132
FireCol2(185) = 246
FireCol1(186) = 196
FireCol2(186) = 246
FireCol1(187) = 196
FireCol2(187) = 246
FireCol1(188) = 196
FireCol2(188) = 246
FireCol1(189) = 196
FireCol2(189) = 246
FireCol1(190) = 198
FireCol2(190) = 246
FireCol1(191) = 198
FireCol2(191) = 246
FireCol1(192) = 6
FireCol2(192) = 247
FireCol1(193) = 6
FireCol2(193) = 247
FireCol1(194) = 6
FireCol2(194) = 247
FireCol1(195) = 70
FireCol2(195) = 247
FireCol1(196) = 70
FireCol2(196) = 247
FireCol1(197) = 70
FireCol2(197) = 247
FireCol1(198) = 70
FireCol2(198) = 247
FireCol1(199) = 70
FireCol2(199) = 247
FireCol1(200) = 70
FireCol2(200) = 247
FireCol1(201) = 134
FireCol2(201) = 247
FireCol1(202) = 136
FireCol2(202) = 247
FireCol1(203) = 136
FireCol2(203) = 247
FireCol1(204) = 136
FireCol2(204) = 247
FireCol1(205) = 200
FireCol2(205) = 247
FireCol1(206) = 200
FireCol2(206) = 247
FireCol1(207) = 200
FireCol2(207) = 247
FireCol1(208) = 200
FireCol2(208) = 247
FireCol1(209) = 200
FireCol2(209) = 247
FireCol1(210) = 200
FireCol2(210) = 247
FireCol1(211) = 200
FireCol2(211) = 247
FireCol1(212) = 200
FireCol2(212) = 247
FireCol1(213) = 200
FireCol2(213) = 247
FireCol1(214) = 202
FireCol2(214) = 247
FireCol1(215) = 202
FireCol2(215) = 247
FireCol1(216) = 202
FireCol2(216) = 247
FireCol1(217) = 202
FireCol2(217) = 247
FireCol1(218) = 202
FireCol2(218) = 247
FireCol1(219) = 202
FireCol2(219) = 247
FireCol1(220) = 202
FireCol2(220) = 247
FireCol1(221) = 202
FireCol2(221) = 247
FireCol1(222) = 202
FireCol2(222) = 247
FireCol1(223) = 202
FireCol2(223) = 247
FireCol1(224) = 202
FireCol2(224) = 247
FireCol1(225) = 202
FireCol2(225) = 247
FireCol1(226) = 202
FireCol2(226) = 247
FireCol1(227) = 202
FireCol2(227) = 247
FireCol1(228) = 202
FireCol2(228) = 247
FireCol1(229) = 202
FireCol2(229) = 247
FireCol1(230) = 202
FireCol2(230) = 247
FireCol1(231) = 202
FireCol2(231) = 247
FireCol1(232) = 202
FireCol2(232) = 247
FireCol1(233) = 202
FireCol2(233) = 247
FireCol1(234) = 204
FireCol2(234) = 247
FireCol1(235) = 204
FireCol2(235) = 247
FireCol1(236) = 204
FireCol2(236) = 247
FireCol1(237) = 204
FireCol2(237) = 247
FireCol1(238) = 204
FireCol2(238) = 247
FireCol1(239) = 204
FireCol2(239) = 247
FireCol1(240) = 204
FireCol2(240) = 247
FireCol1(241) = 204
FireCol2(241) = 247
FireCol1(242) = 204
FireCol2(242) = 247
FireCol1(243) = 206
FireCol2(243) = 247
FireCol1(244) = 206
FireCol2(244) = 247
FireCol1(245) = 206
FireCol2(245) = 247
FireCol1(246) = 206
FireCol2(246) = 247
FireCol1(247) = 206
FireCol2(247) = 247
FireCol1(248) = 206
FireCol2(248) = 247
FireCol1(249) = 206
FireCol2(249) = 247
FireCol1(250) = 206
FireCol2(250) = 247
FireCol1(251) = 208
FireCol2(251) = 247
FireCol1(252) = 208
FireCol2(252) = 247
FireCol1(253) = 208
FireCol2(253) = 247
FireCol1(254) = 208
FireCol2(254) = 247
FireCol1(255) = 208
FireCol2(255) = 247
End Sub

Public Sub AddColdSpots(ByVal Number As Long)
'adds cooling spots so the flame cools unevenly
'variable for the for loop
Dim I As Long
'sets up the randomize function
Randomize Timer
'start the loop
For I = 1 To Number
'creates a cool pixel placed randomly with a random cooling amount
CoolingMap(Int(Rnd * TotInf) + 1) = Int(Rnd * 10)
'end of loop
Next I
'holds the cooling pixel to the right of current
Dim N1 As Long
'holds the cooling pixel to the left of current
Dim N2 As Long
'holds the cooling pixel down from the current
Dim N3 As Long
'holds the cooling pixel up from the current
Dim N4 As Long
'starts the loop (don't need edges)
For I = MinInf To MaxInf
'gets the pixels to the right value
N1 = CoolingMap(I + 1)
'gets the pixels to the left value
N2 = CoolingMap(I - 1)
'gets the pixels underneath value
N3 = CoolingMap(I + fWidth)
'gets the pixels above value
N4 = CoolingMap(I - fWidth)
'gets the average of the pixels around it
NCoolingmap(I) = CByte((N1 + N2 + N3 + N4) / 4)
'end of loop
Next I
For I = 1 To TotInf - fWidth
'copy the pixels back but up one pixel
CoolingMap(I) = NCoolingmap(I + fWidth)
'end of loop
Next I
End Sub

Public Sub AddHotspots(ByVal Number As Long)
'add hot spots so the flame grows from the bottom
'for the loop
Dim I As Long
'setup the randomize function
Randomize Timer
'start of the for loop
For I = 1 To Number
'adds a hotspot to the bottom with a random value
Buffer1(CurBuf, TotInf - Int(Rnd * fWidth) - fWidth) = Int(Rnd * 8) + 247 '255 'Int(Rnd * 191) + 64
'end of loop
Next I
End Sub

Private Sub Command1_Click()
'checks to see if the loop is already running
If Running = True Then
'if running, then stop it
StopIt = True
'if not running then lets start
Else
'let everything know it is running
Running = True
'we don't want to stopit, we just started it
StopIt = False
'change the command so user knows to click to stop
Command1.Caption = "Stop"
Call DoFire_16bit
'loop is stopped so we don't need to stop it anymore
StopIt = False
'loop isn't running anymore
Running = False
'let user know to click command to start fire up
Command1.Caption = "Start"
'end the if statement from above (beginning of sub)
End If
End Sub

Private Sub Form_Activate()
'the loop isn't running
Running = False
'since the loop isn't running we don't need to stop it
StopIt = False
'the current buffer used is the first one
CurBuf = 0
'the buffer to hold the new values is the second one
NewBuf = 1
'we need to get the bitmap information from picture
GetObject Picture1.Image, Len(PicInfo), PicInfo
'setup the buffer to hold the colors
ReDim PicBits(1 To PicInfo.bmWidth * PicInfo.bmHeight * 2) As Byte
'get what the maximum value for our fire loop needs to be
MaxInf = (UBound(PicBits) / 2) - fWidth - 1
'get what the minimum value for our fire loop needs to be
MinInf = fWidth + 1
'find out how many pixels there are in total
TotInf = UBound(PicBits) / 2 - 1
'setup the colors in the 2 arrays (FireCol1, FireCol2)
SetColorArrays
'add some hotspots to start
AddHotspots (50)
'add some coldspots to start
AddColdSpots (250)
End Sub

Private Sub Form_Unload(Cancel As Integer)
'if the user closes the program, make sure loop is stopped
Running = False
'we need to stop the loop
StopIt = True
'end the program
End
End Sub
