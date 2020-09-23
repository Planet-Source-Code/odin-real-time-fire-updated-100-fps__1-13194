VERSION 5.00
Begin VB.Form frmFire_New 
   Caption         =   "Fire"
   ClientHeight    =   1830
   ClientLeft      =   1575
   ClientTop       =   1545
   ClientWidth     =   1560
   LinkTopic       =   "Form2"
   ScaleHeight     =   122
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
      ScaleHeight     =   25
      ScaleMode       =   0  'User
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
Attribute VB_Name = "frmFire_New"
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
'holds red colors used in flame
Dim FireRed(255) As Byte
'holds green colors used in flame
Dim FireGreen(255) As Byte
'holds blue colors used in flame
Dim FireBlue(255) As Byte
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

Public Sub DoFire_24bit()
'this sub is used if the user is in 24-bit color
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
'red is the 3rd byte so lets set it (anyone who knows C++ understands this)
PicBits(Counter + 2) = FireRed(Buffer1(NewBuf, I - fWidth)) '* 4
'green is the 2nd byte so lets set it too
PicBits(Counter + 1) = FireGreen(Buffer1(NewBuf, I - fWidth)) ' * 4.25
'blue is the 1st byte so lets set it too
PicBits(Counter + 0) = FireBlue(Buffer1(NewBuf, I - fWidth)) '* 6
'add three to the counter so we get to the next set of color
Counter = Counter + 3
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
'this function just sets the RGB colors used in
'the flame
FireRed(0) = 0
FireGreen(0) = 0
FireBlue(0) = 0
FireRed(1) = 0
FireGreen(1) = 0
FireBlue(1) = 0
FireRed(2) = 0
FireGreen(2) = 0
FireBlue(2) = 0
FireRed(3) = 0
FireGreen(3) = 0
FireBlue(3) = 0
FireRed(4) = 0
FireGreen(4) = 0
FireBlue(4) = 0
FireRed(5) = 0
FireGreen(5) = 0
FireBlue(5) = 0
FireRed(6) = 0
FireGreen(6) = 0
FireBlue(6) = 0
FireRed(7) = 0
FireGreen(7) = 0
FireBlue(7) = 0
FireRed(8) = 0
FireGreen(8) = 0
FireBlue(8) = 0
FireRed(9) = 0
FireGreen(9) = 0
FireBlue(9) = 0
FireRed(10) = 0
FireGreen(10) = 0
FireBlue(10) = 0
FireRed(11) = 0
FireGreen(11) = 0
FireBlue(11) = 0
FireRed(12) = 0
FireGreen(12) = 0
FireBlue(12) = 0
FireRed(13) = 0
FireGreen(13) = 0
FireBlue(13) = 0
FireRed(14) = 0
FireGreen(14) = 0
FireBlue(14) = 0
FireRed(15) = 0
FireGreen(15) = 0
FireBlue(15) = 0
FireRed(16) = 0
FireGreen(16) = 0
FireBlue(16) = 0
FireRed(17) = 0
FireGreen(17) = 0
FireBlue(17) = 0
FireRed(18) = 0
FireGreen(18) = 0
FireBlue(18) = 0
FireRed(19) = 0
FireGreen(19) = 0
FireBlue(19) = 0
FireRed(20) = 0
FireGreen(20) = 0
FireBlue(20) = 0
FireRed(21) = 0
FireGreen(21) = 0
FireBlue(21) = 0
FireRed(22) = 0
FireGreen(22) = 0
FireBlue(22) = 0
FireRed(23) = 0
FireGreen(23) = 0
FireBlue(23) = 0
FireRed(24) = 0
FireGreen(24) = 0
FireBlue(24) = 0
FireRed(25) = 0
FireGreen(25) = 0
FireBlue(25) = 0
FireRed(26) = 0
FireGreen(26) = 0
FireBlue(26) = 0
FireRed(27) = 0
FireGreen(27) = 0
FireBlue(27) = 0
FireRed(28) = 4
FireGreen(28) = 0
FireBlue(28) = 0
FireRed(29) = 4
FireGreen(29) = 0
FireBlue(29) = 0
FireRed(30) = 4
FireGreen(30) = 0
FireBlue(30) = 0
FireRed(31) = 4
FireGreen(31) = 0
FireBlue(31) = 0
FireRed(32) = 4
FireGreen(32) = 0
FireBlue(32) = 0
FireRed(33) = 4
FireGreen(33) = 0
FireBlue(33) = 0
FireRed(34) = 4
FireGreen(34) = 0
FireBlue(34) = 0
FireRed(35) = 8
FireGreen(35) = 0
FireBlue(35) = 0
FireRed(36) = 8
FireGreen(36) = 0
FireBlue(36) = 0
FireRed(37) = 8
FireGreen(37) = 0
FireBlue(37) = 0
FireRed(38) = 8
FireGreen(38) = 0
FireBlue(38) = 0
FireRed(39) = 8
FireGreen(39) = 0
FireBlue(39) = 0
FireRed(40) = 12
FireGreen(40) = 0
FireBlue(40) = 0
FireRed(41) = 12
FireGreen(41) = 0
FireBlue(41) = 0
FireRed(42) = 12
FireGreen(42) = 0
FireBlue(42) = 0
FireRed(43) = 12
FireGreen(43) = 0
FireBlue(43) = 0
FireRed(44) = 16
FireGreen(44) = 5
FireBlue(44) = 0
FireRed(45) = 16
FireGreen(45) = 5
FireBlue(45) = 0
FireRed(46) = 16
FireGreen(46) = 5
FireBlue(46) = 0
FireRed(47) = 16
FireGreen(47) = 5
FireBlue(47) = 0
FireRed(48) = 20
FireGreen(48) = 5
FireBlue(48) = 0
FireRed(49) = 20
FireGreen(49) = 5
FireBlue(49) = 0
FireRed(50) = 20
FireGreen(50) = 5
FireBlue(50) = 0
FireRed(51) = 24
FireGreen(51) = 5
FireBlue(51) = 0
FireRed(52) = 24
FireGreen(52) = 5
FireBlue(52) = 0
FireRed(53) = 24
FireGreen(53) = 5
FireBlue(53) = 0
FireRed(54) = 28
FireGreen(54) = 5
FireBlue(54) = 0
FireRed(55) = 28
FireGreen(55) = 10
FireBlue(55) = 0
FireRed(56) = 32
FireGreen(56) = 10
FireBlue(56) = 0
FireRed(57) = 32
FireGreen(57) = 10
FireBlue(57) = 0
FireRed(58) = 32
FireGreen(58) = 10
FireBlue(58) = 0
FireRed(59) = 36
FireGreen(59) = 10
FireBlue(59) = 0
FireRed(60) = 36
FireGreen(60) = 10
FireBlue(60) = 0
FireRed(61) = 40
FireGreen(61) = 10
FireBlue(61) = 0
FireRed(62) = 40
FireGreen(62) = 10
FireBlue(62) = 0
FireRed(63) = 44
FireGreen(63) = 10
FireBlue(63) = 0
FireRed(64) = 44
FireGreen(64) = 15
FireBlue(64) = 0
FireRed(65) = 48
FireGreen(65) = 15
FireBlue(65) = 0
FireRed(66) = 48
FireGreen(66) = 15
FireBlue(66) = 0
FireRed(67) = 52
FireGreen(67) = 15
FireBlue(67) = 0
FireRed(68) = 52
FireGreen(68) = 15
FireBlue(68) = 0
FireRed(69) = 56
FireGreen(69) = 15
FireBlue(69) = 0
FireRed(70) = 56
FireGreen(70) = 20
FireBlue(70) = 0
FireRed(71) = 60
FireGreen(71) = 20
FireBlue(71) = 0
FireRed(72) = 60
FireGreen(72) = 20
FireBlue(72) = 0
FireRed(73) = 64
FireGreen(73) = 20
FireBlue(73) = 0
FireRed(74) = 68
FireGreen(74) = 20
FireBlue(74) = 0
FireRed(75) = 68
FireGreen(75) = 20
FireBlue(75) = 0
FireRed(76) = 72
FireGreen(76) = 25
FireBlue(76) = 0
FireRed(77) = 72
FireGreen(77) = 25
FireBlue(77) = 0
FireRed(78) = 76
FireGreen(78) = 25
FireBlue(78) = 0
FireRed(79) = 80
FireGreen(79) = 25
FireBlue(79) = 0
FireRed(80) = 80
FireGreen(80) = 25
FireBlue(80) = 0
FireRed(81) = 84
FireGreen(81) = 30
FireBlue(81) = 0
FireRed(82) = 88
FireGreen(82) = 30
FireBlue(82) = 0
FireRed(83) = 88
FireGreen(83) = 30
FireBlue(83) = 0
FireRed(84) = 92
FireGreen(84) = 30
FireBlue(84) = 0
FireRed(85) = 92
FireGreen(85) = 35
FireBlue(85) = 0
FireRed(86) = 96
FireGreen(86) = 35
FireBlue(86) = 0
FireRed(87) = 100
FireGreen(87) = 35
FireBlue(87) = 0
FireRed(88) = 100
FireGreen(88) = 35
FireBlue(88) = 0
FireRed(89) = 104
FireGreen(89) = 40
FireBlue(89) = 0
FireRed(90) = 108
FireGreen(90) = 40
FireBlue(90) = 0
FireRed(91) = 108
FireGreen(91) = 40
FireBlue(91) = 0
FireRed(92) = 112
FireGreen(92) = 40
FireBlue(92) = 0
FireRed(93) = 116
FireGreen(93) = 45
FireBlue(93) = 0
FireRed(94) = 120
FireGreen(94) = 45
FireBlue(94) = 0
FireRed(95) = 120
FireGreen(95) = 45
FireBlue(95) = 0
FireRed(96) = 124
FireGreen(96) = 45
FireBlue(96) = 0
FireRed(97) = 128
FireGreen(97) = 50
FireBlue(97) = 0
FireRed(98) = 128
FireGreen(98) = 50
FireBlue(98) = 0
FireRed(99) = 132
FireGreen(99) = 50
FireBlue(99) = 0
FireRed(100) = 136
FireGreen(100) = 55
FireBlue(100) = 0
FireRed(101) = 136
FireGreen(101) = 55
FireBlue(101) = 0
FireRed(102) = 140
FireGreen(102) = 55
FireBlue(102) = 0
FireRed(103) = 144
FireGreen(103) = 60
FireBlue(103) = 0
FireRed(104) = 144
FireGreen(104) = 60
FireBlue(104) = 0
FireRed(105) = 148
FireGreen(105) = 60
FireBlue(105) = 0
FireRed(106) = 152
FireGreen(106) = 65
FireBlue(106) = 0
FireRed(107) = 152
FireGreen(107) = 65
FireBlue(107) = 0
FireRed(108) = 156
FireGreen(108) = 65
FireBlue(108) = 0
FireRed(109) = 160
FireGreen(109) = 70
FireBlue(109) = 0
FireRed(110) = 160
FireGreen(110) = 70
FireBlue(110) = 23
FireRed(111) = 164
FireGreen(111) = 70
FireBlue(111) = 23
FireRed(112) = 164
FireGreen(112) = 75
FireBlue(112) = 23
FireRed(113) = 168
FireGreen(113) = 75
FireBlue(113) = 23
FireRed(114) = 172
FireGreen(114) = 75
FireBlue(114) = 23
FireRed(115) = 172
FireGreen(115) = 80
FireBlue(115) = 23
FireRed(116) = 176
FireGreen(116) = 80
FireBlue(116) = 23
FireRed(117) = 176
FireGreen(117) = 80
FireBlue(117) = 23
FireRed(118) = 180
FireGreen(118) = 85
FireBlue(118) = 23
FireRed(119) = 184
FireGreen(119) = 85
FireBlue(119) = 23
FireRed(120) = 184
FireGreen(120) = 90
FireBlue(120) = 23
FireRed(121) = 188
FireGreen(121) = 90
FireBlue(121) = 23
FireRed(122) = 188
FireGreen(122) = 90
FireBlue(122) = 23
FireRed(123) = 192
FireGreen(123) = 95
FireBlue(123) = 23
FireRed(124) = 192
FireGreen(124) = 95
FireBlue(124) = 23
FireRed(125) = 196
FireGreen(125) = 95
FireBlue(125) = 23
FireRed(126) = 196
FireGreen(126) = 100
FireBlue(126) = 23
FireRed(127) = 200
FireGreen(127) = 100
FireBlue(127) = 23
FireRed(128) = 200
FireGreen(128) = 105
FireBlue(128) = 23
FireRed(129) = 204
FireGreen(129) = 105
FireBlue(129) = 23
FireRed(130) = 204
FireGreen(130) = 105
FireBlue(130) = 23
FireRed(131) = 208
FireGreen(131) = 110
FireBlue(131) = 23
FireRed(132) = 208
FireGreen(132) = 110
FireBlue(132) = 23
FireRed(133) = 208
FireGreen(133) = 115
FireBlue(133) = 23
FireRed(134) = 212
FireGreen(134) = 115
FireBlue(134) = 23
FireRed(135) = 212
FireGreen(135) = 115
FireBlue(135) = 23
FireRed(136) = 216
FireGreen(136) = 120
FireBlue(136) = 23
FireRed(137) = 216
FireGreen(137) = 120
FireBlue(137) = 23
FireRed(138) = 216
FireGreen(138) = 125
FireBlue(138) = 23
FireRed(139) = 220
FireGreen(139) = 125
FireBlue(139) = 46
FireRed(140) = 220
FireGreen(140) = 130
FireBlue(140) = 46
FireRed(141) = 220
FireGreen(141) = 130
FireBlue(141) = 46
FireRed(142) = 224
FireGreen(142) = 130
FireBlue(142) = 46
FireRed(143) = 224
FireGreen(143) = 135
FireBlue(143) = 46
FireRed(144) = 224
FireGreen(144) = 135
FireBlue(144) = 46
FireRed(145) = 228
FireGreen(145) = 140
FireBlue(145) = 46
FireRed(146) = 228
FireGreen(146) = 140
FireBlue(146) = 46
FireRed(147) = 228
FireGreen(147) = 145
FireBlue(147) = 46
FireRed(148) = 228
FireGreen(148) = 145
FireBlue(148) = 46
FireRed(149) = 232
FireGreen(149) = 145
FireBlue(149) = 46
FireRed(150) = 232
FireGreen(150) = 150
FireBlue(150) = 46
FireRed(151) = 232
FireGreen(151) = 150
FireBlue(151) = 46
FireRed(152) = 232
FireGreen(152) = 155
FireBlue(152) = 46
FireRed(153) = 236
FireGreen(153) = 155
FireBlue(153) = 46
FireRed(154) = 236
FireGreen(154) = 160
FireBlue(154) = 46
FireRed(155) = 236
FireGreen(155) = 160
FireBlue(155) = 46
FireRed(156) = 236
FireGreen(156) = 160
FireBlue(156) = 46
FireRed(157) = 236
FireGreen(157) = 165
FireBlue(157) = 46
FireRed(158) = 240
FireGreen(158) = 165
FireBlue(158) = 46
FireRed(159) = 240
FireGreen(159) = 170
FireBlue(159) = 69
FireRed(160) = 240
FireGreen(160) = 170
FireBlue(160) = 69
FireRed(161) = 240
FireGreen(161) = 175
FireBlue(161) = 69
FireRed(162) = 240
FireGreen(162) = 175
FireBlue(162) = 69
FireRed(163) = 240
FireGreen(163) = 175
FireBlue(163) = 69
FireRed(164) = 240
FireGreen(164) = 180
FireBlue(164) = 69
FireRed(165) = 244
FireGreen(165) = 180
FireBlue(165) = 69
FireRed(166) = 244
FireGreen(166) = 185
FireBlue(166) = 69
FireRed(167) = 244
FireGreen(167) = 185
FireBlue(167) = 69
FireRed(168) = 244
FireGreen(168) = 185
FireBlue(168) = 69
FireRed(169) = 244
FireGreen(169) = 190
FireBlue(169) = 69
FireRed(170) = 244
FireGreen(170) = 190
FireBlue(170) = 69
FireRed(171) = 244
FireGreen(171) = 195
FireBlue(171) = 69
FireRed(172) = 244
FireGreen(172) = 195
FireBlue(172) = 69
FireRed(173) = 244
FireGreen(173) = 200
FireBlue(173) = 69
FireRed(174) = 244
FireGreen(174) = 200
FireBlue(174) = 69
FireRed(175) = 248
FireGreen(175) = 200
FireBlue(175) = 69
FireRed(176) = 248
FireGreen(176) = 205
FireBlue(176) = 92
FireRed(177) = 248
FireGreen(177) = 205
FireBlue(177) = 92
FireRed(178) = 248
FireGreen(178) = 210
FireBlue(178) = 92
FireRed(179) = 248
FireGreen(179) = 210
FireBlue(179) = 92
FireRed(180) = 248
FireGreen(180) = 210
FireBlue(180) = 92
FireRed(181) = 248
FireGreen(181) = 215
FireBlue(181) = 92
FireRed(182) = 248
FireGreen(182) = 215
FireBlue(182) = 92
FireRed(183) = 248
FireGreen(183) = 215
FireBlue(183) = 92
FireRed(184) = 248
FireGreen(184) = 220
FireBlue(184) = 92
FireRed(185) = 248
FireGreen(185) = 220
FireBlue(185) = 92
FireRed(186) = 248
FireGreen(186) = 225
FireBlue(186) = 92
FireRed(187) = 248
FireGreen(187) = 225
FireBlue(187) = 92
FireRed(188) = 248
FireGreen(188) = 225
FireBlue(188) = 92
FireRed(189) = 248
FireGreen(189) = 230
FireBlue(189) = 92
FireRed(190) = 248
FireGreen(190) = 230
FireBlue(190) = 115
FireRed(191) = 248
FireGreen(191) = 230
FireBlue(191) = 115
FireRed(192) = 248
FireGreen(192) = 235
FireBlue(192) = 115
FireRed(193) = 248
FireGreen(193) = 235
FireBlue(193) = 115
FireRed(194) = 248
FireGreen(194) = 235
FireBlue(194) = 115
FireRed(195) = 248
FireGreen(195) = 240
FireBlue(195) = 115
FireRed(196) = 248
FireGreen(196) = 240
FireBlue(196) = 115
FireRed(197) = 248
FireGreen(197) = 240
FireBlue(197) = 115
FireRed(198) = 248
FireGreen(198) = 245
FireBlue(198) = 115
FireRed(199) = 248
FireGreen(199) = 245
FireBlue(199) = 115
FireRed(200) = 248
FireGreen(200) = 245
FireBlue(200) = 115
FireRed(201) = 248
FireGreen(201) = 250
FireBlue(201) = 115
FireRed(202) = 248
FireGreen(202) = 250
FireBlue(202) = 138
FireRed(203) = 248
FireGreen(203) = 250
FireBlue(203) = 138
FireRed(204) = 248
FireGreen(204) = 250
FireBlue(204) = 138
FireRed(205) = 248
FireGreen(205) = 255
FireBlue(205) = 138
FireRed(206) = 248
FireGreen(206) = 255
FireBlue(206) = 138
FireRed(207) = 248
FireGreen(207) = 255
FireBlue(207) = 138
FireRed(208) = 248
FireGreen(208) = 255
FireBlue(208) = 138
FireRed(209) = 248
FireGreen(209) = 255
FireBlue(209) = 138
FireRed(210) = 248
FireGreen(210) = 255
FireBlue(210) = 138
FireRed(211) = 248
FireGreen(211) = 255
FireBlue(211) = 138
FireRed(212) = 248
FireGreen(212) = 255
FireBlue(212) = 138
FireRed(213) = 248
FireGreen(213) = 255
FireBlue(213) = 138
FireRed(214) = 248
FireGreen(214) = 255
FireBlue(214) = 161
FireRed(215) = 248
FireGreen(215) = 255
FireBlue(215) = 161
FireRed(216) = 248
FireGreen(216) = 255
FireBlue(216) = 161
FireRed(217) = 248
FireGreen(217) = 255
FireBlue(217) = 161
FireRed(218) = 248
FireGreen(218) = 255
FireBlue(218) = 161
FireRed(219) = 248
FireGreen(219) = 255
FireBlue(219) = 161
FireRed(220) = 248
FireGreen(220) = 255
FireBlue(220) = 161
FireRed(221) = 248
FireGreen(221) = 255
FireBlue(221) = 161
FireRed(222) = 248
FireGreen(222) = 255
FireBlue(222) = 161
FireRed(223) = 248
FireGreen(223) = 255
FireBlue(223) = 161
FireRed(224) = 248
FireGreen(224) = 255
FireBlue(224) = 184
FireRed(225) = 248
FireGreen(225) = 255
FireBlue(225) = 184
FireRed(226) = 248
FireGreen(226) = 255
FireBlue(226) = 184
FireRed(227) = 248
FireGreen(227) = 255
FireBlue(227) = 184
FireRed(228) = 248
FireGreen(228) = 255
FireBlue(228) = 184
FireRed(229) = 248
FireGreen(229) = 255
FireBlue(229) = 184
FireRed(230) = 248
FireGreen(230) = 255
FireBlue(230) = 184
FireRed(231) = 248
FireGreen(231) = 255
FireBlue(231) = 184
FireRed(232) = 248
FireGreen(232) = 255
FireBlue(232) = 184
FireRed(233) = 248
FireGreen(233) = 255
FireBlue(233) = 184
FireRed(234) = 248
FireGreen(234) = 255
FireBlue(234) = 207
FireRed(235) = 248
FireGreen(235) = 255
FireBlue(235) = 207
FireRed(236) = 248
FireGreen(236) = 255
FireBlue(236) = 207
FireRed(237) = 248
FireGreen(237) = 255
FireBlue(237) = 207
FireRed(238) = 248
FireGreen(238) = 255
FireBlue(238) = 207
FireRed(239) = 248
FireGreen(239) = 255
FireBlue(239) = 207
FireRed(240) = 248
FireGreen(240) = 255
FireBlue(240) = 207
FireRed(241) = 248
FireGreen(241) = 255
FireBlue(241) = 207
FireRed(242) = 248
FireGreen(242) = 255
FireBlue(242) = 207
FireRed(243) = 248
FireGreen(243) = 255
FireBlue(243) = 230
FireRed(244) = 248
FireGreen(244) = 255
FireBlue(244) = 230
FireRed(245) = 248
FireGreen(245) = 255
FireBlue(245) = 230
FireRed(246) = 248
FireGreen(246) = 255
FireBlue(246) = 230
FireRed(247) = 248
FireGreen(247) = 255
FireBlue(247) = 230
FireRed(248) = 248
FireGreen(248) = 255
FireBlue(248) = 230
FireRed(249) = 248
FireGreen(249) = 255
FireBlue(249) = 230
FireRed(250) = 248
FireGreen(250) = 255
FireBlue(250) = 230
FireRed(251) = 248
FireGreen(251) = 255
FireBlue(251) = 253
FireRed(252) = 248
FireGreen(252) = 255
FireBlue(252) = 253
FireRed(253) = 248
FireGreen(253) = 255
FireBlue(253) = 253
FireRed(254) = 248
FireGreen(254) = 255
FireBlue(254) = 253
FireRed(255) = 248
FireGreen(255) = 255
FireBlue(255) = 253
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
If PicInfo.bmBitsPixel = 24 Then DoFire_24bit
If PicInfo.bmBitsPixel = 16 Then DoFire_24bit
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
ReDim PicBits(1 To PicInfo.bmWidth * PicInfo.bmHeight * (PicInfo.bmBitsPixel / 8)) As Byte
'get what the maximum value for our fire loop needs to be
MaxInf = (UBound(PicBits) / (PicInfo.bmBitsPixel / 8)) - fWidth - 1
'get what the minimum value for our fire loop needs to be
MinInf = fWidth + 1
'find out how many pixels there are in total
TotInf = UBound(PicBits) / (PicInfo.bmBitsPixel / 8) - 1
'setup the colors in the 3 arrays (FireRed, FireGreen, FireBlue)
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
