VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Fire"
   ClientHeight    =   1830
   ClientLeft      =   1545
   ClientTop       =   1545
   ClientWidth     =   1560
   LinkTopic       =   "Form1"
   ScaleHeight     =   1830
   ScaleWidth      =   1560
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      Height          =   255
      Left            =   20
      TabIndex        =   1
      Top             =   1560
      Width           =   615
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   1500
      Left            =   20
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
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'used to get the bitmap information from picturebox
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
'sets the pixel colors in the picturebox
Private Declare Function SetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, lpBits As Any) As Long
'this will be used to get FPS
Private Declare Function GetTickCount Lib "kernel32" () As Long
'this will be used to speed up buffer updates
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
'width of the fire area
Const fWidth = 100
'height of the fire area
Const fHeight = 100
'holds the luminance of each pixel
Dim Buffer1(1, 1 To 10000) As Byte
'holds the cooling amount of each pixel
Dim CoolingMap(1 To 10000) As Byte
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
'holds the number to step
Dim nStep As Integer
'Holds the current cooling location
Dim curCooling As Long

Public Function GetCooling(ByVal numI As Long) As Byte
    'Holds the location of the current pixel in coolingmap
    Dim Loc As Long
    'Check if the current location has looped
    If curCooling > TotInf Then curCooling = curCooling - TotInf
    'Check if the current location has looped
    If curCooling < 0 Then curCooling = curCooling + TotInf
    'Caclculate the cooling position
    Loc = numI + curCooling
    'Make sure the location hasn't looped
    If Loc > (TotInf + 1) Then Loc = Loc - (TotInf + 1)
    'Return the correct amount to cool
    GetCooling = CoolingMap(Loc)
End Function

Public Sub DoFire(ByVal iStep As Integer)
'this sub calculates and draws each frame
    
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
        DoEvents
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
            Col = GetCooling(I)
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
            Counter = Counter + iStep
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
        AddHotspots (15)
        'draws the new image
        SetBitmapBits Picture1.Image, UBound(PicBits), PicBits(1)
        'updates the picturebox
        Picture1.Refresh
        'allows the loop to see changes in the StopIt variable
        DoEvents
        'adds one to frames
        Frames = Frames + 1
        
        'Causing the flame to "flow"
        'curCooling = curCooling + Int(Rnd * fWidth * 6) - (fWidth * 3) + Int(Rnd * 10) - 5
        'Causes the flame to rise "normally"
        curCooling = curCooling + fWidth
        
    'continue loop until StopIt doesn't equal false
    Loop While StopIt = False
    'gets the current time
    ET = GetTickCount()
    'calculates the frames per second and displays them
    Label1.Caption = Format(Frames / ((ET - ST) / 1000), "0.00") & " FPS"
End Sub

Public Sub Set_Color_16()
'This function generates the color array need to
'display "accurate" fire in 16 bit color.
    FireBlue(0) = 0
    FireGreen(0) = 0
    FireBlue(1) = 0
    FireGreen(1) = 0
    FireBlue(2) = 0
    FireGreen(2) = 0
    FireBlue(3) = 0
    FireGreen(3) = 0
    FireBlue(4) = 0
    FireGreen(4) = 0
    FireBlue(5) = 0
    FireGreen(5) = 0
    FireBlue(6) = 0
    FireGreen(6) = 0
    FireBlue(7) = 0
    FireGreen(7) = 0
    FireBlue(8) = 0
    FireGreen(8) = 0
    FireBlue(9) = 0
    FireGreen(9) = 0
    FireBlue(10) = 0
    FireGreen(10) = 0
    FireBlue(11) = 0
    FireGreen(11) = 0
    FireBlue(12) = 0
    FireGreen(12) = 0
    FireBlue(13) = 0
    FireGreen(13) = 0
    FireBlue(14) = 0
    FireGreen(14) = 0
    FireBlue(15) = 0
    FireGreen(15) = 0
    FireBlue(16) = 0
    FireGreen(16) = 0
    FireBlue(17) = 0
    FireGreen(17) = 0
    FireBlue(18) = 0
    FireGreen(18) = 0
    FireBlue(19) = 0
    FireGreen(19) = 0
    FireBlue(20) = 0
    FireGreen(20) = 0
    FireBlue(21) = 0
    FireGreen(21) = 0
    FireBlue(22) = 0
    FireGreen(22) = 0
    FireBlue(23) = 0
    FireGreen(23) = 0
    FireBlue(24) = 0
    FireGreen(24) = 0
    FireBlue(25) = 0
    FireGreen(25) = 0
    FireBlue(26) = 0
    FireGreen(26) = 0
    FireBlue(27) = 0
    FireGreen(27) = 0
    FireBlue(28) = 0
    FireGreen(28) = 0
    FireBlue(29) = 0
    FireGreen(29) = 0
    FireBlue(30) = 0
    FireGreen(30) = 0
    FireBlue(31) = 0
    FireGreen(31) = 0
    FireBlue(32) = 0
    FireGreen(32) = 0
    FireBlue(33) = 0
    FireGreen(33) = 0
    FireBlue(34) = 0
    FireGreen(34) = 0
    FireBlue(35) = 0
    FireGreen(35) = 0
    FireBlue(36) = 0
    FireGreen(36) = 0
    FireBlue(37) = 0
    FireGreen(37) = 0
    FireBlue(38) = 0
    FireGreen(38) = 0
    FireBlue(39) = 0
    FireGreen(39) = 0
    FireBlue(40) = 0
    FireGreen(40) = 8
    FireBlue(41) = 0
    FireGreen(41) = 8
    FireBlue(42) = 0
    FireGreen(42) = 8
    FireBlue(43) = 0
    FireGreen(43) = 8
    FireBlue(44) = 0
    FireGreen(44) = 8
    FireBlue(45) = 0
    FireGreen(45) = 8
    FireBlue(46) = 0
    FireGreen(46) = 8
    FireBlue(47) = 0
    FireGreen(47) = 8
    FireBlue(48) = 0
    FireGreen(48) = 16
    FireBlue(49) = 0
    FireGreen(49) = 16
    FireBlue(50) = 0
    FireGreen(50) = 16
    FireBlue(51) = 0
    FireGreen(51) = 16
    FireBlue(52) = 0
    FireGreen(52) = 16
    FireBlue(53) = 0
    FireGreen(53) = 16
    FireBlue(54) = 0
    FireGreen(54) = 24
    FireBlue(55) = 64
    FireGreen(55) = 24
    FireBlue(56) = 64
    FireGreen(56) = 24
    FireBlue(57) = 64
    FireGreen(57) = 24
    FireBlue(58) = 64
    FireGreen(58) = 24
    FireBlue(59) = 64
    FireGreen(59) = 32
    FireBlue(60) = 64
    FireGreen(60) = 32
    FireBlue(61) = 64
    FireGreen(61) = 32
    FireBlue(62) = 64
    FireGreen(62) = 32
    FireBlue(63) = 64
    FireGreen(63) = 40
    FireBlue(64) = 64
    FireGreen(64) = 40
    FireBlue(65) = 64
    FireGreen(65) = 40
    FireBlue(66) = 64
    FireGreen(66) = 40
    FireBlue(67) = 64
    FireGreen(67) = 48
    FireBlue(68) = 64
    FireGreen(68) = 48
    FireBlue(69) = 64
    FireGreen(69) = 48
    FireBlue(70) = 128
    FireGreen(70) = 48
    FireBlue(71) = 128
    FireGreen(71) = 56
    FireBlue(72) = 128
    FireGreen(72) = 56
    FireBlue(73) = 128
    FireGreen(73) = 56
    FireBlue(74) = 128
    FireGreen(74) = 64
    FireBlue(75) = 128
    FireGreen(75) = 64
    FireBlue(76) = 192
    FireGreen(76) = 64
    FireBlue(77) = 192
    FireGreen(77) = 64
    FireBlue(78) = 192
    FireGreen(78) = 72
    FireBlue(79) = 192
    FireGreen(79) = 72
    FireBlue(80) = 192
    FireGreen(80) = 72
    FireBlue(81) = 192
    FireGreen(81) = 80
    FireBlue(82) = 192
    FireGreen(82) = 80
    FireBlue(83) = 192
    FireGreen(83) = 80
    FireBlue(84) = 192
    FireGreen(84) = 88
    FireBlue(85) = 0
    FireGreen(85) = 89
    FireBlue(86) = 0
    FireGreen(86) = 89
    FireBlue(87) = 0
    FireGreen(87) = 97
    FireBlue(88) = 0
    FireGreen(88) = 97
    FireBlue(89) = 0
    FireGreen(89) = 97
    FireBlue(90) = 0
    FireGreen(90) = 105
    FireBlue(91) = 0
    FireGreen(91) = 105
    FireBlue(92) = 0
    FireGreen(92) = 105
    FireBlue(93) = 64
    FireGreen(93) = 113
    FireBlue(94) = 64
    FireGreen(94) = 113
    FireBlue(95) = 64
    FireGreen(95) = 113
    FireBlue(96) = 64
    FireGreen(96) = 121
    FireBlue(97) = 128
    FireGreen(97) = 121
    FireBlue(98) = 128
    FireGreen(98) = 121
    FireBlue(99) = 128
    FireGreen(99) = 129
    FireBlue(100) = 128
    FireGreen(100) = 129
    FireBlue(101) = 128
    FireGreen(101) = 129
    FireBlue(102) = 128
    FireGreen(102) = 137
    FireBlue(103) = 192
    FireGreen(103) = 137
    FireBlue(104) = 192
    FireGreen(104) = 137
    FireBlue(105) = 192
    FireGreen(105) = 137
    FireBlue(106) = 192
    FireGreen(106) = 145
    FireBlue(107) = 192
    FireGreen(107) = 145
    FireBlue(108) = 192
    FireGreen(108) = 145
    FireBlue(109) = 0
    FireGreen(109) = 154
    FireBlue(110) = 0
    FireGreen(110) = 154
    FireBlue(111) = 0
    FireGreen(111) = 154
    FireBlue(112) = 64
    FireGreen(112) = 154
    FireBlue(113) = 64
    FireGreen(113) = 162
    FireBlue(114) = 64
    FireGreen(114) = 162
    FireBlue(115) = 64
    FireGreen(115) = 162
    FireBlue(116) = 64
    FireGreen(116) = 170
    FireBlue(117) = 64
    FireGreen(117) = 170
    FireBlue(118) = 128
    FireGreen(118) = 170
    FireBlue(119) = 128
    FireGreen(119) = 178
    FireBlue(120) = 128
    FireGreen(120) = 178
    FireBlue(121) = 128
    FireGreen(121) = 178
    FireBlue(122) = 128
    FireGreen(122) = 178
    FireBlue(123) = 192
    FireGreen(123) = 186
    FireBlue(124) = 192
    FireGreen(124) = 186
    FireBlue(125) = 192
    FireGreen(125) = 186
    FireBlue(126) = 0
    FireGreen(126) = 187
    FireBlue(127) = 0
    FireGreen(127) = 195
    FireBlue(128) = 0
    FireGreen(128) = 195
    FireBlue(129) = 0
    FireGreen(129) = 195
    FireBlue(130) = 0
    FireGreen(130) = 195
    FireBlue(131) = 64
    FireGreen(131) = 203
    FireBlue(132) = 64
    FireGreen(132) = 203
    FireBlue(133) = 64
    FireGreen(133) = 203
    FireBlue(134) = 64
    FireGreen(134) = 203
    FireBlue(135) = 64
    FireGreen(135) = 203
    FireBlue(136) = 128
    FireGreen(136) = 211
    FireBlue(137) = 128
    FireGreen(137) = 211
    FireBlue(138) = 192
    FireGreen(138) = 211
    FireBlue(139) = 194
    FireGreen(139) = 211
    FireBlue(140) = 194
    FireGreen(140) = 211
    FireBlue(141) = 194
    FireGreen(141) = 211
    FireBlue(142) = 194
    FireGreen(142) = 219
    FireBlue(143) = 2
    FireGreen(143) = 220
    FireBlue(144) = 2
    FireGreen(144) = 220
    FireBlue(145) = 66
    FireGreen(145) = 220
    FireBlue(146) = 66
    FireGreen(146) = 220
    FireBlue(147) = 66
    FireGreen(147) = 220
    FireBlue(148) = 66
    FireGreen(148) = 220
    FireBlue(149) = 66
    FireGreen(149) = 228
    FireBlue(150) = 130
    FireGreen(150) = 228
    FireBlue(151) = 130
    FireGreen(151) = 228
    FireBlue(152) = 130
    FireGreen(152) = 228
    FireBlue(153) = 130
    FireGreen(153) = 228
    FireBlue(154) = 194
    FireGreen(154) = 228
    FireBlue(155) = 194
    FireGreen(155) = 228
    FireBlue(156) = 194
    FireGreen(156) = 228
    FireBlue(157) = 2
    FireGreen(157) = 229
    FireBlue(158) = 2
    FireGreen(158) = 237
    FireBlue(159) = 4
    FireGreen(159) = 237
    FireBlue(160) = 4
    FireGreen(160) = 237
    FireBlue(161) = 68
    FireGreen(161) = 237
    FireBlue(162) = 68
    FireGreen(162) = 237
    FireBlue(163) = 68
    FireGreen(163) = 237
    FireBlue(164) = 68
    FireGreen(164) = 237
    FireBlue(165) = 68
    FireGreen(165) = 237
    FireBlue(166) = 132
    FireGreen(166) = 237
    FireBlue(167) = 132
    FireGreen(167) = 237
    FireBlue(168) = 132
    FireGreen(168) = 237
    FireBlue(169) = 196
    FireGreen(169) = 237
    FireBlue(170) = 196
    FireGreen(170) = 237
    FireBlue(171) = 196
    FireGreen(171) = 237
    FireBlue(172) = 196
    FireGreen(172) = 237
    FireBlue(173) = 4
    FireGreen(173) = 238
    FireBlue(174) = 4
    FireGreen(174) = 238
    FireBlue(175) = 4
    FireGreen(175) = 246
    FireBlue(176) = 4
    FireGreen(176) = 246
    FireBlue(177) = 4
    FireGreen(177) = 246
    FireBlue(178) = 68
    FireGreen(178) = 246
    FireBlue(179) = 68
    FireGreen(179) = 246
    FireBlue(180) = 68
    FireGreen(180) = 246
    FireBlue(181) = 132
    FireGreen(181) = 246
    FireBlue(182) = 132
    FireGreen(182) = 246
    FireBlue(183) = 132
    FireGreen(183) = 246
    FireBlue(184) = 132
    FireGreen(184) = 246
    FireBlue(185) = 132
    FireGreen(185) = 246
    FireBlue(186) = 196
    FireGreen(186) = 246
    FireBlue(187) = 196
    FireGreen(187) = 246
    FireBlue(188) = 196
    FireGreen(188) = 246
    FireBlue(189) = 196
    FireGreen(189) = 246
    FireBlue(190) = 198
    FireGreen(190) = 246
    FireBlue(191) = 198
    FireGreen(191) = 246
    FireBlue(192) = 6
    FireGreen(192) = 247
    FireBlue(193) = 6
    FireGreen(193) = 247
    FireBlue(194) = 6
    FireGreen(194) = 247
    FireBlue(195) = 70
    FireGreen(195) = 247
    FireBlue(196) = 70
    FireGreen(196) = 247
    FireBlue(197) = 70
    FireGreen(197) = 247
    FireBlue(198) = 70
    FireGreen(198) = 247
    FireBlue(199) = 70
    FireGreen(199) = 247
    FireBlue(200) = 70
    FireGreen(200) = 247
    FireBlue(201) = 134
    FireGreen(201) = 247
    FireBlue(202) = 136
    FireGreen(202) = 247
    FireBlue(203) = 136
    FireGreen(203) = 247
    FireBlue(204) = 136
    FireGreen(204) = 247
    FireBlue(205) = 200
    FireGreen(205) = 247
    FireBlue(206) = 200
    FireGreen(206) = 247
    FireBlue(207) = 200
    FireGreen(207) = 247
    FireBlue(208) = 200
    FireGreen(208) = 247
    FireBlue(209) = 200
    FireGreen(209) = 247
    FireBlue(210) = 200
    FireGreen(210) = 247
    FireBlue(211) = 200
    FireGreen(211) = 247
    FireBlue(212) = 200
    FireGreen(212) = 247
    FireBlue(213) = 200
    FireGreen(213) = 247
    FireBlue(214) = 202
    FireGreen(214) = 247
    FireBlue(215) = 202
    FireGreen(215) = 247
    FireBlue(216) = 202
    FireGreen(216) = 247
    FireBlue(217) = 202
    FireGreen(217) = 247
    FireBlue(218) = 202
    FireGreen(218) = 247
    FireBlue(219) = 202
    FireGreen(219) = 247
    FireBlue(220) = 202
    FireGreen(220) = 247
    FireBlue(221) = 202
    FireGreen(221) = 247
    FireBlue(222) = 202
    FireGreen(222) = 247
    FireBlue(223) = 202
    FireGreen(223) = 247
    FireBlue(224) = 202
    FireGreen(224) = 247
    FireBlue(225) = 202
    FireGreen(225) = 247
    FireBlue(226) = 202
    FireGreen(226) = 247
    FireBlue(227) = 202
    FireGreen(227) = 247
    FireBlue(228) = 202
    FireGreen(228) = 247
    FireBlue(229) = 202
    FireGreen(229) = 247
    FireBlue(230) = 202
    FireGreen(230) = 247
    FireBlue(231) = 202
    FireGreen(231) = 247
    FireBlue(232) = 202
    FireGreen(232) = 247
    FireBlue(233) = 202
    FireGreen(233) = 247
    FireBlue(234) = 204
    FireGreen(234) = 247
    FireBlue(235) = 204
    FireGreen(235) = 247
    FireBlue(236) = 204
    FireGreen(236) = 247
    FireBlue(237) = 204
    FireGreen(237) = 247
    FireBlue(238) = 204
    FireGreen(238) = 247
    FireBlue(239) = 204
    FireGreen(239) = 247
    FireBlue(240) = 204
    FireGreen(240) = 247
    FireBlue(241) = 204
    FireGreen(241) = 247
    FireBlue(242) = 204
    FireGreen(242) = 247
    FireBlue(243) = 206
    FireGreen(243) = 247
    FireBlue(244) = 206
    FireGreen(244) = 247
    FireBlue(245) = 206
    FireGreen(245) = 247
    FireBlue(246) = 206
    FireGreen(246) = 247
    FireBlue(247) = 206
    FireGreen(247) = 247
    FireBlue(248) = 206
    FireGreen(248) = 247
    FireBlue(249) = 206
    FireGreen(249) = 247
    FireBlue(250) = 206
    FireGreen(250) = 247
    FireBlue(251) = 208
    FireGreen(251) = 247
    FireBlue(252) = 208
    FireGreen(252) = 247
    FireBlue(253) = 208
    FireGreen(253) = 247
    FireBlue(254) = 208
    FireGreen(254) = 247
    FireBlue(255) = 208
    FireGreen(255) = 247
End Sub

Public Sub Set_Color_2432()
'This function generates the color array need to
'display "accurate" fire in 24 + 32 bit color.
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

Public Sub GenerateCoolMap()
'creates a cooling map so the flame cools unevenly
    'Variable for the for loop
    Dim I As Long
    'Sets up the randomize function
    Randomize Timer
    'Holds the randomly generated number
    Dim intRand As Integer
    'Holds the randomly generated cooling amount
    Dim intCool As Integer
    'a buffer to hold the previous cooling amount
    Dim NCoolingMap(1 To 10000) As Byte

    For I = 1 To 250
        'Generate the location to add coldspot
        intRand = Int(Rnd * (TotInf - fWidth - fWidth)) + fWidth + 1
        'Generate the cooling amount
        intCool = Int(Rnd * 25) + 10
        'Set the cooling amount
        CoolingMap(intRand) = intCool
        CoolingMap(intRand + 1) = intCool
        CoolingMap(intRand - 1) = intCool
        'CoolingMap(intRand + fWidth) = intCool
        'CoolingMap(intRand - fWidth) = intCool
    Next I
    
    'Holds the second for loop
    Dim S As Long
    'Smooth the cooling map 25 times
    For S = 0 To 25
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
            'NCoolingmap(I) = CByte((N1 + N2 + N3 + N4) / 4)
            CopyMemory NCoolingMap(I), CByte((N1 + N2 + N3 + N4) / 4), 1
        Next I
        'Copy the data into the coolingmap
        CopyMemory CoolingMap(1), NCoolingMap(1), TotInf
    Next S
End Sub

Public Sub AddHotspots(ByVal Number As Long)
'add hot spots so the flame grows from the bottom
'for the loop
    Dim I As Long
    'setup the randomize function
    Randomize Timer

    For I = 1 To Number
        'adds a hotspot to the bottom with a random value
        Buffer1(CurBuf, TotInf - Int(Rnd * fWidth * 2) - fWidth) = Int(Rnd * 25) + 220
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
        
        'Start the rendering of fire
        Call DoFire(nStep)
        
        'loop is stopped so we don't need to stop it anymore
        StopIt = False
        'loop isn't running anymore
        Running = False
        'let user know to click command to start fire up
        Command1.Caption = "Start"
        'end the if statement from above (beginning of sub)
    End If
End Sub

Private Sub Form_Load()
    'Get information about the picturebox
    GetObject Picture1.Image, Len(PicInfo), PicInfo
    'Check if the color depth used on this machine
    If PicInfo.bmBitsPixel < 16 Then
        'Aren't using the minimum 16-bit color
        MsgBox "This program must run at a minimum of 16-bit color (65535)", vbInformation + vbOKOnly, "Invalid Colormode"
        'Stop the programs execution
        End
    ElseIf PicInfo.bmBitsPixel = 16 Then
        '16-bit color is being used
        Set_Color_16
        '2 bytes per pixel
        nStep = 2
    ElseIf PicInfo.bmBitsPixel = 24 Then
        '24-bit color is being used
        Set_Color_2432
        '3 bytes per pixel
        nStep = 3
    ElseIf PicInfo.bmBitsPixel = 32 Then
        '32-bit color is being used
        Set_Color_2432
        '4 bytes per pixel
        nStep = 4
    Else
        'Higher than 32-bit color?
        MsgBox "You are using a color mode higher than 32-bit, please lower to run program", vbInformation + vbOKOnly, "Invalid Colormode"
        'End the program
        End
    End If
    'The current cooling location is 0
    curCooling = 0
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
    'add some hotspots to start
    AddHotspots (5)
    'Create the coolingmap
    GenerateCoolMap
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'if the user closes the program, make sure loop is stopped
    Running = False
    'we need to stop the loop
    StopIt = True
    'end the program
    End
End Sub
