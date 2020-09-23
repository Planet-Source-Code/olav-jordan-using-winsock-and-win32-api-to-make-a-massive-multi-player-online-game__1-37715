VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmBotz 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "WAR BOTZ"
   ClientHeight    =   6225
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   7410
   FontTransparent =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   415
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   494
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock wskConnector 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Frame fraVendor 
      Height          =   4575
      Left            =   480
      TabIndex        =   0
      Top             =   600
      Visible         =   0   'False
      Width           =   2895
      Begin VB.VScrollBar vsrHealth 
         Height          =   375
         Left            =   960
         Max             =   0
         Min             =   10
         TabIndex        =   27
         Top             =   960
         Width           =   375
      End
      Begin VB.CheckBox chkRocket 
         Caption         =   "Rocket  300"
         Height          =   255
         Left            =   1560
         TabIndex        =   22
         Top             =   3480
         Width           =   1215
      End
      Begin VB.CheckBox chkArmour 
         Caption         =   "Armour  1000"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   3480
         Width           =   1335
      End
      Begin VB.CheckBox chkGun 
         Caption         =   "Sub Machine Gun 800"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   20
         Top             =   3120
         Width           =   1935
      End
      Begin VB.CheckBox chkGun 
         Caption         =   "Shot Gun 400"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   19
         Top             =   2760
         Width           =   1335
      End
      Begin VB.VScrollBar vsrSheild 
         Enabled         =   0   'False
         Height          =   375
         Left            =   960
         Max             =   0
         Min             =   10
         TabIndex        =   9
         Top             =   2400
         Width           =   375
      End
      Begin VB.CommandButton cmdBuy 
         Caption         =   "&Buy"
         Default         =   -1  'True
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   4200
         Width           =   2535
      End
      Begin VB.VScrollBar vsrBullets 
         Enabled         =   0   'False
         Height          =   375
         Index           =   1
         Left            =   960
         Max             =   0
         Min             =   1000
         TabIndex        =   5
         Top             =   1920
         Width           =   375
      End
      Begin VB.VScrollBar vsrBullets 
         Enabled         =   0   'False
         Height          =   375
         Index           =   0
         Left            =   960
         Max             =   0
         Min             =   1000
         TabIndex        =   4
         Top             =   1440
         Width           =   375
      End
      Begin VB.Label Label11 
         Caption         =   "Health"
         Height          =   255
         Left            =   1560
         TabIndex        =   28
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label lblHealth 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   375
         Left            =   480
         TabIndex        =   26
         Top             =   960
         Width           =   375
      End
      Begin VB.Label Label10 
         Caption         =   "5"
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   960
         Width           =   135
      End
      Begin VB.Label Label9 
         Caption         =   "Shields"
         Height          =   255
         Left            =   1560
         TabIndex        =   18
         Top             =   2400
         Width           =   1095
      End
      Begin VB.Label Label8 
         Caption         =   "Sub M. amo."
         Height          =   255
         Left            =   1560
         TabIndex        =   17
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "Shot Gun amo"
         Height          =   255
         Left            =   1560
         TabIndex        =   16
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "Cost"
         Height          =   255
         Left            =   2160
         TabIndex        =   15
         Top             =   3840
         Width           =   495
      End
      Begin VB.Label Label5 
         Caption         =   "Have"
         Height          =   255
         Left            =   2160
         TabIndex        =   14
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label4 
         Caption         =   "40"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   2400
         Width           =   255
      End
      Begin VB.Label Label3 
         Caption         =   "20"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   1920
         Width           =   255
      End
      Begin VB.Label Label2 
         Caption         =   "10"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1440
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "Cost"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Width           =   735
      End
      Begin VB.Label lblSheild 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   375
         Left            =   480
         TabIndex        =   8
         Top             =   2400
         Width           =   375
      End
      Begin VB.Label lblCost 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   3840
         Width           =   1815
      End
      Begin VB.Label lblBullets 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   375
         Index           =   1
         Left            =   480
         TabIndex        =   3
         Top             =   1920
         Width           =   375
      End
      Begin VB.Label lblBullets 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   375
         Index           =   0
         Left            =   480
         TabIndex        =   2
         Top             =   1440
         Width           =   375
      End
      Begin VB.Label lblHave 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.PictureBox picInfo 
      FontTransparent =   0   'False
      Height          =   5775
      Left            =   4200
      ScaleHeight     =   381
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   181
      TabIndex        =   24
      Top             =   0
      Width           =   2775
   End
   Begin VB.PictureBox picArena 
      Height          =   5775
      Left            =   0
      ScaleHeight     =   381
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   269
      TabIndex        =   23
      Top             =   0
      Width           =   4095
   End
   Begin VB.Timer tmrBullet 
      Interval        =   1
      Left            =   120
      Top             =   600
   End
   Begin MSWinsockLib.Winsock wskBot 
      Left            =   120
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuConnect 
         Caption         =   "&Connect"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuDisconnect 
         Caption         =   "&Disconnect"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuBuy 
         Caption         =   "&Buy Weapons"
      End
      Begin VB.Menu mnuSeperator 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
         Shortcut        =   {F3}
      End
   End
End
Attribute VB_Name = "frmBotz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'program:         WarBotz
'programmer:      Olav Jordan
'Purpose:         entertaining online game
'
'if anyone has any suggestions for enhancements to this game please post on psc
'i will run a server for this game from 2:00pm (eastern time us/canada) to 2:00am (eastern time us/canada) from august 6 to the 12 2002
'TO PLAY ON MY SERVER U MUST CONNECT WITH COMPILED EXE THAT COMES WITH GAME
'I WILL NOT GIVE OUT MY IP TO ANYONE
'the exe connects to my ip that is the only difference between source and exe it is NOT a virus

'RECT
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function CopyRect Lib "user32" (lpDestRect As RECT, lpSourceRect As RECT) As Long
Private Declare Function IsRectEmpty Lib "user32" (lpRect As RECT) As Long
Private Declare Function IntersectRect Lib "user32" (lpDestRect As RECT, lpSrc1Rect As RECT, lpSrc2Rect As RECT) As Long

'DC/BITMAPS
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

'SCREEN RESOLUTION
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function CreateDC Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As String, lpInitData As Any) As Long
Private Declare Function ChangeDisplaySettings Lib "user32" Alias "ChangeDisplaySettingsA" (lpDevMode As Any, ByVal dwFlags As Long) As Long
Private Declare Function EnumDisplaySettings Lib "user32" Alias "EnumDisplaySettingsA" (ByVal lpszDeviceName As Long, ByVal iModeNum As Long, lpDevMode As Any) As Boolean

'TEXT
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long

'SOUND,TIME,KEYBOARDEVENTS,MESSAGE QUEUE
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Private Declare Function GetQueueStatus Lib "user32" (ByVal fuFlags As Long) As Long

Private Enum E_EVENTS 'MESSAGE QUEUE EVENTS
   QS_TIMER = &H10
   QS_POSTMESSAGE = &H8
   QS_PAINT = &H20
   QS_MOUSEMOVE = &H2
   QS_MOUSEBUTTON = &H4
   QS_KEY = &H1
   QS_HOTKEY = &H80
   QS_MOUSE = (QS_MOUSEMOVE Or QS_MOUSEBUTTON)
   QS_INPUT = (QS_MOUSE Or QS_KEY)
   QS_ALLEVENTS = (QS_INPUT Or QS_POSTMESSAGE Or QS_TIMER Or QS_PAINT Or QS_HOTKEY)
End Enum

Private Enum E_TEXT 'TEXT DISPLAY
   TRANSPARENT = 1

   DT_CALCRECT = &H400
   DT_SINGLELINE = &H20
   DT_CENTER = &H1
End Enum

Private Enum E_RESOLUTION 'RESOLUTION INFORMATION
   CCHDEVICENAME = 32
   CCHFORMNAME = 32
   DM_BITSPERPEL = &H40000
   DM_PELSWIDTH = &H80000
   DM_PELSHEIGHT = &H100000
   CDS_TEST = &H4
   BITSPIXEL = 12         '  Number of bits per pixel
End Enum

Private Enum E_SOUNDS   'SOUND INFORMATION
   SND_ASYNC = &H1         '  play asynchronously
   SND_NOWAIT = &H2000      '  don't wait if the driver is busy
   SND_LOOP = &H8         '  loop the sound until next sndPlaySound
   SND_NODEFAULT = &H2         '  silence not default, if sound not found
   SND_NOSTOP = &H10        '  don't stop any currently playing sound
End Enum

Private Enum E_GRAPHICS    'GRAPHICS INFORMATION
   IMAGE_BITMAP = 0
   
   LR_LOADFROMFILE = &H10
   LR_CREATEDIBSECTION = &H2000
End Enum

Private Enum E_PLAYER      'PLAYER INFORMATION
   LIFE_HEIGHT = 16

   MAX_HEALTH = 50
   
   PLAYER_WIDTH = 24
   PLAYER_HEIGHT = 32
End Enum

Private Type BITMAP '14 bytes    'BITMAP INFORMATION
        bmType        As Long
        bmWidth       As Long
        bmHeight      As Long
        bmWidthBytes  As Long
        bmPlanes      As Integer
        bmBitsPixel   As Integer
        bmBits        As Long
End Type

Private Type RECT          'RECTANGLES
   Left        As Long
   Top         As Long
   Right       As Long
   Bottom      As Long
End Type

Private Type typInfo       'BASIC INFORMATION OF PLAYER
   Lives             As Long
   Health            As Long
   Money             As Long
   Bullet(2)         As Long
   Connected         As Boolean
   ID                As Long
   Armour            As Long
   Sheilds           As Long
   ShieldOn          As Long
   SheildsHoldTime   As Long
   Named             As String
End Type

Private Type typSlide         'INFORMATION ON SLIDE OF ANIMATION
   Frame             As Long
   Direction         As Long
End Type

Private Type typMovement      'MOVEMENT/TIMEING OF PLAYER/ITEMS
   MoveRate             As Long
   TBeforeMove          As Long
   TBeforeTurn          As Long
   TBeforeChangeGun     As Long
   TLastMove            As Long
   TLastTurn            As Long
   TLastChangeGun       As Long
End Type

Private Type typPlayer     'PLAYER TOTAL INFORMATION
   Pos         As RECT
   Info        As typInfo
   Pic         As typSlide
   Move        As typMovement
End Type

Private Type typIMAGE_INFO    'INFORMATION ON BITMAPS
   DC       As Long
   Info     As BITMAP
End Type

Private Type typBuffer        'BUFFERS
   Back        As Long
   Clean       As Long
   hBlt        As Long
   Info        As BITMAP
End Type

Private Type typBullet        'INFORMATION ON BULLETS
   Pos            As RECT
   Damage         As Long
   MoveRate       As Long
   Fired          As Boolean
   Direction      As Long
   TLastMove      As Long
   ID             As Long
End Type

Private Type typEnemyBullet      'INFORMATION ON ENEMIES
   Info() As typBullet
End Type

Private Type DEVMODE             'SCREEN RESOLUTION
        dmDeviceName          As String * CCHDEVICENAME
        dmSpecVersion         As Integer
        dmDriverVersion       As Integer
        dmSize                As Integer
        dmDriverExtra         As Integer
        dmFields              As Long
        dmOrientation         As Integer
        dmPaperSize           As Integer
        dmPaperLength         As Integer
        dmPaperWidth          As Integer
        dmScale               As Integer
        dmCopies              As Integer
        dmDefaultSource       As Integer
        dmPrintQuality        As Integer
        dmColor               As Integer
        dmDuplex              As Integer
        dmYResolution         As Integer
        dmTTOption            As Integer
        dmCollate             As Integer
        dmFormName            As String * CCHFORMNAME
        dmUnusedPadding       As Integer
        dmBitsPerPel          As Long
        dmPelsWidth           As Long
        dmPelsHeight          As Long
        dmDisplayFlags        As Long
        dmDisplayFrequency    As Long
End Type

Private Type typMonitor
   ScaleWidth     As Long
   ScaleHeight    As Long
End Type

Dim ScreenDC As Long                      'DC TO SCREEN
Dim Monitor As typMonitor                 'MONITOR INFORMATION

Dim IP As String                          'IP ADRESS

Dim Player As typPlayer                   'CHARACTERS
Dim Enemy() As typPlayer

Dim PlayerImage As typIMAGE_INFO          'BITMAPS
Dim EnemyImage As typIMAGE_INFO
Dim BulletImage As typIMAGE_INFO
Dim HealthBarImage As typIMAGE_INFO

Dim ArenaBuffer As typBuffer              'BUFFERS
Dim InfoBuffer As typBuffer

Dim BulletInfo() As typBullet             'BULLETS
Dim EnemyBulletInfo() As typEnemyBullet

Dim ptrBullet As Long                     '"POINTER" TO BULLET

Dim Continue As Boolean                   'SENDING COMPLETION CHECK
Dim Playing As Boolean                    'GAME LOOP CHECK

Dim TLastPause As Long                    'INTERVAL TO PAUSE

'***************************MISCELANIOS RUTINES**************************************

'OBJECT CREATION *********************

Private Function InitBuffer(ByVal DC As Long, ByVal Width As Long, ByVal Height As Long) As Long
   'create an object we can draw to off screen
   
   Dim RetVal As Long
   
   InitBuffer = CreateCompatibleDC(0)
   
   If InitBuffer = 0 Then
      MsgBox "cannot generate dc", vbOKOnly, "ERROR"
      Exit Function
   End If
   
   RetVal = CreateCompatibleBitmap(DC, Width, Height)
   
   If RetVal = 0 Then
      MsgBox "cannot create ArenaBuffer", vbOKOnly, "ERROR"
      DeleteDC InitBuffer
      
      Exit Function
   End If
   
   SelectObject InitBuffer, RetVal
   
   DeleteObject RetVal
End Function

Private Function GenerateDC(ByVal FILE_NAME As String, ByRef Pic As BITMAP) As Long
   'load and hold a picture in memory

   Dim RetVal As Long
   
   GenerateDC = CreateCompatibleDC(0)
   
   If GenerateDC = 0 Then
      MsgBox "cannot generate dc", vbOKOnly, "ERROR"
      Exit Function
   End If
   
   RetVal = LoadImage(0, App.Path & "\Bitmaps\" & FILE_NAME, IMAGE_BITMAP, 0, 0, LR_LOADFROMFILE Or LR_CREATEDIBSECTION)
   
   If RetVal = 0 Then
      MsgBox "cannot load image", vbOKOnly, "ERROR"
      DeleteDC GenerateDC
      
      Exit Function
   End If
   
   SelectObject GenerateDC, RetVal
   GetObject RetVal, Len(Pic), Pic
   
   DeleteObject RetVal
End Function

'RESOLUTION CHANGE*******************
Private Function ChangeRes(ByRef X As Long, ByRef Y As Long, ByRef Bits As Long) As Long
   'change resolution
   Dim devm As DEVMODE
   
   EnumDisplaySettings 0&, 0&, devm
   devm.dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT Or DM_BITSPERPEL
   devm.dmPelsWidth = X 'ScreenWidth
   devm.dmPelsHeight = Y 'ScreenHeight
   devm.dmBitsPerPel = Bits '(can be 8, 16, 24, 32 or even 4)
   
   ChangeDisplaySettings devm, CDS_TEST
End Function

Private Sub DoRes()
   'change resolution to 640 by 480
   ScreenDC = CreateDC("DISPLAY", vbNullString, vbNullString, ByVal 0&)
   ChangeRes 640, 480, GetDeviceCaps(ScreenDC, BITSPIXEL)
End Sub
   
Private Sub FixRes()
   'return res to original
   ChangeRes Monitor.ScaleWidth, Monitor.ScaleHeight, GetDeviceCaps(ScreenDC, BITSPIXEL)
   DeleteDC ScreenDC
End Sub

'TIME OPERATIONS*********************
Private Function FrameRate()
   'hold program so it runs at 60 fps on any computer
   Static HoldLastTime As Long
   
   Do While GetTickCount - HoldLastTime < 15
      If GetQueueStatus(QS_ALLEVENTS) <> 0 Then DoEvents
   Loop
   
   HoldLastTime = GetTickCount
End Function

Private Sub Pause()
   'pause the game to let user buy items
   If GetTickCount - TLastPause > 60000 Then
      mnuBuy_Click
      TLastPause = GetTickCount
      Send "/e "
      
      Do While GetTickCount - TLastPause < 10000
         If GetQueueStatus(QS_ALLEVENTS) <> 0 Then DoEvents
      Loop
      
      SFX "music"
      TLastPause = GetTickCount
      
      If fraVendor.Visible = True Then
         mnuBuy_Click
      End If
   End If
End Sub

'BUTTONS/BUYING ITEMS**********************
Private Sub mnuBuy_Click()
   vsrSheild.Min = 10 - Player.Info.Sheilds
   vsrHealth.Min = MAX_HEALTH - Player.Info.Health
   
   fraVendor.Visible = Not (fraVendor.Visible)
   lblHave.Caption = Player.Info.Money
End Sub

Private Sub chkArmour_Click()
   'select to buy an armour upgrade
   If chkArmour.Value = 1 Then
      lblCost.Caption = Val(lblCost.Caption) + 1000
      vsrSheild.Enabled = True
   ElseIf chkArmour = 0 Then
      lblCost.Caption = Val(lblCost.Caption) - 1000
      vsrSheild.Value = 0
      vsrSheild.Enabled = False
   End If
End Sub

Private Sub chkGun_Click(Index As Integer)
   'select to buy a gun upgrade
   If chkGun(Index).Value = 1 Then
      lblCost.Caption = Val(lblCost.Caption) + (Index + 1) * 400
      vsrBullets(Index).Enabled = True
   ElseIf chkGun(Index).Value = 0 Then
      lblCost.Caption = lblCost.Caption - (Index + 1) * 400
      vsrBullets(Index).Value = 0
      vsrBullets(Index).Enabled = False
   End If
End Sub

Private Sub chkRocket_Click()
   'speed up the user upgrade
   If chkRocket.Value = 1 Then
      lblCost.Caption = Val(lblCost.Caption) + 300
   ElseIf chkRocket.Value = 0 Then
      lblCost.Caption = Val(lblCost.Caption) - 300
   End If
End Sub

Private Sub vsrBullets_Change(Index As Integer)
   'amount of bullets to buy of gun type
   Static OldVal(1) As Long
   
   lblCost.Caption = lblCost.Caption + (10 + (10 * Index)) * (vsrBullets(Index).Value - OldVal(Index))
   OldVal(Index) = vsrBullets(Index).Value
   lblBullets(Index).Caption = vsrBullets(Index).Value
End Sub

Private Sub vsrHealth_Change()
   'buy back health points
   Static OldVal As Long
   
   lblHealth.Caption = vsrHealth.Value
   
   lblCost.Caption = lblCost.Caption + (vsrHealth.Value - OldVal) * 5
   
   OldVal = vsrHealth.Value
End Sub

Private Sub vsrSheild_Change()
   'select to buy sheilds
   Static OldVal As Long
   
   lblSheild.Caption = vsrSheild.Value
   
   lblCost.Caption = lblCost.Caption + (vsrSheild.Value - OldVal) * 40
   
   OldVal = vsrSheild.Value
End Sub

Private Sub cmdBuy_Click()
   'buy all selected items
   Dim I As Long
   
   If Player.Info.Money - Val(lblCost.Caption) >= 0 Then    'give items only if player can afford it
      Player.Info.Money = Player.Info.Money - Val(lblCost.Caption)
      
      If chkRocket.Value = 0 Then               'give the upgrade to player
         Player.Move.TBeforeMove = 100
      Else
         chkRocket.Value = 2
         Player.Move.TBeforeMove = 50
      End If
      
      If chkArmour.Value = 0 Then
         Player.Info.Armour = 0
      Else
         chkArmour.Value = 2
         Player.Info.Armour = 1
      End If
      
      For I = vsrBullets.LBound To vsrBullets.UBound
         If chkGun(I).Value = 0 Then
            Player.Info.Bullet(I + 1) = 0
         Else
            chkGun(I).Value = 2
            Player.Info.Bullet(I + 1) = Player.Info.Bullet(I + 1) + vsrBullets(I).Value
         End If
         
         vsrBullets(I).Value = 0
      Next
      
      Player.Info.Sheilds = Player.Info.Sheilds + vsrSheild.Value
      vsrSheild.Value = 0
      
      Player.Info.Health = Player.Info.Health + vsrHealth.Value
      vsrHealth.Value = 0
      
      lblCost.Caption = 0
      
      fraVendor.Visible = False
                                 'display new information to screen
      StretchBlt InfoBuffer.Back, 0, 88, picInfo.ScaleWidth - picInfo.ScaleWidth * ((50 - Player.Info.Health) / 50), LIFE_HEIGHT, HealthBarImage.DC, 0, 0, HealthBarImage.Info.bmWidth - HealthBarImage.Info.bmWidth * ((50 - Player.Info.Health) / 50), LIFE_HEIGHT, vbSrcCopy
      StretchBlt InfoBuffer.Back, picInfo.ScaleWidth - picInfo.ScaleWidth * ((50 - Player.Info.Health) / 50), 88, picInfo.ScaleWidth, LIFE_HEIGHT, HealthBarImage.DC, HealthBarImage.Info.bmWidth - HealthBarImage.Info.bmWidth * ((50 - Player.Info.Health) / 50), LIFE_HEIGHT, HealthBarImage.Info.bmWidth, LIFE_HEIGHT, vbSrcCopy
      WriteText InfoBuffer.Back, picInfo.ScaleWidth / 2, 220, Player.Info.Sheilds
      
      If ptrBullet <> 0 Then
         WriteText InfoBuffer.Back, picInfo.ScaleWidth / 2, 176, Player.Info.Bullet(ptrBullet)
      Else
         WriteText InfoBuffer.Back, picInfo.ScaleWidth / 2, 176, "%"
      End If
      
      WriteText InfoBuffer.Back, picInfo.ScaleWidth / 2, 132, Player.Info.Money
      WriteText InfoBuffer.Back, picInfo.ScaleWidth / 2, 264, Player.Info.Armour
   End If
End Sub

'MESSAGE SENDS**********************************
Private Function InitConnect() As String     'send information of player connecting or connected to new user
   InitConnect = "/i " & Player.Info.ID & " "
End Function

Private Function BulletMove(ByVal Ptr As Long) As String    'send when a bullet is moving to other players
   BulletMove = "/b " & Player.Info.ID & " " & Ptr & " " & BulletInfo(Ptr).Pos.Left & " " & BulletInfo(Ptr).Pos.Top & " " & BulletInfo(Ptr).Damage & " "
End Function

Private Function BulletFinished(ByVal Ptr As Long) As String   'bullet is at edge of screen stop from doing check/move
   BulletFinished = "/f " & Player.Info.ID & " " & Ptr & " "
End Function

Private Function PlayerData() As String         'send information to other players when player moves
   PlayerData = "/d " & Player.Info.ID & " " & Player.Pic.Direction & " " & Player.Pos.Left & " " & Player.Pos.Top & " "
End Function

Private Function Send(ByVal Text As String)     'send and wait till complete
   If CheckConnection Then
      wskBot.SendData Text
      
      Continue = True
      Do While Continue = True
         If GetQueueStatus(QS_ALLEVENTS) <> 0 Then DoEvents
         If wskBot.State <> sckConnected Then Continue = False
      Loop
   End If
End Function

Private Sub wskBot_SendComplete()         'leave the send loop when message is sent
   Continue = False
End Sub

'MESSAGE ARRIVAL*****************************
Private Function Section(ByRef PointMSG As Long, ByVal MSG As String) As String
   '***********
   'each piece of information is seperated by a space
   'section returns each piece of information before a space is found
   '***********
   Do While Mid(MSG, PointMSG, 1) <> " " And PointMSG <= Len(MSG)
      Section = Section & Mid(MSG, PointMSG, 1)
      PointMSG = PointMSG + 1
   Loop
   
   PointMSG = PointMSG + 1
End Function

Private Sub wskConnector_DataArrival(ByVal bytesTotal As Long)
   'check for connection to server and set port to connect with
   Dim ID As String
      
   wskConnector.GetData ID
   Player.Info.ID = ID
   
   Continue = False
End Sub

Private Function CheckConnection() As Boolean
   'check for connection with server and give error message if disconnected
   If wskBot.State = sckConnected Then
      CheckConnection = True
   Else
      CheckConnection = False
      mnuConnect.Enabled = True
      mnuDisconnect.Enabled = False
      
      If Player.Info.Connected = True Then
         MsgBox "You have been disconnected from the server", vbOKOnly, "Disconnect"
         Player.Info.Connected = False
      End If
   End If
End Function

'CONDITIONAL CHECKS***************************
Private Sub wskConnector_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
   'connection to server errors
   Select Case Number
      Case sckAddressInUse
         MsgBox "Adress in use try again soon", vbOKOnly Or vbCritical, "Error"
         wskConnector.Close
      Case sckConnectionRefused
         MsgBox "Wrong ip or" & vbNewLine & "Server is down", vbOKOnly Or vbCritical, "Error"
         wskConnector.Close
      Case 10065
         MsgBox "Bad IP entered", vbOKOnly Or vbCritical, "Error"
         wskConnector.Close
      Case Else
         MsgBox "unhandled error reset game", vbOKOnly Or vbCritical, "Error"
         wskConnector.Close
   End Select
End Sub

Private Sub wskBot_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
   'connection errors
   If Player.Info.Connected = True Then
      MsgBox "Unhandled error", vbOKOnly Or vbCritical, "Error"
      Player.Info.Connected = False
   End If
End Sub

Private Function CheckRect(ByRef Rect1 As RECT, ByRef Rect2 As RECT, ByVal X As Long, ByVal Y As Long) As Long
   'check if one rect is inside another
   Dim DestRect As RECT
   Dim TempRect As RECT
      
   SetRect TempRect, Rect1.Left, Rect1.Top, Rect1.Right, Rect1.Bottom
   OffsetRect TempRect, X, Y
   
   IntersectRect DestRect, TempRect, Rect2
   CheckRect = IsRectEmpty(DestRect)
End Function

Private Function CheckPlayer(ByVal X As Long, ByVal Y As Long) As Long
   'check if player is touching any other players
   Dim I As Long

   For I = LBound(Enemy) To UBound(Enemy)
      If Enemy(I).Info.Connected = True Then
         CheckPlayer = CheckRect(Player.Pos, Enemy(I).Pos, X, Y)
         If CheckPlayer = 0 Then Exit Function
      End If
   Next
End Function

'GRAPHICS/VISUALS***************************************************

Private Function OffSetPic(ByRef Pic As RECT, ByVal DC As Long, ByVal Xmov As Long, ByVal Ymov As Long, ByVal Width As Long, ByVal Height As Long, ByVal Direction As Long, ByVal Frame As Long)
   'move and redraw player/bullet
   BitBlt ArenaBuffer.Back, Pic.Left, Pic.Top, Width, Height, ArenaBuffer.Clean, Pic.Left, Pic.Top, vbSrcCopy
   
   OffsetRect Pic, Xmov, Ymov
      
   BitBlt ArenaBuffer.Back, Pic.Left, Pic.Top, Width, Height, DC, Frame * Width + 3 * Width, Direction * Height, vbSrcAnd
   BitBlt ArenaBuffer.Back, Pic.Left, Pic.Top, Width, Height, DC, Frame * Width, Direction * Height, vbSrcPaint
End Function

Private Sub FrameChange()
   'switch to next animation frame
   If Player.Pic.Frame < 2 Then
      Player.Pic.Frame = Player.Pic.Frame + 1
   Else
      Player.Pic.Frame = 0
   End If
End Sub

Private Sub WriteText(ByVal DC As Long, ByVal X As Long, ByVal Y As Long, ByVal Text As String)
   'write text
   Dim TextRect As RECT
   
   TextRect.Left = X       'set position to write
   TextRect.Top = Y
   
   BitBlt DC, 0, Y, picInfo.ScaleWidth, 22, InfoBuffer.Clean, 0, Y, vbSrcCopy    'remove old text
   DrawText DC, Text, Len(Text), TextRect, DT_CALCRECT                           'set the rectangle to proper size
   TextRect.Left = TextRect.Left - (TextRect.Right - TextRect.Left) / 2          'write text in rectangle
   DrawText DC, Text, Len(Text), TextRect, 0&                                    'write text to device context
   
   BitBlt picInfo.hdc, 0, 0, picInfo.ScaleWidth, picInfo.ScaleHeight, InfoBuffer.Back, 0, 0, vbSrcCopy   'display on screen
End Sub

Private Sub picArena_Paint()        'if part of the image is covered up redraw it
   BitBlt picArena.hdc, 0, 0, picArena.ScaleWidth, picArena.ScaleHeight, ArenaBuffer.Back, 0, 0, vbSrcCopy
   BitBlt picInfo.hdc, 0, 0, picInfo.ScaleWidth, picInfo.ScaleHeight, InfoBuffer.Back, 0, 0, vbSrcCopy
End Sub

Private Sub picInfo_Paint()         'if part of the image is covered up redraw it
   BitBlt picInfo.hdc, 0, 0, picInfo.ScaleWidth, picInfo.ScaleHeight, InfoBuffer.Back, 0, 0, vbSrcCopy
   BitBlt picArena.hdc, 0, 0, picArena.ScaleWidth, picArena.ScaleHeight, ArenaBuffer.Back, 0, 0, vbSrcCopy
End Sub

'SOUND*****************************************
Private Sub SFX(ByVal FILE_NAME As String)         'play sound
   sndPlaySound App.Path & "\Sounds\" & FILE_NAME & ".wav", SND_ASYNC Or SND_NODEFAULT
End Sub

'ITEM EFFECTS******************************
Private Sub DoSheilds()
   'rais armour and leave for 10 seconds per sheild used
   If Player.Info.ShieldOn > 0 Then
      If GetTickCount - Player.Info.SheildsHoldTime > 10000 Then
         Player.Info.ShieldOn = Player.Info.ShieldOn - 1
         Player.Info.SheildsHoldTime = GetTickCount
         
         If Player.Info.ShieldOn = 0 Then
            Player.Info.Armour = Player.Info.Armour - 10
            WriteText InfoBuffer.Back, picInfo.ScaleWidth / 2, 264, Player.Info.Armour
         End If
      End If
   End If
End Sub

Private Sub DoBullet()
   'move a bullet that has been shot
   Dim I As Long
   
   For I = LBound(BulletInfo) To UBound(BulletInfo)
      With BulletInfo(I)
         If .Fired = True Then
            If GetTickCount - .TLastMove >= 1 Then    'move once every millisecond
               .TLastMove = GetTickCount
               
               Select Case BulletInfo(I).Direction    'move in properdirection
                  Case 0
                     If BulletInfo(I).Pos.Bottom > 0 Then      'draw bullet moving
                        OffSetPic BulletInfo(I).Pos, BulletImage.DC, 0, -BulletInfo(I).MoveRate, BulletImage.Info.bmWidth / 2, BulletImage.Info.bmHeight, 0, 0
                        Send BulletMove(I)                     'send information of bullet to other players
                     Else     'when the bullet reaches the edge of the playing field
                        BulletInfo(I).Fired = False            'destroy bullet
                        Send BulletFinished(I)                 'tell other players bullet is destroyed
                     End If
                  Case 1
                     If BulletInfo(I).Pos.Left < picArena.ScaleWidth Then
                        OffSetPic BulletInfo(I).Pos, BulletImage.DC, BulletInfo(I).MoveRate, 0, BulletImage.Info.bmWidth / 2, BulletImage.Info.bmHeight, 0, 0
                        Send BulletMove(I)
                     Else
                        BulletInfo(I).Fired = False
                        Send BulletFinished(I)
                     End If
                  Case 2
                     If BulletInfo(I).Pos.Top < picArena.ScaleHeight Then
                        OffSetPic BulletInfo(I).Pos, BulletImage.DC, 0, BulletInfo(I).MoveRate, BulletImage.Info.bmWidth / 2, BulletImage.Info.bmHeight, 0, 0
                        Send BulletMove(I)
                     Else
                        BulletInfo(I).Fired = False
                        Send BulletFinished(I)
                     End If
                  Case 3
                     If BulletInfo(I).Pos.Right > 0 Then
                        OffSetPic BulletInfo(I).Pos, BulletImage.DC, -BulletInfo(I).MoveRate, 0, BulletImage.Info.bmWidth / 2, BulletImage.Info.bmHeight, 0, 0
                        Send BulletMove(I)
                     Else
                        BulletInfo(I).Fired = False
                        Send BulletFinished(I)
                     End If
               End Select
            End If
         End If
      End With
   Next
                                                      'draw to screen
   BitBlt picArena.hdc, 0, 0, picArena.ScaleWidth, picArena.ScaleHeight, ArenaBuffer.Back, 0, 0, vbSrcCopy
End Sub

'PLAYER ACTIONS****************************************
Private Sub KeyBoardEvent()
   'get player input
   
   'directions:
   '        0 = north
   '        1 = east
   '        2 = south
   '        3 = west
   
   Static Bullet As Long
   Static HoldTime As Long

   If GetKeyState(vbKeyEscape) < 0 Then      'disconnect from server
      wskBot.Close
   End If
   
   If Player.Info.Connected = True And Player.Info.Health > 0 Then
      If GetTickCount - Player.Move.TLastChangeGun > Player.Move.TBeforeChangeGun Then
         If GetKeyState(vbKeyControl) < 0 Then                 'switch to another gun
            Player.Move.TLastChangeGun = GetTickCount
            
            If ptrBullet < 2 Then
               ptrBullet = ptrBullet + 1
               WriteText InfoBuffer.Back, picInfo.ScaleWidth / 2, 176, Player.Info.Bullet(ptrBullet)
            Else
               ptrBullet = 0
               WriteText InfoBuffer.Back, picInfo.ScaleWidth / 2, 176, "%"
            End If
         ElseIf GetKeyState(vbKey1) < 0 Then
            ptrBullet = 0
            WriteText InfoBuffer.Back, picInfo.ScaleWidth / 2, 176, "%"
         ElseIf GetKeyState(vbKey2) < 0 Then
            ptrBullet = 1
            WriteText InfoBuffer.Back, picInfo.ScaleWidth / 2, 176, Player.Info.Bullet(ptrBullet)
         ElseIf GetKeyState(vbKey3) < 0 Then
            ptrBullet = 2
            WriteText InfoBuffer.Back, picInfo.ScaleWidth / 2, 176, Player.Info.Bullet(ptrBullet)
         End If
      End If
      
      If GetKeyState(vbKeyReturn) < 0 Then               'turn shields on
         If GetTickCount - Player.Move.TLastChangeGun > Player.Move.TBeforeChangeGun Then
            If Player.Info.Sheilds > 0 Then
               If Player.Info.ShieldOn = 0 Then Player.Info.SheildsHoldTime = GetTickCount
               Player.Move.TLastChangeGun = GetTickCount
               Player.Info.Sheilds = Player.Info.Sheilds - 1
               
               WriteText InfoBuffer.Back, picInfo.ScaleWidth / 2, 220, Player.Info.Sheilds
               If Player.Info.Armour < 10 Then Player.Info.Armour = Player.Info.Armour + 10
               Player.Info.ShieldOn = Player.Info.ShieldOn + 1
               WriteText InfoBuffer.Back, picInfo.ScaleWidth / 2, 264, Player.Info.Armour
            End If
         End If
      End If
            
      If GetKeyState(vbKeyUp) < 0 Then          'move player forward depending on direction they face
         If GetTickCount - Player.Move.TLastMove > Player.Move.TBeforeMove Then
            Player.Move.TLastMove = GetTickCount
            
            Select Case Player.Pic.Direction
               Case 0
                  If Player.Pos.Top - Player.Move.MoveRate >= 0 And CheckPlayer(0, -Player.Move.MoveRate) <> 0 Then
                     OffSetPic Player.Pos, PlayerImage.DC, 0, -Player.Move.MoveRate, PLAYER_WIDTH, PLAYER_HEIGHT, Player.Pic.Direction, Player.Pic.Frame
                     Send PlayerData
                  End If
               Case 1
                  If Player.Pos.Right + Player.Move.MoveRate <= picArena.ScaleWidth And CheckPlayer(Player.Move.MoveRate, 0) <> 0 Then
                     OffSetPic Player.Pos, PlayerImage.DC, Player.Move.MoveRate, 0, PLAYER_WIDTH, PLAYER_HEIGHT, Player.Pic.Direction, Player.Pic.Frame
                     Send PlayerData
                  End If
               Case 2
                  If Player.Pos.Bottom + Player.Move.MoveRate <= picArena.ScaleHeight And CheckPlayer(0, Player.Move.MoveRate) <> 0 Then
                     OffSetPic Player.Pos, PlayerImage.DC, 0, Player.Move.MoveRate, PLAYER_WIDTH, PLAYER_HEIGHT, Player.Pic.Direction, Player.Pic.Frame
                     Send PlayerData
                  End If
               Case 3
                  If Player.Pos.Left - Player.Move.MoveRate >= 0 And CheckPlayer(-Player.Move.MoveRate, 0) <> 0 Then
                     OffSetPic Player.Pos, PlayerImage.DC, -Player.Move.MoveRate, 0, PLAYER_WIDTH, PLAYER_HEIGHT, Player.Pic.Direction, Player.Pic.Frame
                     Send PlayerData
                  End If
            End Select
            
            FrameChange
         End If
      End If
         
      If GetKeyState(vbKeyRight) < 0 Then    'turn player to the right
         If GetTickCount - Player.Move.TLastTurn > Player.Move.TBeforeTurn Then
            Player.Move.TLastTurn = GetTickCount
            
            If Player.Pic.Direction < 3 Then
               Player.Pic.Direction = Player.Pic.Direction + 1
               OffSetPic Player.Pos, PlayerImage.DC, 0, 0, PLAYER_WIDTH, PLAYER_HEIGHT, Player.Pic.Direction, Player.Pic.Frame
               Send PlayerData
            Else
               Player.Pic.Direction = 0
               OffSetPic Player.Pos, PlayerImage.DC, 0, 0, PLAYER_WIDTH, PLAYER_HEIGHT, Player.Pic.Direction, Player.Pic.Frame
               Send PlayerData
            End If
         End If
      End If
      
      If GetKeyState(vbKeyDown) < 0 Then     'move backwards depending on what direction player if facing
         If GetTickCount - Player.Move.TLastMove > Player.Move.TBeforeMove Then
            Player.Move.TLastMove = GetTickCount
         
            Select Case Player.Pic.Direction
               Case 2
                  If Player.Pos.Top - Player.Move.MoveRate >= 0 And CheckPlayer(0, -Player.Move.MoveRate) <> 0 Then
                     OffSetPic Player.Pos, PlayerImage.DC, 0, -Player.Move.MoveRate, PLAYER_WIDTH, PLAYER_HEIGHT, Player.Pic.Direction, Player.Pic.Frame
                     Send PlayerData
                  End If
               Case 3
                  If Player.Pos.Right + Player.Move.MoveRate <= picArena.ScaleWidth And CheckPlayer(Player.Move.MoveRate, 0) <> 0 Then
                     OffSetPic Player.Pos, PlayerImage.DC, Player.Move.MoveRate, 0, PLAYER_WIDTH, PLAYER_HEIGHT, Player.Pic.Direction, Player.Pic.Frame
                     Send PlayerData
                  End If
               Case 0
                  If Player.Pos.Bottom + Player.Move.MoveRate <= picArena.ScaleHeight And CheckPlayer(0, Player.Move.MoveRate) <> 0 Then
                     OffSetPic Player.Pos, PlayerImage.DC, 0, Player.Move.MoveRate, PLAYER_WIDTH, PLAYER_HEIGHT, Player.Pic.Direction, Player.Pic.Frame
                     Send PlayerData
                  End If
               Case 1
                  If Player.Pos.Left - Player.Move.MoveRate >= 0 And CheckPlayer(-Player.Move.MoveRate, 0) <> 0 Then
                     OffSetPic Player.Pos, PlayerImage.DC, -Player.Move.MoveRate, 0, PLAYER_WIDTH, PLAYER_HEIGHT, Player.Pic.Direction, Player.Pic.Frame
                     Send PlayerData
                  End If
            End Select
            
            FrameChange
         End If
      End If
      
      If GetKeyState(vbKeyLeft) < 0 Then     'turn player to left
         If GetTickCount - Player.Move.TLastTurn > Player.Move.TBeforeTurn Then
            Player.Move.TLastTurn = GetTickCount
            
            If Player.Pic.Direction > 0 Then
               Player.Pic.Direction = Player.Pic.Direction - 1
               OffSetPic Player.Pos, PlayerImage.DC, 0, 0, PLAYER_WIDTH, PLAYER_HEIGHT, Player.Pic.Direction, Player.Pic.Frame
               Send PlayerData
            Else
               Player.Pic.Direction = 3
               OffSetPic Player.Pos, PlayerImage.DC, 0, 0, PLAYER_WIDTH, PLAYER_HEIGHT, Player.Pic.Direction, Player.Pic.Frame
               Send PlayerData
            End If
         End If
      End If
      
      If GetKeyState(vbKeyHome) < 0 Then        'side step to the left
         If GetTickCount - Player.Move.TLastMove > Player.Move.TBeforeMove Then
            Player.Move.TLastMove = GetTickCount
            
            Select Case Player.Pic.Direction
               Case 0
                  If Player.Pos.Left - Player.Move.MoveRate >= 0 And CheckPlayer(-Player.Move.MoveRate, 0) <> 0 Then
                     OffSetPic Player.Pos, PlayerImage.DC, -Player.Move.MoveRate, 0, PLAYER_WIDTH, PLAYER_HEIGHT, Player.Pic.Direction, Player.Pic.Frame
                     Send PlayerData
                  End If
               Case 1
                  If Player.Pos.Top - Player.Move.MoveRate >= 0 And CheckPlayer(0, -Player.Move.MoveRate) <> 0 Then
                     OffSetPic Player.Pos, PlayerImage.DC, 0, -Player.Move.MoveRate, PLAYER_WIDTH, PLAYER_HEIGHT, Player.Pic.Direction, Player.Pic.Frame
                     Send PlayerData
                  End If
               Case 2
                  If Player.Pos.Right + Player.Move.MoveRate <= picArena.ScaleWidth And CheckPlayer(Player.Move.MoveRate, 0) <> 0 Then
                     OffSetPic Player.Pos, PlayerImage.DC, Player.Move.MoveRate, 0, PLAYER_WIDTH, PLAYER_HEIGHT, Player.Pic.Direction, Player.Pic.Frame
                     Send PlayerData
                  End If
               Case 3
                  If Player.Pos.Bottom + Player.Move.MoveRate <= picArena.ScaleHeight And CheckPlayer(0, Player.Move.MoveRate) <> 0 Then
                     OffSetPic Player.Pos, PlayerImage.DC, 0, Player.Move.MoveRate, PLAYER_WIDTH, PLAYER_HEIGHT, Player.Pic.Direction, Player.Pic.Frame
                     Send PlayerData
                  End If
            End Select
            
            FrameChange
         End If
      End If
      
      If GetKeyState(vbKeyPageUp) < 0 Then      'side step to the right
         If GetTickCount - Player.Move.TLastMove > Player.Move.TBeforeMove Then
            Player.Move.TLastMove = GetTickCount
            
            Select Case Player.Pic.Direction
               Case 2
                  If Player.Pos.Left - Player.Move.MoveRate >= 0 And CheckPlayer(-Player.Move.MoveRate, 0) <> 0 Then
                     OffSetPic Player.Pos, PlayerImage.DC, -Player.Move.MoveRate, 0, PLAYER_WIDTH, PLAYER_HEIGHT, Player.Pic.Direction, Player.Pic.Frame
                     Send PlayerData
                  End If
               Case 3
                  If Player.Pos.Top - Player.Move.MoveRate >= 0 And CheckPlayer(0, -Player.Move.MoveRate) <> 0 Then
                     OffSetPic Player.Pos, PlayerImage.DC, 0, -Player.Move.MoveRate, PLAYER_WIDTH, PLAYER_HEIGHT, Player.Pic.Direction, Player.Pic.Frame
                     Send PlayerData
                  End If
               Case 0
                  If Player.Pos.Right + Player.Move.MoveRate <= picArena.ScaleWidth And CheckPlayer(Player.Move.MoveRate, 0) <> 0 Then
                     OffSetPic Player.Pos, PlayerImage.DC, Player.Move.MoveRate, 0, PLAYER_WIDTH, PLAYER_HEIGHT, Player.Pic.Direction, Player.Pic.Frame
                     Send PlayerData
                  End If
               Case 1
                  If Player.Pos.Bottom + Player.Move.MoveRate <= picArena.ScaleHeight And CheckPlayer(0, Player.Move.MoveRate) <> 0 Then
                     OffSetPic Player.Pos, PlayerImage.DC, 0, Player.Move.MoveRate, PLAYER_WIDTH, PLAYER_HEIGHT, Player.Pic.Direction, Player.Pic.Frame
                     Send PlayerData
                  End If
            End Select
            
            FrameChange
         End If
      End If
      
      If GetKeyState(vbKeySpace) < 0 Then          'shoot bullet in direction facing
         If GetTickCount - HoldTime > 1001 - 500 * ptrBullet Then
            If Player.Info.Bullet(ptrBullet) > 0 Then
               HoldTime = GetTickCount
               
               If ptrBullet = 0 Then               'infinate bullets
                  For Bullet = LBound(BulletInfo) To UBound(BulletInfo)
                     If Not (BulletInfo(Bullet).Fired) Then       'select dead bullet
                        Exit For                                  'so we dont have to resize array
                     End If
                     
                     If Bullet = UBound(BulletInfo) Then    'resize array of bullets if needed
                        Bullet = Bullet + 1
                        ReDim Preserve BulletInfo(Bullet)
                        
                        Exit For
                     End If
                  Next
                  
                  BulletInfo(Bullet).MoveRate = 2           'set bullet properties
                  BulletInfo(Bullet).Damage = 2
                  BulletInfo(Bullet).Fired = True
               ElseIf Player.Info.Bullet(ptrBullet) > 0 Then      'buyable bullets
                  For Bullet = LBound(BulletInfo) To UBound(BulletInfo)
                     If Not (BulletInfo(Bullet).Fired) Then
                        Exit For                               'select dead bullet to avoid resize of array
                     End If
                     
                     If Bullet = UBound(BulletInfo) Then       'resize array of bullets if needed
                        Bullet = Bullet + 1
                        ReDim Preserve BulletInfo(Bullet)
                        
                        Exit For
                     End If
                  Next
                  
                  BulletInfo(Bullet).MoveRate = 2 * ptrBullet              'set enhanced bullet properties
                  Player.Info.Bullet(ptrBullet) = Player.Info.Bullet(ptrBullet) - 1
                  BulletInfo(Bullet).Damage = 8 - 3 * ptrBullet
                  BulletInfo(Bullet).Fired = True
                  WriteText InfoBuffer.Back, picInfo.ScaleWidth / 2, 176, Player.Info.Bullet(ptrBullet)
               End If
               
               If BulletInfo(Bullet).Fired Then                'set to fire and tell other players bullet is fired
                  BulletInfo(Bullet).Direction = Player.Pic.Direction
                  SFX "Shoot"
                  Send "/s " & Player.Info.ID & " " & Bullet & " "
                  
                  Select Case BulletInfo(Bullet).Direction
                     Case 0
                        SetRect BulletInfo(Bullet).Pos, Player.Pos.Left + PLAYER_WIDTH / 2 - BulletImage.Info.bmWidth / 4, Player.Pos.Top - BulletImage.Info.bmHeight, Player.Pos.Left + PLAYER_WIDTH / 2 + BulletImage.Info.bmWidth / 4, Player.Pos.Top
                     Case 1
                        SetRect BulletInfo(Bullet).Pos, Player.Pos.Right, Player.Pos.Top + PLAYER_HEIGHT / 2 - BulletImage.Info.bmHeight / 2, Player.Pos.Right + BulletImage.Info.bmWidth / 2, Player.Pos.Top + PLAYER_HEIGHT / 2 + BulletImage.Info.bmHeight / 2
                     Case 2
                        SetRect BulletInfo(Bullet).Pos, Player.Pos.Left + PLAYER_WIDTH / 2 - BulletImage.Info.bmWidth / 4, Player.Pos.Bottom, Player.Pos.Left + PLAYER_WIDTH / 2 + BulletImage.Info.bmWidth / 4, Player.Pos.Bottom + BulletImage.Info.bmHeight
                     Case 3
                         SetRect BulletInfo(Bullet).Pos, Player.Pos.Left - BulletImage.Info.bmWidth / 2, Player.Pos.Top + PLAYER_HEIGHT / 2 - BulletImage.Info.bmHeight / 2, Player.Pos.Left, Player.Pos.Top + PLAYER_HEIGHT / 2 + BulletImage.Info.bmHeight / 2
                  End Select
               End If
            End If
         End If
      End If
   End If
   
   BitBlt picArena.hdc, 0, 0, picArena.ScaleWidth, picArena.ScaleHeight, ArenaBuffer.Back, 0, 0, vbSrcCopy
End Sub

'ENEMY/ITEM/PLAYER INTERACTIONS  AND MESSAGES************************************
Private Sub DataArrival(Optional ByVal bytesTotal As Long)
   'handle information recieved from other users
   '/commands:
   '        /d = other player moved/turned
   '        /i = new user joined or player is recieving information on other players connected already
   '        /l = other player has left
   '        /s = other player shot a bullet
   '        /b = other player's bullet is moving
   '        /f = other player's bullet has reached edge of playing area
   '        /k = player's bullet has killed another player
   '        /h = player's bullet has hit another player
   '        /e = pause game
   
   Dim Message As String   'holds message sent to player
   Dim ptrMSG As Long      '"pointer" to position in message
   Dim Temp As Long        'temporary storage for misc. information
   Dim Temp2 As Long
   Dim RetVal As Long
   Dim I As Long           'counters
   Dim J As Long
   
   ptrMSG = 1           'set pointer to beggining of message
   wskBot.GetData Message  'get message
   
   Do While ptrMSG < Len(Message)      'process whole message
      Select Case Section(ptrMSG, Message)
         Case "/d"
            Temp = Section(ptrMSG, Message)
            
            For I = LBound(Enemy) To UBound(Enemy)       'move proper enemy to new location
               With Enemy(I)
                  If .Info.ID = Temp Then
                     .Pic.Direction = Section(ptrMSG, Message)
                     OffSetPic .Pos, EnemyImage.DC, Section(ptrMSG, Message) - .Pos.Left, Section(ptrMSG, Message) - .Pos.Top, PLAYER_WIDTH, PLAYER_HEIGHT, .Pic.Direction, .Pic.Frame
                     
                     If .Pic.Frame < 2 Then
                        .Pic.Frame = .Pic.Frame + 1
                     Else
                        .Pic.Frame = 0
                     End If
                     
                     'Exit For
                  End If
               End With
            Next
         Case "/i"                              'resize array of enemies if needed
            Temp = Section(ptrMSG, Message)
            
            For I = LBound(Enemy) To UBound(Enemy)
               If Enemy(I).Info.ID = Temp And Enemy(I).Info.Connected = True Then
                  Temp = 0
                  Exit For
               End If
            Next
            
            If Temp <> 0 Then
               For I = LBound(Enemy) To UBound(Enemy)
                  If Enemy(I).Info.Connected = False Then
                     If I = UBound(Enemy) Then
                        ReDim Preserve Enemy(UBound(Enemy) + 1)
                        ReDim Preserve EnemyBulletInfo(UBound(Enemy))
                        ReDim Preserve EnemyBulletInfo(UBound(Enemy)).Info(0)
                     End If
                     
                     Enemy(I).Info.ID = Temp             'set enemy properties
                     Enemy(I).Info.Connected = True
                     SetRect Enemy(I).Pos, 0, 0, PLAYER_WIDTH, PLAYER_HEIGHT
                                          
                     Exit For
                  End If
               Next

               Send InitConnect
               Send PlayerData
            End If
         Case "/l"                              'resize array if needed
            Temp = Section(ptrMSG, Message)
                                       
            For I = LBound(Enemy) To UBound(Enemy)
               If Enemy(I).Info.ID = Temp Then
                  Enemy(I).Info.Connected = False
                                                'remove enemy from playing area
                  BitBlt ArenaBuffer.Back, Enemy(I).Pos.Left, Enemy(I).Pos.Top, PLAYER_WIDTH, PLAYER_HEIGHT, ArenaBuffer.Clean, Enemy(I).Pos.Left, Enemy(I).Pos.Top, vbSrcCopy
                  
                  If I = UBound(Enemy) - 1 Then
                     ReDim Preserve Enemy(UBound(Enemy) - 1)
                     ReDim Preserve EnemyBulletInfo(UBound(Enemy))
                  End If
                  
                  Exit For
               End If
            Next
         Case "/s"
            Temp = Section(ptrMSG, Message)
            Temp2 = Section(ptrMSG, Message)
            
            For I = LBound(Enemy) To UBound(Enemy)
               If Enemy(I).Info.ID = Temp Then                       'resize proper enemies bullet array if needed
                  For J = LBound(EnemyBulletInfo(I).Info) To UBound(EnemyBulletInfo(I).Info)
                     If EnemyBulletInfo(I).Info(J).Fired = False Then         'take dead bullet if possible
                        EnemyBulletInfo(I).Info(J).ID = Temp2
                        EnemyBulletInfo(I).Info(J).Fired = True
                        Exit For
                     ElseIf J = UBound(EnemyBulletInfo(I).Info) Then
                        ReDim Preserve EnemyBulletInfo(I).Info(J + 1)
                        EnemyBulletInfo(I).Info(UBound(EnemyBulletInfo(I).Info)).ID = Temp2
                        EnemyBulletInfo(I).Info(UBound(EnemyBulletInfo(I).Info)).Fired = True
                        Exit For
                     End If
                  Next
                  
                  SetRect EnemyBulletInfo(I).Info(J).Pos, EnemyBulletInfo(I).Info(J).Pos.Left, EnemyBulletInfo(I).Info(J).Pos.Top, EnemyBulletInfo(I).Info(J).Pos.Left + BulletImage.Info.bmWidth / 2, EnemyBulletInfo(I).Info(J).Pos.Top + BulletImage.Info.bmHeight
                  Exit For
               End If
            Next
         Case "/b"
            Temp = Section(ptrMSG, Message)
            Temp2 = Section(ptrMSG, Message)
            
            For I = LBound(Enemy) To UBound(Enemy)
               If Enemy(I).Info.ID = Temp Then
                  For J = LBound(EnemyBulletInfo(I).Info) To UBound(EnemyBulletInfo(I).Info)
                     If EnemyBulletInfo(I).Info(J).ID = Temp2 Then         'move proper bullet
                        OffSetPic EnemyBulletInfo(I).Info(J).Pos, BulletImage.DC, Section(ptrMSG, Message) - EnemyBulletInfo(I).Info(Temp2).Pos.Left, Section(ptrMSG, Message) - EnemyBulletInfo(I).Info(Temp2).Pos.Top, BulletImage.Info.bmWidth / 2, BulletImage.Info.bmHeight, 0, 0
                        Exit For
                     If J = UBound(EnemyBulletInfo(I).Info) Then Exit Sub
                     End If
                  Next
                  
                  If CheckRect(Player.Pos, EnemyBulletInfo(I).Info(J).Pos, 0, 0) = 0 And Player.Info.Health > 0 Then
                     RetVal = Val(Section(ptrMSG, Message))       'check if bullet hit player
                     Player.Info.Health = Player.Info.Health - IIf(RetVal - Player.Info.Armour > 0, RetVal - Player.Info.Armour, 0)
                     
                     If Player.Info.Health <= 0 Then              'send players death info and reset player information
                        Send "/k " & Enemy(I).Info.ID & " "
                        Player.Info.Lives = Player.Info.Lives - 1
                        WriteText InfoBuffer.Back, picInfo.ScaleWidth / 2, 44, Player.Info.Lives
                        SFX "Death"
                        BitBlt ArenaBuffer.Back, Player.Pos.Left, Player.Pos.Top, PLAYER_WIDTH, PLAYER_HEIGHT, ArenaBuffer.Clean, Player.Pos.Left, Player.Pos.Top, vbSrcCopy
                        SetPlayer
                        
                        StretchBlt InfoBuffer.Back, 0, 88, picInfo.ScaleWidth - picInfo.ScaleWidth * ((MAX_HEALTH - Player.Info.Health) / MAX_HEALTH), LIFE_HEIGHT, HealthBarImage.DC, 0, 0, HealthBarImage.Info.bmWidth - HealthBarImage.Info.bmWidth * ((MAX_HEALTH - Player.Info.Health) / MAX_HEALTH), LIFE_HEIGHT, vbSrcCopy
                        BitBlt picInfo.hdc, 0, 0, picInfo.ScaleWidth, picInfo.ScaleHeight, InfoBuffer.Back, 0, 0, vbSrcCopy

                        If Player.Info.Lives = 0 Then
                           Player.Info.Lives = 3
                           SFX "Lose"
                        End If
                     Else                                      'lose life
                        Send "/h " & Player.Info.ID & " " & Player.Info.Health & " "
                        
                        StretchBlt InfoBuffer.Back, 0, 88, picInfo.ScaleWidth - picInfo.ScaleWidth * ((MAX_HEALTH - Player.Info.Health) / MAX_HEALTH), LIFE_HEIGHT, HealthBarImage.DC, 0, 0, HealthBarImage.Info.bmWidth - HealthBarImage.Info.bmWidth * ((MAX_HEALTH - Player.Info.Health) / MAX_HEALTH), LIFE_HEIGHT, vbSrcCopy
                        StretchBlt InfoBuffer.Back, picInfo.ScaleWidth - picInfo.ScaleWidth * ((MAX_HEALTH - Player.Info.Health) / MAX_HEALTH), 88, picInfo.ScaleWidth, LIFE_HEIGHT, HealthBarImage.DC, HealthBarImage.Info.bmWidth - HealthBarImage.Info.bmWidth * ((MAX_HEALTH - Player.Info.Health) / MAX_HEALTH), LIFE_HEIGHT, HealthBarImage.Info.bmWidth, LIFE_HEIGHT, vbSrcCopy
                        BitBlt picInfo.hdc, 0, 0, picInfo.ScaleWidth, picInfo.ScaleHeight, InfoBuffer.Back, 0, 0, vbSrcCopy
                        
                        SFX "Shot"
                     End If
                  End If
                  
                  Exit For
               End If
            Next
         Case "/f"
            Temp = Section(ptrMSG, Message)
            Temp2 = Section(ptrMSG, Message)
            
            For I = LBound(Enemy) To UBound(Enemy)
               If Enemy(I).Info.ID = Temp Then
                  For J = LBound(EnemyBulletInfo(I).Info) To UBound(EnemyBulletInfo(I).Info)
                     If EnemyBulletInfo(I).Info(J).ID = Temp2 Then
                        EnemyBulletInfo(I).Info(J).Fired = False     'proper enemy bullet is dead
                        
                        Exit For
                     End If
                  Next
                  
                  Exit For
               End If
            Next
         Case "/k"
            If Player.Info.ID = Val(Section(ptrMSG, Message)) Then
                              'receive money for items
                              'depends on how many lives enemy had
               Player.Info.Money = Player.Info.Money + 100 * (Player.Info.Lives + 1)
               
               WriteText InfoBuffer.Back, picInfo.ScaleWidth / 2, 132, Player.Info.Money
            End If
         Case "/h"
            Temp = Section(ptrMSG, Message)
            
            For I = LBound(Enemy) To UBound(Enemy)          'lower enemy health
               If Enemy(I).Info.ID = Temp Then
                  Enemy(I).Info.Health = Section(ptrMSG, Message)
                  
                  Exit For
               End If
            Next
         Case "/e"
            TLastPause = GetTickCount           'pause game for 10 seconds
            mnuBuy_Click                        'bring up buy window with time to buy items
            
            Do While GetTickCount - TLastPause < 10000
               If GetQueueStatus(QS_ALLEVENTS) <> 0 Then DoEvents
            Loop
            
            SFX "music"
            TLastPause = 2 * GetTickCount
            
            If fraVendor.Visible = True Then
               mnuBuy_Click
            End If
      End Select
   Loop
   
   BitBlt picArena.hdc, 0, 0, picArena.ScaleWidth, picArena.ScaleHeight, ArenaBuffer.Back, 0, 0, vbSrcCopy
End Sub

'**************************MAIN GAME LOOP******************************
Private Sub Game()         'main game loop where every thing is called
   Static BulletHoldTime As Long
   
   Do While Playing
      KeyBoardEvent
      DoSheilds
      DoBullet
      Pause
      
      If wskBot.BytesReceived > 0 Then DataArrival
      If GetQueueStatus(QS_ALLEVENTS) <> 0 Then DoEvents
      
      FrameRate
   Loop
End Sub

'**************************INITIALIZATIONS************************************
Private Sub Form_Initialize()       'set dynamic arrays so ubound/lbound does not give error
   ReDim Enemy(0)
   ReDim BulletInfo(0)
   ReDim EnemyBulletInfo(0)
   ReDim EnemyBulletInfo(0).Info(0)
End Sub

Private Sub Form_Load()
   'display form before entering loop (must have this)
   Dim I As Long
   Dim Text As String
   Dim Dimentions(1) As Long
   
   Me.Show
               'original resolution is saved to file incase of crash
   If Dir(App.Path & "\Resolution.txt") <> "" Then       'check if file with original res exists
      Open App.Path & "\Resolution.txt" For Input As #1
         Do While Not (EOF(1)) And I < 2
            Input #1, Text
            Dimentions(I) = CLng(Text)
            I = I + 1
         Loop
      Close #1
   Else                                                  'write the res to file if none exists
      Open App.Path & "\Resolution.txt" For Output As #1
         Write #1, Screen.Width / Screen.TwipsPerPixelX
         Write #1, Screen.Height / Screen.TwipsPerPixelY
      Close #1
   End If

         'get/set resolution
   Monitor.ScaleWidth = Screen.Width / Screen.TwipsPerPixelX
   Monitor.ScaleHeight = Screen.Height / Screen.TwipsPerPixelY
   
   If Monitor.ScaleWidth <> Dimentions(0) Or Monitor.ScaleHeight <> Dimentions(1) Then
                  'if file info does not match current res then set the res to match the file
      Monitor.ScaleWidth = Dimentions(0)
      Monitor.ScaleHeight = Dimentions(1)
   End If
   
'   DoRes         commented out because of testing
'   Me.WindowState = vbMaximized

   picArena.Move 10, 10, Me.ScaleWidth - 110, Me.ScaleHeight - 20
   picInfo.Move Me.ScaleWidth - 100, 10, 90, Me.ScaleHeight - 20
   
            'load bitmap and save to memory
   PlayerImage.DC = GenerateDC("Player.bmp", PlayerImage.Info)
   EnemyImage.DC = GenerateDC("Enemy.bmp", EnemyImage.Info)
   BulletImage.DC = GenerateDC("Bullet.bmp", BulletImage.Info)
   HealthBarImage.DC = GenerateDC("LifeBar.bmp", HealthBarImage.Info)
   
            'create buffer to draw on
   InfoBuffer.hBlt = GenerateDC("InfoBack.bmp", InfoBuffer.Info)
   InfoBuffer.Back = InitBuffer(picInfo.hdc, picInfo.ScaleWidth, picInfo.ScaleHeight)
   InfoBuffer.Clean = InitBuffer(picInfo.hdc, picInfo.ScaleWidth, picInfo.ScaleHeight)
   
   ArenaBuffer.hBlt = GenerateDC("BackGround.bmp", ArenaBuffer.Info)
   ArenaBuffer.Back = InitBuffer(picArena.hdc, picArena.ScaleWidth, picArena.ScaleHeight)
   ArenaBuffer.Clean = InitBuffer(picArena.hdc, picArena.ScaleWidth, picArena.ScaleHeight)
   
            'stretch backgrounds to fit buffer/display to screen
   StretchBlt InfoBuffer.Clean, 0, 0, picInfo.ScaleWidth, picInfo.ScaleHeight, InfoBuffer.hBlt, 0, 0, InfoBuffer.Info.bmWidth, InfoBuffer.Info.bmHeight, vbSrcCopy
   BitBlt InfoBuffer.Back, 0, 0, picInfo.ScaleWidth, picInfo.ScaleHeight, InfoBuffer.Clean, 0, 0, vbSrcCopy
   BitBlt picInfo.hdc, 0, 0, picInfo.ScaleWidth, picInfo.ScaleHeight, InfoBuffer.Back, 0, 0, vbSrcCopy
   
   StretchBlt ArenaBuffer.Clean, 0, 0, picArena.ScaleWidth, picArena.ScaleHeight, ArenaBuffer.hBlt, 0, 0, ArenaBuffer.Info.bmWidth, ArenaBuffer.Info.bmHeight, vbSrcCopy
   BitBlt ArenaBuffer.Back, 0, 0, picArena.ScaleWidth, picArena.ScaleHeight, ArenaBuffer.Clean, 0, 0, vbSrcCopy
   BitBlt picArena.hdc, 0, 0, picArena.ScaleWidth, picArena.ScaleHeight, ArenaBuffer.Back, 0, 0, vbSrcCopy
   
            'set background of text box to transparent
   SetBkMode InfoBuffer.Back, TRANSPARENT
End Sub

Private Sub mnuConnect_Click()
   'get ip/name
   IP = InputBox("Enter IP to connect to", "IP adress", "0.0.0.0")
   Player.Info.Named = InputBox("Enter Your name", "Name", "PLAYER")
   
   'display players information
   WriteText InfoBuffer.Back, picInfo.ScaleWidth / 2, 0, Player.Info.Named
   WriteText InfoBuffer.Back, picInfo.ScaleWidth / 2, 22, "LIFE"
   WriteText InfoBuffer.Back, picInfo.ScaleWidth / 2, 44, 3
   
   WriteText InfoBuffer.Back, picInfo.ScaleWidth / 2, 66, "HEALTH"
   StretchBlt InfoBuffer.Back, 0, 88, picInfo.ScaleWidth - picInfo.ScaleWidth * ((50 - 50) / 50), LIFE_HEIGHT, HealthBarImage.DC, 0, 0, HealthBarImage.Info.bmWidth - HealthBarImage.Info.bmWidth * ((50 - 50) / 50), LIFE_HEIGHT, vbSrcCopy
   
   WriteText InfoBuffer.Back, picInfo.ScaleWidth / 2, 110, "MONEY"
   WriteText InfoBuffer.Back, picInfo.ScaleWidth / 2, 132, "0"
   
   WriteText InfoBuffer.Back, picInfo.ScaleWidth / 2, 154, "BULLETS"
   WriteText InfoBuffer.Back, picInfo.ScaleWidth / 2, 176, "%"
   
   WriteText InfoBuffer.Back, picInfo.ScaleWidth / 2, 198, "SHEILD"
   WriteText InfoBuffer.Back, picInfo.ScaleWidth / 2, 220, "0"
   
   WriteText InfoBuffer.Back, picInfo.ScaleWidth / 2, 242, "ARMOUR"
   WriteText InfoBuffer.Back, picInfo.ScaleWidth / 2, 264, "0"
   
   mnuConnect.Enabled = False
   mnuDisconnect.Enabled = True
   
   'connect to server
   Connect
End Sub

Private Sub Connect()
   wskConnector.Close
   wskConnector.Connect IP, "1"  'connect to server on defalut port(1)

   Continue = True
   Do While Continue          'wait for port to connect to
      If GetQueueStatus(QS_ALLEVENTS) <> 0 Then DoEvents
      If wskConnector.State = sckClosed Then Exit Sub
   Loop

   wskBot.Close
   wskBot.Connect IP, Str(Player.Info.ID)       'connect to proper port
   wskConnector.Close            'close default connection
End Sub

Private Sub wskBot_Connect()
   'set player start location and properties
   SetPlayer
   Player.Info.Lives = 3
   Player.Move.MoveRate = 10
   Player.Move.TBeforeMove = 100
   Player.Move.TBeforeTurn = 200
   Player.Move.TBeforeChangeGun = 200
   Player.Info.Bullet(0) = 1
   TLastPause = GetTickCount
   Player.Info.Money = 0
   
   Send InitConnect        'tell other players of your connection
   sndPlaySound App.Path & "\Sounds\Music.wav", SND_LOOP Or SND_ASYNC Or SND_NODEFAULT
   Playing = True
   Game        'enter game loop
End Sub

Private Sub SetPlayer()
   'set player start/spawn location and some properties
   If Player.Info.ID / 4 = Int(Player.Info.ID / 4) Then
      SetRect Player.Pos, picArena.ScaleWidth - PLAYER_WIDTH, picArena.ScaleHeight - PLAYER_HEIGHT, picArena.ScaleWidth, picArena.ScaleHeight
   ElseIf Player.Info.ID / 3 = Int(Player.Info.ID / 3) Then
      SetRect Player.Pos, 0, picArena.ScaleHeight - PLAYER_HEIGHT, PLAYER_WIDTH, picArena.ScaleHeight
   ElseIf Player.Info.ID / 2 = Int(Player.Info.ID / 2) Then
      SetRect Player.Pos, picArena.ScaleWidth - PLAYER_WIDTH, 0, picArena.ScaleWidth, PLAYER_HEIGHT
   Else
      SetRect Player.Pos, 0, 0, PLAYER_WIDTH, PLAYER_HEIGHT
   End If
   
   Player.Info.Health = MAX_HEALTH
   
               'draw player
   OffSetPic Player.Pos, PlayerImage.DC, 0, 0, PLAYER_WIDTH, PLAYER_HEIGHT, Player.Pic.Direction, Player.Pic.Frame
   Send PlayerData
   
   Player.Info.Connected = True
End Sub

'DISCONECTIONS/DEINITIALIZE
Private Sub mnuDisconnect_Click()      'disconnect from server
   If wskBot.State = sckConnected Then
      Send "/l " & Player.Info.ID & " "      'tell other players of your disconnection
   End If
   
   wskBot.Close
   mnuConnect.Enabled = True
   mnuDisconnect.Enabled = False
   Playing = False
End Sub

Private Sub mnuExit_Click()            'exit game
   Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If wskBot.State = sckConnected Then
      Send "/l " & Player.Info.ID & " "   'tell other players of your disconnection
   End If
                     'delete from memory all device contetxt to avoid memory leak(must have)
   If PlayerImage.DC <> 0 Then
      DeleteDC PlayerImage.DC
   End If
   
   If EnemyImage.DC <> 0 Then
      DeleteDC EnemyImage.DC
   End If
   
   If BulletImage.DC <> 0 Then
      DeleteDC BulletImage.DC
   End If
   
   If ArenaBuffer.Back <> 0 Then
      DeleteDC ArenaBuffer.Back
   End If
   
   If ArenaBuffer.Clean <> 0 Then
      DeleteDC ArenaBuffer.Clean
   End If
   
   If ArenaBuffer.hBlt <> 0 Then
      DeleteDC ArenaBuffer.hBlt
   End If
   
   If InfoBuffer.hBlt <> 0 Then
      DeleteDC InfoBuffer.hBlt
   End If
   
   If InfoBuffer.Back <> 0 Then
      DeleteDC InfoBuffer.Back
   End If
   
   If InfoBuffer.Clean <> 0 Then
      DeleteDC InfoBuffer.Clean
   End If
   
   If HealthBarImage.DC <> 0 Then
      DeleteDC HealthBarImage.DC
   End If
                  'stop any currently playing sounds
   sndPlaySound vbNullString, 0&
                  
'   FixRes         'set resolution back
   Playing = False   'exit game loop
   TLastPause = 0
End Sub

