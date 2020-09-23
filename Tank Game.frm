VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Tank Game"
   ClientHeight    =   6156
   ClientLeft      =   132
   ClientTop       =   876
   ClientWidth     =   7728
   LinkTopic       =   "Form1"
   ScaleHeight     =   513
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   644
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   252
      Left            =   240
      TabIndex        =   28
      Top             =   720
      Visible         =   0   'False
      Width           =   492
   End
   Begin VB.Timer Timer4 
      Interval        =   1
      Left            =   7440
      Top             =   480
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   7680
      Top             =   120
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      Height          =   288
      Left            =   1560
      TabIndex        =   27
      Text            =   "Text2"
      Top             =   120
      Visible         =   0   'False
      Width           =   1332
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   7800
      Top             =   840
   End
   Begin VB.CommandButton bb 
      BackColor       =   &H00000000&
      Enabled         =   0   'False
      Height          =   384
      Index           =   9
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   1200
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.CommandButton bb 
      BackColor       =   &H00000000&
      Enabled         =   0   'False
      Height          =   384
      Index           =   8
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   1200
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.CommandButton bb 
      BackColor       =   &H00000000&
      Enabled         =   0   'False
      Height          =   384
      Index           =   7
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   1200
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.CommandButton bb 
      BackColor       =   &H00000000&
      Enabled         =   0   'False
      Height          =   384
      Index           =   6
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   1200
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Timer btimer 
      Enabled         =   0   'False
      Index           =   9
      Interval        =   25
      Left            =   8400
      Top             =   1920
   End
   Begin VB.Timer btimer 
      Enabled         =   0   'False
      Index           =   8
      Interval        =   25
      Left            =   8400
      Top             =   1920
   End
   Begin VB.Timer btimer 
      Enabled         =   0   'False
      Index           =   7
      Interval        =   25
      Left            =   8280
      Top             =   2040
   End
   Begin VB.Timer btimer 
      Enabled         =   0   'False
      Index           =   6
      Interval        =   25
      Left            =   8280
      Top             =   2280
   End
   Begin VB.Timer ghost 
      Enabled         =   0   'False
      Index           =   1
      Interval        =   1000
      Left            =   7920
      Top             =   840
   End
   Begin VB.Timer btimer 
      Enabled         =   0   'False
      Index           =   5
      Interval        =   25
      Left            =   8280
      Top             =   1920
   End
   Begin VB.Timer btimer 
      Enabled         =   0   'False
      Index           =   4
      Interval        =   25
      Left            =   8280
      Top             =   1800
   End
   Begin VB.Timer btimer 
      Enabled         =   0   'False
      Index           =   3
      Interval        =   25
      Left            =   8280
      Top             =   1920
   End
   Begin VB.Timer btimer 
      Enabled         =   0   'False
      Index           =   2
      Interval        =   25
      Left            =   8400
      Top             =   1920
   End
   Begin VB.Timer btimer 
      Enabled         =   0   'False
      Index           =   1
      Interval        =   25
      Left            =   8400
      Top             =   1680
   End
   Begin VB.Timer btimer 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   25
      Left            =   8160
      Top             =   2040
   End
   Begin VB.CommandButton bb 
      BackColor       =   &H00000000&
      Enabled         =   0   'False
      Height          =   384
      Index           =   5
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   1320
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.CommandButton bb 
      BackColor       =   &H00000000&
      Enabled         =   0   'False
      Height          =   384
      Index           =   4
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   1320
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.CommandButton bb 
      BackColor       =   &H00000000&
      Enabled         =   0   'False
      Height          =   384
      Index           =   3
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   1320
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.CommandButton bb 
      BackColor       =   &H00000000&
      Enabled         =   0   'False
      Height          =   384
      Index           =   2
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   1320
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.CommandButton bb 
      BackColor       =   &H00000000&
      Enabled         =   0   'False
      Height          =   384
      Index           =   1
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   1320
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.CommandButton bb 
      BackColor       =   &H00000000&
      Enabled         =   0   'False
      Height          =   384
      Index           =   0
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   1320
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Timer gflash 
      Interval        =   150
      Left            =   8040
      Top             =   240
   End
   Begin VB.Timer ghost 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   1000
      Left            =   8040
      Top             =   840
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   9000
      Top             =   0
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   288
      Left            =   3240
      TabIndex        =   14
      Text            =   "Text1"
      Top             =   0
      Visible         =   0   'False
      Width           =   2892
   End
   Begin VB.Timer Timer 
      Enabled         =   0   'False
      Index           =   9
      Interval        =   25
      Left            =   8520
      Top             =   360
   End
   Begin VB.Timer Timer 
      Enabled         =   0   'False
      Index           =   8
      Interval        =   25
      Left            =   8520
      Top             =   360
   End
   Begin VB.Timer Timer 
      Enabled         =   0   'False
      Index           =   7
      Interval        =   25
      Left            =   8520
      Top             =   360
   End
   Begin VB.Timer Timer 
      Enabled         =   0   'False
      Index           =   6
      Interval        =   25
      Left            =   8520
      Top             =   360
   End
   Begin VB.Timer Timer 
      Enabled         =   0   'False
      Index           =   5
      Interval        =   25
      Left            =   8520
      Top             =   360
   End
   Begin VB.Timer Timer 
      Enabled         =   0   'False
      Index           =   4
      Interval        =   25
      Left            =   8520
      Top             =   360
   End
   Begin VB.Timer Timer 
      Enabled         =   0   'False
      Index           =   3
      Interval        =   25
      Left            =   8520
      Top             =   360
   End
   Begin VB.Timer Timer 
      Enabled         =   0   'False
      Index           =   2
      Interval        =   25
      Left            =   8520
      Top             =   360
   End
   Begin VB.Timer Timer 
      Enabled         =   0   'False
      Index           =   1
      Interval        =   25
      Left            =   8520
      Top             =   360
   End
   Begin VB.Timer Timer 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   25
      Left            =   8520
      Top             =   360
   End
   Begin VB.CommandButton bullet 
      BackColor       =   &H00000000&
      Enabled         =   0   'False
      Height          =   384
      Index           =   9
      Left            =   8832
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   768
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.CommandButton bullet 
      BackColor       =   &H00000000&
      Enabled         =   0   'False
      Height          =   384
      Index           =   8
      Left            =   8832
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   768
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.CommandButton bullet 
      BackColor       =   &H00000000&
      Enabled         =   0   'False
      Height          =   384
      Index           =   7
      Left            =   8832
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   768
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.CommandButton bullet 
      BackColor       =   &H00000000&
      Enabled         =   0   'False
      Height          =   384
      Index           =   6
      Left            =   8832
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   768
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.CommandButton bullet 
      BackColor       =   &H00000000&
      Enabled         =   0   'False
      Height          =   384
      Index           =   5
      Left            =   8832
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   768
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.CommandButton bullet 
      BackColor       =   &H00000000&
      Enabled         =   0   'False
      Height          =   384
      Index           =   4
      Left            =   8832
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   768
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.CommandButton bullet 
      BackColor       =   &H00000000&
      Enabled         =   0   'False
      Height          =   384
      Index           =   3
      Left            =   8832
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   768
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.CommandButton bullet 
      BackColor       =   &H00000000&
      Enabled         =   0   'False
      Height          =   384
      Index           =   2
      Left            =   8832
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   768
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.CommandButton bullet 
      BackColor       =   &H00000000&
      Enabled         =   0   'False
      Height          =   384
      Index           =   1
      Left            =   8832
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   768
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.CommandButton bullet 
      BackColor       =   &H00000000&
      Enabled         =   0   'False
      Height          =   384
      Index           =   0
      Left            =   8832
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   768
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.ListBox Chat 
      Appearance      =   0  'Flat
      CausesValidation=   0   'False
      Enabled         =   0   'False
      Height          =   1092
      IntegralHeight  =   0   'False
      Left            =   120
      MousePointer    =   12  'No Drop
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   4680
      Width           =   7572
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   600
      Top             =   120
      _ExtentX        =   995
      _ExtentY        =   995
      _Version        =   393216
      Protocol        =   4
      URL             =   "http://"
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   0
      Top             =   120
      _ExtentX        =   593
      _ExtentY        =   593
      _Version        =   393216
   End
   Begin VB.Image ghosty 
      Height          =   432
      Left            =   8880
      Picture         =   "Tank Game.frx":0000
      Top             =   4200
      Visible         =   0   'False
      Width           =   396
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Height          =   252
      Left            =   6480
      TabIndex        =   16
      Top             =   600
      Width           =   1332
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Height          =   252
      Left            =   6480
      TabIndex        =   15
      Top             =   240
      Width           =   1332
   End
   Begin VB.Image GuySkin 
      Height          =   432
      Index           =   4
      Left            =   8880
      Picture         =   "Tank Game.frx":0E52
      Top             =   1920
      Visible         =   0   'False
      Width           =   396
   End
   Begin VB.Image GuySkin 
      Height          =   396
      Index           =   3
      Left            =   8880
      Picture         =   "Tank Game.frx":1CA4
      Top             =   2520
      Visible         =   0   'False
      Width           =   432
   End
   Begin VB.Image GuySkin 
      Height          =   432
      Index           =   2
      Left            =   8880
      Picture         =   "Tank Game.frx":2AD2
      Top             =   3120
      Visible         =   0   'False
      Width           =   396
   End
   Begin VB.Image GuySkin 
      Height          =   396
      Index           =   1
      Left            =   9000
      Picture         =   "Tank Game.frx":3924
      Top             =   3720
      Visible         =   0   'False
      Width           =   432
   End
   Begin VB.Image picbackground 
      Height          =   768
      Index           =   0
      Left            =   8640
      Picture         =   "Tank Game.frx":4752
      Top             =   4920
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Label lblhosting 
      BackColor       =   &H00FF80FF&
      Caption         =   "Hosting"
      Height          =   252
      Left            =   7080
      TabIndex        =   2
      Top             =   5880
      Visible         =   0   'False
      Width           =   612
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFC0&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   6360
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   2892
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   640
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label lblConnected 
      BackColor       =   &H0000FF00&
      Caption         =   "Connected"
      Height          =   252
      Left            =   0
      TabIndex        =   0
      Top             =   5880
      Visible         =   0   'False
      Width           =   7092
   End
   Begin VB.Image Guy 
      Height          =   396
      Index           =   2
      Left            =   3480
      Picture         =   "Tank Game.frx":6D94
      Top             =   3480
      Width           =   432
   End
   Begin VB.Image Guy 
      Height          =   396
      Index           =   1
      Left            =   0
      Picture         =   "Tank Game.frx":7BC2
      Top             =   0
      Width           =   432
   End
   Begin VB.Menu mConnectHost 
      Caption         =   "Connect/Host"
      Begin VB.Menu mHost 
         Caption         =   "Host a game"
      End
      Begin VB.Menu mConnect 
         Caption         =   "Connect"
      End
      Begin VB.Menu mDisconnect 
         Caption         =   "Disconnect"
      End
      Begin VB.Menu mgetIp 
         Caption         =   "Get External IP"
      End
   End
   Begin VB.Menu mOption 
      Caption         =   "Options"
      Begin VB.Menu mPort 
         Caption         =   "Choose Port"
      End
      Begin VB.Menu atk 
         Caption         =   "Attack!"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mDrawLines 
         Caption         =   "Draw Lines"
      End
      Begin VB.Menu mSay 
         Caption         =   "Say"
         Shortcut        =   ^S
      End
      Begin VB.Menu mSkins 
         Caption         =   "Cloak"
         Shortcut        =   {F3}
      End
   End
   Begin VB.Menu mHelp 
      Caption         =   "Help"
      Begin VB.Menu mAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'SIMPLE MULTIPLAYER GAME
'Created by FST
Private outofbullets As Integer
Private datacount As Integer
Private character(30) As String

'Understanding of Cells
'This whole program is based on cells as you can notice. That means that i am creating
'squares on the form, where the guys can move. So i am limiting the positions they can be in
'This has its up and downs, down side is that you cannot see your guy nicely going from one
'position to another, he looks as though he is jumping there, appearing out of nowhere. But the
'good part about making cells is that you can monitor your guys, then later on in game you
'can add some objects in the form, stones, monsters, you get to know exact positions where they are
'and you dont have to use the RECT collision detection

'Also remember, winsock cant be perfect. If you have AOL you can lag. Connection, computer...
'some packets get lost, mixed up or even joined. We will discuss this later

'The next known bug is that when you start this application in Visual Basic and someone connects to you,
'then you try to say something, the text you send he wont see! BUt you will see the text he sends.
'I wasnt able to identify why this happens, but it is like that and i cant fix it
'So what you do when you want to use SAY feature appropriatly, create an EXE of this application
'and run that EXE, the SAY feature will be flawless when you run this as EXE...

'for FOR..NEXT loops
Dim score(1) As Integer 'score(0) = you, score(1) = them
Dim i As Integer, r As Integer, c As Integer
Dim dir As String
Dim strr As String
Dim direction(9) As Integer 'direction bullet was shot at, used 9 so i can easily change # of bullets
'positions of both guys
Dim GuyXpos(2) As Integer 'should have commented this while i was programming, i think i made guyXpos(1) you and 2 them
Dim GuyYpos(2) As Integer ' look above

'if lines option is selected
Dim lines As Boolean 'to draw lines for easy shooting

'winsock stuff
Dim port As Single, hostIP As String

'winsock data arrival Dims, they will be needed later on
Dim Data As String
Dim Data2() As String
Dim data3() As String

'temporary variable for messagebox input yes/no
Dim temp As Integer

'for graphics loop in painting the picture with grass
Dim x As Integer, y As Integer

'for chat
Private namez As String

'to make sure there are no variables that we are using that are not dimmed
Option Explicit

Sub DrawLines()
'this procedure is called if the user selects the option to draw squares...

'this For..Next loop does horizontal lines
For i = 1 To 20
Form1.Line (i * 32, 0)-(i * 32, 480)
Next i

'this For..Next loop does vertical lines
For i = 1 To 15
Form1.Line (0, i * 32)-(640, i * 32)
Next i
'the form is 640 X 480 big

'this boolean holds the current state... false = no lines, true = draw lines
lines = True
End Sub


Private Sub atk_Click()
Dim num As Integer 'number of bullet thats available
Dim count As Integer 'to see if there are bullets
For i = 0 To 5
If bullet(i).Left = Guy(1).Left And bullet(i).Top = Guy(1).Top And bullet(i).Visible = False Then 'if bullet is in same spot as your tank and hasnt been shot
num = i 'get number of that bullet
count = 1 'make sure that it wont output that ur out of bullets
End If
Next i
If count = 1 Then 'if u have bullets
bullet(num).Visible = True 'the bullet that u found LAST is shot
Text1.Text = num 'for testing purposes
Timer(num).Enabled = True 'makes the timer for speed of bullet start
direction(num) = dir 'direction = way ur facing, check out keydown
Winsock1.SendData "003|" 'tells other computer u shot a bullet
Else
Chat.RemoveItem (0) 'removes first thing of textbox so that it looks like its scrolling up... same as CHAT
outofbullets = outofbullets + 1 'amount of times that this has occured
Chat.AddItem "* Out of bullets *   (time #" & outofbullets & ")" 'outputs that ur out of bullets and how many times you have already tried to shoot when u had no bullets (just for their knowledge on how much of a waster they are)
End If
End Sub

Private Sub btimer_Timer(Index As Integer) 'NOTE: any word with B at begining is opponents (bad)
' this just checks their position so bullet goes right way
If bdirection(Index) = 1 Then
bb(Index).Top = bb(Index).Top - 32
ElseIf bdirection(Index) = 2 Then
bb(Index).Left = bb(Index).Left - 32
ElseIf bdirection(Index) = 3 Then
bb(Index).Top = bb(Index).Top + 32
Else
bb(Index).Left = bb(Index).Left + 32
End If
If bb(Index).Left < 0 Or bb(Index).Left > 20 * 32 Or bb(Index).Top < 0 Or bb(Index).Top > 12 * 32 Then 'if it goes off the screen then
'set its position to b on their tank and make it invisible... now its usable again
bb(Index).Visible = False
bb(Index).Left = Guy(2).Left
bb(Index).Top = Guy(2).Top
btimer(Index).Enabled = False
End If
End Sub

Private Sub Chat_Click()
Form1.SetFocus 'just incase they click on chat
End Sub

Private Sub Chat_GotFocus()
Form1.SetFocus 'incase they mess something up somehow, after all, they are bound to find a way
End Sub

Private Sub Chat_ItemCheck(Item As Integer)
Form1.SetFocus '...
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'this occurs when the user releases a key. We use this Select Case to find out which key
'the user released
rpt = 0
Select Case KeyCode
'if its key up
Case vbKeyUp

    'then we first check if the Yposition of the player isnt 1, therefore he cant go more up!
    If GuyYpos(1) <> 1 Then
    rpt = 1
    For i = 0 To 9
    'moves bullets along with them
If bullet(i).Top = Guy(1).Top And bullet(i).Left = Guy(1).Left And bullet(i).Visible = False Then
bullet(i).Top = bullet(i).Top - 32
End If
Next i
    'but if it is not, change his Y position one up
    GuyYpos(1) = GuyYpos(1) - 1 'changes actual position of character
    dir = 1 'makes direction be 1 (up), used for skins and bullet directions... same for other 3 instances
    End If
Case vbKeyRight

    'We check if the position of the player isnt already maximum of going right
    If GuyXpos(1) <> 20 Then
    rpt = 1
        'moves bullets along with them
    For i = 0 To 9
If bullet(i).Top = Guy(1).Top And bullet(i).Left = Guy(1).Left And bullet(i).Visible = False Then
bullet(i).Left = bullet(i).Left + 32
End If
Next i
    GuyXpos(1) = GuyXpos(1) + 1
    dir = 4
    End If
Case vbKeyDown

    'here we check if the player isnt at the bottom of the screen and he cant go any more down
    If GuyYpos(1) <> 12 Then
    rpt = 1
        'moves bullets along with them
    For i = 0 To 9
If bullet(i).Top = Guy(1).Top And bullet(i).Left = Guy(1).Left And bullet(i).Visible = False Then
bullet(i).Top = bullet(i).Top + 32
End If
Next i
GuyYpos(1) = GuyYpos(1) + 1
    dir = 3
    End If
Case vbKeyLeft

    'if the user pressed left, but his position is already minimum left value, we cant go anymore left
    If GuyXpos(1) <> 1 Then
    rpt = 1
        'moves bullets along with them
    For i = 0 To 9
If bullet(i).Top = Guy(1).Top And bullet(i).Left = Guy(1).Left And bullet(i).Visible = False Then
bullet(i).Left = bullet(i).Left - 32
End If
Next i
GuyXpos(1) = GuyXpos(1) - 1
    dir = 2
    End If
End Select
Guy(1).Picture = GuySkin(5 - dir).Picture 'changes skin

If Winsock1.State <> 0 Then
Winsock1.SendData "999" & "/" & 5 - dir & "|" 'send em the position he is now in and let the other player interpret it
'the less amount of data sent the quicker everything runs
End If
'here we move the actual guy to the updated X and Y positions
'we use the .Move command instead of setting .Left and .Top separatly which takes longer both to you
'and the computer to process it
Guy(1).Move GuyXpos(1) * 32 - 32, GuyYpos(1) * 32 - 32

'finally, we need to send the new data over to the other player, we use On Error Resume Next,
'just to make sure the program doesnt crash if the user moves and he is not connected yet
On Error Resume Next
    
'This line can be difficult to understand so let me explain
'Programming with winsock is hard. If you are constantly sending a lot of packets at the same time,
'sometimes they get mixed up, or joined. Most of time though they are ok, but sometimes it doesnt
'work out. Lets have an example to illustrate this.
'lets say
'X = 5
'y = 15
'Now we need to send these over to the other side. We do this by joining these two numbers and putting
'a dummie between them, so we are sending one string in fact, but then later on on dataarrival section
'we extract each information by using the Split() function. Therefore we could do this:
'Winsock1.SendData GuyXpos(1) & "/" & GuyYpos(1)
'it seems like a good syntax from what i told you so far. But remember ! I said above that the packets
'sometimes join. So we are sending 5/15 very fast and eventually sometimes the other guy gets a string
'5/155/15. See the difference ? Two packets are joined and the Y coordinate becomes 155!
'Therefore we add "|" in the end. Its importance will be documented in dataarrival section.
If rpt = 1 Then
Winsock1.SendData GuyXpos(1) & "/" & GuyYpos(1) & "|"
End If
For i = 0 To 9
If bullet(i).Visible = False Then
bullet(i).Move GuyXpos(1) * 32 - 32, GuyYpos(1) * 32 - 32
End If
Next i
End Sub

Private Sub Form_Load()

atk.Enabled = False 'makes sure u dont atk @ begining before u are connected
'sets all bullets at same position as u
For i = 0 To 9
bullet(i).Top = Guy(1).Top
bullet(i).Left = Guy(1).Left
bb(i).Top = Guy(2).Top
bb(i).Left = Guy(2).Left
Next i
'asks for name to be used in chat
namez = InputBox("What is your name?")
dir = 2
'makes chatbox blank, that way when a message is sent, top item is deleted and new item is placed on bottom, this makes an actual chat effect without making the forum lose focus
Chat.AddItem ""
Chat.AddItem ""
Chat.AddItem ""
Chat.AddItem ""
Chat.AddItem ""
'to make sure all numbers generated are TRULY random
Randomize 'added later: just realized that no random numbers are used
'we set the default port
port = 5432
'we set the default position for both characters
GuyXpos(1) = 1
GuyYpos(1) = 1
GuyXpos(2) = 10
GuyYpos(2) = 10

'we move these characters to the default positions that were given above
Guy(1).Move GuyXpos(1) * 32 - 32, GuyYpos(1) * 32 - 32
Guy(2).Move GuyXpos(2) * 32 - 32, GuyYpos(2) * 32 - 32

'this code is the code that creates the grass

    'first we make sure y is 0
    y = 0
    'then we cicle through the rows(r)
    For r = 1 To 14
        'everytime a row changes, we need to assign new X
       x = 0
        'we cycle through all columns
        For c = 1 To 20
            'we use this function to paint the form1 with the picture of grass
            Form1.PaintPicture picbackground(0).Picture, x, y
            'increase x by 32 (one square, or cell if you like)
           x = x + 32
        Next c
        'add 32 to Y
        y = y + 32
    Next r
    'note: u need to wait a while if u run that in debug mode! thats 14 * 20 = 280 cycles!
    'insure graphic presistance
    Form1.Picture = Form1.Image
Chat.Enabled = False

End Sub

Private Sub Form_Unload(Cancel As Integer)
'we want to close the winsock before we close the program
Winsock1.Close
End Sub




Private Sub gflash_Timer()
'this makes the flashing G when they are a ghost
If ghost(0).Enabled = True Then 'if they are actually a ghost
If Guy(2).Picture = ghosty Then 'if they show the G symbol
Guy(2).Picture = GuySkin(5 - dir) 'make their pic normal (later realized that this is YOUR skin... oh well, when they shoot they change to the correct direction anyways)
Else
Guy(2).Picture = ghosty
End If
End If
End Sub

Private Sub ghost_Timer(Index As Integer)
ghost(Index).Enabled = False 'after it finishes, disable this
If rempic <> 0 Then
Guy(2).Picture = GuySkin(5 - rempic).Picture
Else
Guy(2).Picture = GuySkin(4 - rempic).Picture
End If
End Sub


Private Sub mAbout_Click()
'if the user clicks About, display this message
MsgBox "This game was made by Yonatan Naamad." & vbCrLf & vbCrLf & "It is an open source two player tank game, in which you basically try to kill the other person.  If you find any bugs, or if you get an error and you know what caused it, email me @ cached@gmail.com and I'll fix it and add a thank-you to this program for u :).  Knock yourselves out, or should I say your opponents? :Þ" & vbCrLf & vbCrLf & "Current Thank Yous: none", , "About"
End Sub

Private Sub mConnect_Click()
Timer3.Enabled = True
atk.Enabled = True
On Error Resume Next
'If the user clicks Connect, we need to get the IP of the hosting computer first !
hostIP = InputBox("Enter the host's computer name or ip address:" & vbCrLf & "(Be careful not to include any unnecessary spaces etc or error message will be generated.")
'then we connect to this IP, on the default port
Winsock1.Connect hostIP, port
'we let the user know that he is connected


'we set up the starting default positions in a case that the user had moved before clicking connect
'notice, that in mHOST sub, we do the same thing, but we do it the other way,
'we let the GuyXpos(2) = 10, GuyYpos(2) = 10 and the GuyXpos(1) = 1 and GuyYpos(1) = 1
'This is a really hard part to explain, you really need to think about it. If you click connect,
'We set up your position 10, 10 and the host will be commanding the 1,1 guy
'In Host Sub, its vice versa, the Guy(1) is given the positions in 1,1, and the Guy2 in 10,10 !
GuyXpos(2) = 1
GuyYpos(2) = 1
GuyXpos(1) = 10
GuyYpos(1) = 10
'we move the pictures to updated positions
Guy(2).Move GuyXpos(2) * 32 - 32, GuyYpos(2) * 32 - 32
Guy(1).Move GuyXpos(1) * 32 - 32, GuyYpos(1) * 32 - 32

End Sub

Private Sub mDrawLines_Click()
'If the user clicks DrawLines

'and Lines are already Drawen (This option of DrawLines works as toggle so we
'need to figure out the previous state)
If lines = True Then
    'clear the lines and set lines = false
    lines = False
    Form1.Cls
Else
' if the lines is False however,
    'we set lines = true and we call DrawLines procedure to make the actual lines
    lines = True
    DrawLines
End If
End Sub

Private Sub mDisconnect_Click()
'all we need to do here really is just to close Winsock1
Winsock1.Close
'and we let the user know that he is not hosting anymore and that he is not connected
lblhosting.Visible = False
lblConnected.Visible = False
atk.Enabled = False
End Sub

Private Sub mgetIp_Click()
'This Sub makes sure you get the external IP

'temporary Dims
Dim a As Integer, b As Integer
Dim strURL As String, strIP As String
'we open the www.whatismyip.com URL and read it whole into strURL, this can take some time
strURL = Inet1.OpenURL("http://www.whatismyip.com/")
'we find where the part before the IP is written
a = InStr(1, strURL, "<TITLE>Your ip is ")
'we find the part after the IP
b = InStr(1, strURL, " WhatIsMyIP.com</TITLE>")
' NOTE: the strURL doesnt have the actual Text you see in the browser in it, it has the HTML
'code of the site in it! You need to watch out for that

'And the IP itself is between these two!
If Not a = 0 Then
strIP = Mid(strURL, a + 18, b - (a + 18)) 'if it confuses u, 18 is the length of the things we searched for above
'(<title>Your ip is ) and the other thing... it finds the characters between the ending of the first one and begining of 2nd
temp = MsgBox("Your IP is: " & strIP & vbCrLf & "Would you like to copy it to the clipboard ?", vbYesNo, "Your IP")
Else
MsgBox "Error: are you connected to the internet?"
End If
'We let the user know what his IP is in form of message box
'if user clicked Yes and he wants to copy the IP into the clipboard
If temp = 6 Then
    'first clear the clipboard
    Clipboard.Clear
    'then assign new clipboard text - (our IP)
    Clipboard.SetText strIP
    
End If
End Sub

Private Sub mHost_Click()
'when we host, we make label lblhost visible to let the user know that he is hosting a game
lblhosting.Visible = True
'we assign local port
Winsock1.LocalPort = port
'and tell winsock to listen at the port
Winsock1.Listen
'we set the Xpositions and Ypositions. See the mConnect sub for more documentation!
'This is a crucial part of winsock understanding... compare this sub to the mConnect sub,
'Particulary the positions im assigning!
GuyXpos(1) = 1
GuyYpos(1) = 1
GuyXpos(2) = 10
GuyYpos(2) = 10

'we move the guys to the positions
Guy(2).Move GuyXpos(2) * 32 - 32, GuyYpos(2) * 32 - 32
Guy(1).Move GuyXpos(1) * 32 - 32, GuyYpos(1) * 32 - 32
End Sub

Private Sub mPort_Click()
'We change the port
port = InputBox("Enter new port! I dont recommend lower port than 5000! default port = 5432")
End Sub

Private Sub mSay_Click()
'this is to make sure that if we are not connected and click SAY the program doesnt crash
On Error Resume Next
'what do you want to say?
strr = InputBox("What do you want to say?")
'if you didnt click Cancel or OK without typing anything
If strr <> "" Then
'we send the data with prefix 998 and again in the end we attach the | in case the packets are joined
Winsock1.SendData "998" & "/" & namez & ": " & strr & "|"
'removes first item of chatbox and adds new item to make chatroom effect without allowing focus to it
Chat.RemoveItem (0)
Chat.AddItem namez & ": " & strr
End If
End Sub

Private Sub mSkins_Click()

'if we are not connected
If lblConnected.Visible = False Then
MsgBox "NOTE: If you cloak before you connect, the other computer will still see you as uncloaked."
End If

'pretty easy, we only read in a new image
'Guy(1).Picture = GuySkin(Index).Picture

On Error Resume Next
'we send through the winsock these two > an identifier 999 and the index # of the skin
'in dataarrival sub, we check if the first data is 999 and if yes, we assign the skin with the # after
'to the guy(2) .... (see Data Arrival Sub for more documentation)
'notice the "|" in the end, we will discuss this in data arrival section
Winsock1.SendData "999" & "/" & 0 & "|"
End Sub

Private Sub Timer_Timer(Index As Integer)
'for direction: top = 1, left = 2, bottom = 3, right = 4
'the index is the # of bullet we are talking about
If direction(Index) = 1 Then
bullet(Index).Top = bullet(Index).Top - 32
ElseIf direction(Index) = 2 Then
bullet(Index).Left = bullet(Index).Left - 32
ElseIf direction(Index) = 3 Then
bullet(Index).Top = bullet(Index).Top + 32
Else
bullet(Index).Left = bullet(Index).Left + 32
End If
'if bullet goes off screen
If bullet(Index).Left < 0 Or bullet(Index).Left > 20 * 32 Or bullet(Index).Top < 0 Or bullet(Index).Top > 12 * 32 Then
'make it usable
bullet(Index).Visible = False
bullet(Index).Left = Guy(1).Left
bullet(Index).Top = Guy(1).Top

Timer(Index).Enabled = False
End If

'if u hit
If (Guy(2).Left = bullet(Index).Left And Guy(2).Top = bullet(Index).Top And bullet(Index).Visible = True) Then
'if they arent a ghost
If ghost(0).Enabled = False Then
'if you are connected
If Winsock1.State <> 0 Then
'tell em the bad news
Winsock1.SendData "002"
End If
'w00t u score
score(0) = score(0) + 1
'this is an interesting part... to make it fun for both ppl that are skilled in this and for noobs
'i decided to make it so that for every hit you get, your bullets become a bit slower
'(1/100th of a second more per position change)
For i = 0 To 9
Timer(i).Interval = Timer(i).Interval + 10
Next i
ghost(0).Enabled = True
End If
End If

End Sub

Private Sub Timer1_Timer()
'display score
Label2.Caption = "You: " & score(0)
Label3.Caption = "Them: " & score(1)
End Sub

Private Sub Timer2_Timer()
If rempic = 1 Or rempic = 2 Or rempic = 3 Or rempic = 4 Then
Guy(2).Picture = GuySkin(rempic).Picture
End If
End Sub

Private Sub Timer3_Timer()
'makes sure EVERYTHING is positioned correctly
For i = 0 To 9
bullet(i).Top = Guy(1).Top
bullet(i).Left = Guy(1).Left
bb(i).Top = Guy(2).Top
bb(i).Left = Guy(2).Left
Next i
Timer3.Enabled = False
End Sub

Private Sub Timer4_Timer()
'positions enemy bullets correctly
For i = 0 To 9
If btimer(i).Enabled = False Then
bb(i).Left = Guy(2).Left
bb(i).Top = Guy(2).Top
End If
Next i
End Sub

Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
'if the other person is requesting connection, this Sub is executed
lblConnected.Visible = True
atk.Enabled = True
'if Winsock is already in use, we close it
If Winsock1.State <> sckClosed Then
Winsock1.Close
End If

'we accept the user
Winsock1.Accept requestID
Winsock1.SendData "001"
'we let the user know that we are connected to the other person and that he joined the game
lblConnected.Visible = True
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
'now the juicy part
'This is the command to receive the data if there is any incoming
Winsock1.GetData Data, vbString, bytesTotal
'The data incoming is stored in the variable 'Data'

'whoa why the hell did i write this? better not tamper with it.
If Data = "003|" Then
Call batk
ElseIf Data = "003|003|" Then
Call batk
Call batk
ElseIf Data = "003|003|003|" Then
Call batk
Call batk
Call batk
ElseIf Data = "003|003|003|003|" Then
Call batk
Call batk
Call batk
Call batk
ElseIf Data = "003|003|003|003|003|" Then
Call batk
Call batk
Call batk
Call batk
Call batk
Else
'So we test. If the lenght of the data received is more then 6 and the packet joining has occured
If Data = "002" Then
score(1) = score(1) + 1
For i = 0 To 9
btimer(i).Interval = btimer(i).Interval + 10
Next i
Else
If Len(Data) > 6 Then

'first we use a temporary variable to store the first part. We extract the 5/15 out of the 5/15|5/15
data3 = Split(Data, "|")
'then from the first part we extract both values Data2(0) becomes the X (5) and Data2(1) becomes Y (15)
Data2 = Split(data3(0), "/")

'we print both
Label1.Caption = "Xpos = " & Data2(0) & "  .. Ypos = " & Data2(1)

'if the data2(0) is 999 and the skin change has occured
If CInt(Data2(0)) = 999 Then
'change the skin
Guy(2).Picture = GuySkin(CInt(Data2(1)))
rempic = CInt(Data2(1))
'parses the data by moving the guy
If (CInt(Data2(1))) = 4 Then
If Guy(2).Top > 0 Then
Guy(2).Top = Guy(2).Top - 32
End If
ElseIf (CInt(Data2(1))) = 3 Then
If Guy(2).Left > 0 Then
Guy(2).Left = Guy(2).Left - 32
End If
ElseIf (CInt(Data2(1))) = 2 Then
If Guy(2).Top < 32 * 11 Then
Guy(2).Top = Guy(2).Top + 32
End If
ElseIf (CInt(Data2(1))) = 1 Then
If Guy(2).Left < 19 * 32 Then
Guy(2).Left = Guy(2).Left + 32
End If
End If
'but if the prefix is 998 (for SAY message)
ElseIf CInt(Data2(0)) = 998 Then
'we do all the crap, display the label, put in text, move it appropriatly
Chat.RemoveItem (0)
Chat.AddItem Data2(1)
Else
'update the position of second player
'Guy(2).Move CInt(Data2(0)) * 32 - 32, CInt(Data2(1)) * 32 - 32

End If

ElseIf Len(Data) = 3 Then
'If Data = "001" Then
lblConnected.Visible = True
'End If
Else
'but if the packets joining hasnt occured and the Data arrived is 5/15| then
'we first split these terms by /, thus giving us 5 and 15|
Data2 = Split(Data, "/")
'I hope you notice that the Y coordinate is 15| instead of 15. So we need to cut the last character
'so now the Data2(0) is the X (5) and Data2(1) is the Y coordinate (15)
Data2(1) = Left$(Data2(1), Len(Data2(1)) - 1)

'we print both
Label1.Caption = "Xpos = " & Data2(0) & "  .. Ypos = " & Data2(1)

'if the data2(0) is 999 and the skin change has occured
If CInt(Data2(0)) = 999 Then
'change the skin
Guy(2).Picture = GuySkin(CInt(Data2(1)))
'If (CInt(Data2(1))) = 1 Then
'Guy(2).Move GuyXpos(2) * 32 - 32, GuyYpos(2) * 32 - 32
'ElseIf (CInt(Data2(1))) = 2 Then
'Guy(2).Left = Guy(2).Left + 32
'ElseIf (CInt(Data2(1))) = 3 Then
'Guy(2).Top = Guy(2).Top - 32
'Else
'Guy(2).Left = Guy(2).Left - 32
'End If
'but if the prefix is 998 (for SAY message)
ElseIf CInt(Data2(0)) = 998 Then
'we do all the crap, display the label, put in text, move it appropriatly
Chat.RemoveItem (0)
Chat.AddItem Data2(1)
Else
'update the position of second player


End If

End If
End If
End If
End Sub



