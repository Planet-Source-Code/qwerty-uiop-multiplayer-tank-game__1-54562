Attribute VB_Name = "Module1"
'SIMPLE MULTIPLAYER GAME
'Created by FST with help by Mephisto's walking guys game
Public outofbullets As Integer
Public rpt As Integer
Public datacount As Integer
Public rempick As Single
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
Public score(1) As Integer
Public i As Integer, r As Integer, c As Integer
Public dir As String
Public strr As String
Public direction(9) As Integer
'positions of both guys
Public GuyXpos(2) As Integer
Public GuyYpos(2) As Integer

'if lines option is selected
Public lines As Boolean

'winsock stuff
Public port As Single, hostIP As String

'winsock data arrival publics, they will be needed later on
Public Data As String
Public Data2() As String
Public data3() As String

'temporary variable for messagebox input yes/no
Public temp As Integer

'for graphics loop in painting the picture with grass
Public x As Integer, y As Integer
Public namez As String
'to make sure there are no variables that we are using that are not publicmed
Public bnum As Integer
Public bdirection(9) As Integer
Public bcount As Integer
Public rempic As Single

Public Function batk()
'the bad guy attack function... should be mostly self explanitory
bcount = 0
For i = 0 To 9
If Form1.bb(i).Left = Form1.Guy(2).Left And Form1.bb(i).Top = Form1.Guy(2).Top And Form1.bb(i).Visible = False Then
bnum = i
bcount = 1
End If
Next i
If bcount = 1 Then
Form1.Text2.Text = bnum
Form1.bb(bnum).Visible = True
If 4 = rempic Then
bdirection(bnum) = 5 - rempic
ElseIf 3 = rempic Then
bdirection(bnum) = 5 - rempic
ElseIf 2 = rempic Then
bdirection(bnum) = 5 - rempic
ElseIf 1 = rempic Then
bdirection(bnum) = 5 - rempic
End If
If rempic <> 0 Then
Form1.Guy(2).Picture = Form1.GuySkin(rempic).Picture
End If
Form1.btimer(bnum).Enabled = True
If 4 = rempic Then
bdirection(bnum) = 5 - rempic
ElseIf 3 = rempic Then
bdirection(bnum) = 5 - rempic
ElseIf 2 = rempic Then
bdirection(bnum) = 5 - rempic
ElseIf 1 = rempic Then
bdirection(bnum) = 5 - rempic
End If
End If
End Function
