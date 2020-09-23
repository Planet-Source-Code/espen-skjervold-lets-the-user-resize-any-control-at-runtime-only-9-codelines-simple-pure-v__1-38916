VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4320
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7380
   LinkTopic       =   "Form1"
   ScaleHeight     =   4320
   ScaleWidth      =   7380
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   1815
      Left            =   1080
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "resize any control runtime.frx":0000
      Top             =   720
      Width           =   3135
   End
   Begin VB.Label Label1 
      Caption         =   "<-->"
      Height          =   255
      Left            =   4200
      TabIndex        =   0
      ToolTipText     =   "click and drag to resize"
      Top             =   2520
      Width           =   255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Resize any control at runtime
'Espen Skjervold, 2002
'To resize another control, change text1 with the control's name,
'and place label1 in the lower right corner




Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'resets the position of label1 if placed outside the form
If Label1.Left < 0 Or Label1.Top < 0 Then Label1.Left = Text1.Left + Text1.Width: Label1.Top = Text1.Top + Text1.Height
If Label1.Left > Form1.Width Or Label1.Top > Form1.Height Then Label1.Left = Text1.Left: Label1.Top = Text1.Top - 200
End Sub


Private Sub label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo err:

If Button = 1 Then
    Label1.Left = Label1.Left - Label1.Width / 2 + X
    Label1.Top = Label1.Top - Label1.Height / 2 + Y
    
    Text1.Width = Label1.Left - Text1.Left
    Text1.Height = Label1.Top - Text1.Top
 
End If

    
err:
End Sub

