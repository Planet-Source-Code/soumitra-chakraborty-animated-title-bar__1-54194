VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4695
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7620
   LinkTopic       =   "Form1"
   ScaleHeight     =   4695
   ScaleWidth      =   7620
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   3240
      Top             =   2160
   End
   Begin VB.Timer Timer1 
      Interval        =   80
      Left            =   3240
      Top             =   2520
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'variabiles for scrolling-used by Timer1
Dim a As String
Dim t As String
    Dim b As Integer
    Dim I, p As Integer



Private Sub Form_Load()
CheckAgain
End Sub

Private Sub Text1_Change()

End Sub

Private Sub Timer1_Timer()
'****************************************

t = Left(a, b)
Form1.Caption = t
b = b + 1
If b > I Then
    b = 0
    Timer1.Enabled = False
    Timer2.Enabled = True
    p = 0
End If
End Sub


Sub CheckAgain()
a = "Your Animated Title by Soumitra Chakraborty"
I = Len(a)
    b = 0
    p = 0
End Sub

Private Sub Timer2_Timer()
p = p + 1
If p Mod 2 = 0 Then
    Form1.Caption = a
Else
    Form1.Caption = ""
End If
If p = 20 Then
    Timer2.Enabled = False
    Timer1.Enabled = True
    
End If
End Sub
