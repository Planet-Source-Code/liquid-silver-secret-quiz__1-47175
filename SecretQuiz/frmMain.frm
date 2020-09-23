VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Secret Quiz"
   ClientHeight    =   1620
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3615
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1620
   ScaleWidth      =   3615
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Height          =   300
      Left            =   1350
      TabIndex        =   2
      Top             =   1200
      Width           =   930
   End
   Begin VB.Timer tmrMain 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   45
      Top             =   15
   End
   Begin VB.Label lblDescription 
      Alignment       =   2  'Center
      Caption         =   $"frmMain.frx":0E42
      Height          =   795
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   3405
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "Secret Quiz"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   630
      TabIndex        =   0
      Top             =   15
      Width           =   2310
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdStart_Click()
    Dim Data As String
    Dim Count As Integer
    
    Me.Hide
    
    Count = 0
    
StartPlace:
    
    Count = Count + 1
    
    Randomize
    Data = LoadResString(Int(Rnd * 50) + 1)
    frmBox.lblQuestion = Str(Count) + ". " + Split(Data, "|")(0)
    Continue = False
    frmBox.Show
    
    tmrMain.Enabled = True
    
    Do Until Continue = True
        DoEvents
    Loop
    
    tmrMain.Enabled = False
    frmBox.txtMain = ""
    
    If Answer <> Int(Trim(Split(Data, "|")(1))) Then
        Call MsgBox("Sorry, wrong answer, try again later.")
        End
    End If
    
    If Count < 10 Then GoTo StartPlace
    
    Call MsgBox("Well Done! The password is: 'verysecretcode'")
    
    End
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub tmrMain_Timer()
    Call MsgBox("Sorry, time is up, try again later.")
    End
End Sub
