VERSION 5.00
Begin VB.Form frmBox 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Secret Quiz"
   ClientHeight    =   1770
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3615
   Icon            =   "frmBox.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1770
   ScaleWidth      =   3615
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtMain 
      Height          =   1035
      Left            =   120
      TabIndex        =   0
      Top             =   645
      Width           =   3375
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   450
      Left            =   2880
      TabIndex        =   1
      Top             =   135
      Width           =   660
   End
   Begin VB.Label lblQuestion 
      Caption         =   "Enter data"
      Height          =   420
      Left            =   135
      TabIndex        =   2
      Top             =   135
      Width           =   2595
   End
End
Attribute VB_Name = "frmBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
On Error GoTo Error
    
    Answer = Int(Trim(txtMain.Text))
    Me.Hide
    Continue = True
    Exit Sub
    
Error:
    Call MsgBox("Sorry, wrong answer, try again later.")
    End
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub txtMain_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call cmdOK_Click
    End If
End Sub
