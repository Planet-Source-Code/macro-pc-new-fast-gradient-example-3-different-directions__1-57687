VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4350
   ClientLeft      =   60
   ClientTop       =   525
   ClientWidth     =   6375
   LinkTopic       =   "Form1"
   ScaleHeight     =   4350
   ScaleWidth      =   6375
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   2400
      TabIndex        =   0
      Top             =   1680
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Call CreateGradient(Me, 2, RGB(0, 0, 300), RGB(0, 300, 0))
End Sub
