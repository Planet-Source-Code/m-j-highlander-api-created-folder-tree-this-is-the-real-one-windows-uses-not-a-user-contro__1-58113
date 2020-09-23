VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7215
   ClientLeft      =   2550
   ClientTop       =   2535
   ClientWidth     =   11460
   LinkTopic       =   "Form1"
   ScaleHeight     =   481
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   764
   Begin VB.TextBox txtPath 
      Height          =   375
      Left            =   180
      TabIndex        =   1
      Top             =   180
      Width           =   9915
   End
   Begin VB.PictureBox picBrowse 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   4995
      Left            =   240
      ScaleHeight     =   333
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   314
      TabIndex        =   0
      Top             =   660
      Width           =   4710
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub PathChange()
  
  txtPath = gs_CurrentDirectory

End Sub
Private Sub Form_Load()

Show

Set DialogContainer = picBrowse 'container for the Treeview
BrowseForFolder App.Path
    
'ChangePath App.Path

End Sub
Private Sub Form_Unload(Cancel As Integer)
  
  CloseUp

End Sub


