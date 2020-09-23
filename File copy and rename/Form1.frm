VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Runtime rename and copy Enjoy it !"
   ClientHeight    =   795
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5220
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   795
   ScaleWidth      =   5220
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   4695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' This Piece of code will let you copy a file from anywhere and rename it
' Note * The file will stay the same but will be only copiied to another location
' and  renamed, and will not gain any effect
' If you like this code :) i think is usefull :) u can vote me , i discovered this when i was creating a new MSN Messenger 10 Built in bots in 1 :)
' Will be available on http://www.vbdotlb.connect.to



Private Sub Command1_Click()
On Error Resume Next
Dim def As String
def = "c:\SCANDISK.LOG"

FileCopy def, "c:\windows\desktop\Renamed.log"
End Sub

