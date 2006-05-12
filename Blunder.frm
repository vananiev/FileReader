VERSION 5.00
Begin VB.Form Blunder 
   BorderStyle     =   0  'None
   Caption         =   "Mistake"
   ClientHeight    =   2025
   ClientLeft      =   4995
   ClientTop       =   4530
   ClientWidth     =   4725
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2025
   ScaleWidth      =   4725
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   3240
      TabIndex        =   0
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   600
      Left            =   480
      Picture         =   "Blunder.frx":0000
      Stretch         =   -1  'True
      Top             =   360
      Width           =   600
   End
   Begin VB.Label Label1 
      Caption         =   "Blunder of Explorer.exe"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   600
      Width           =   2295
   End
End
Attribute VB_Name = "Blunder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CancelButton_Click()
Unload Blunder
End Sub

Private Sub OKButton_Click()
Unload Blunder
End Sub
