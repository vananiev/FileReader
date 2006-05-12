VERSION 5.00
Begin VB.Form EnterKey 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Введите ключ"
   ClientHeight    =   1725
   ClientLeft      =   6315
   ClientTop       =   5850
   ClientWidth     =   4275
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1725
   ScaleWidth      =   4275
   Begin VB.CommandButton OkButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   3240
      TabIndex        =   3
      Top             =   1080
      Width           =   855
   End
   Begin VB.TextBox txtKey 
      Height          =   285
      Left            =   840
      TabIndex        =   2
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "Key"
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "В текстовое поле введите 8 значный ключ для зашифровки текста."
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   3975
   End
End
Attribute VB_Name = "EnterKey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim n As Integer

Private Sub OkButton_Click()
If Len(txtKey) <> 8 Then
For n = Len(txtKey) To 7
txtKey = txtKey & " "
Next n
End If
Hide
End Sub

Private Sub txtKey_Change()
If Len(txtKey) > 8 Then MsgBox "Вы должны ввести неболее 8 символов", vbOKOnly, "Mistake": txtKey = Left(txtKey, 8)
End Sub
