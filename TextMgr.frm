VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form TextMgr 
   Caption         =   "Form1"
   ClientHeight    =   7980
   ClientLeft      =   2175
   ClientTop       =   1635
   ClientWidth     =   9885
   LinkTopic       =   "Form1"
   ScaleHeight     =   7980
   ScaleWidth      =   9885
   Begin VB.TextBox txtFileUnc 
      Height          =   2895
      Left            =   0
      TabIndex        =   7
      Top             =   4080
      Width           =   4935
   End
   Begin VB.TextBox txtFileAsk 
      Height          =   2895
      Left            =   5040
      TabIndex        =   6
      Top             =   4080
      Width           =   4695
   End
   Begin VB.VScrollBar scrHeightFile 
      Height          =   3015
      LargeChange     =   500
      Left            =   9480
      Max             =   32000
      SmallChange     =   115
      TabIndex        =   5
      Top             =   480
      Width           =   255
   End
   Begin VB.TextBox txtFile 
      Height          =   3000
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   9400
   End
   Begin VB.CommandButton cmdOpenFile 
      Caption         =   "Open  File"
      Height          =   495
      Left            =   8640
      TabIndex        =   0
      Top             =   7080
      Width           =   975
   End
   Begin MSComDlg.CommonDialog dlgfile 
      Left            =   9360
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label3 
      Caption         =   "In Unicode:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   3720
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "In ASKILL:"
      Height          =   255
      Left            =   5040
      TabIndex        =   2
      Top             =   3720
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "‘ормат текста:"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "TextMgr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mfsysObject As New Scripting.FileSystemObject
Dim strFile As String
Dim bitFileAsk() As Byte
Dim bitFileUnc() As Byte
Dim intCount As Integer
Sub cmdOpenFile_Click()
txtFile = ""
txtFileAsk = ""
txtFileUnc = ""
' объ€вл€ем объект текстового потока
Dim tstrOpen As TextStream
Dim strFileName As String
' открываем стандартное диалоговое окно
dlgfile.ShowOpen
strFileName = dlgfile.FileName
' провер€ем, было ли указано им€ файла
If strFileName = "" Then Exit Sub
' провер€ем, нет ли уже такого файла
If Not mfsysObject.FileExists(strFileName) Then
Dim intCreate As Integer
intCreate = MsgBox("File not found. Create it?", vbYesNo)
If intCreate = vbNo Then
Exit Sub
End If
End If
' открываем текстовый поток
Set tstrOpen = mfsysObject.OpenTextFile(strFileName, ForReading, True)
' провер€ем, не нулева€ ли длина у данного файла
If tstrOpen.AtEndOfStream Then
' очищаем текстовое поле, но ничего не считываем,
' так как у файла нулева€ длина
strFile = ""
Else
' считываем и отображаем текстовый поток
strFile = tstrOpen.ReadAll
End If
txtFile = strFile
' ќтображаем в ASKILL
bitFileAsk = StrConv(strFile, vbFromUnicode)
For intCount = LBound(bitFileAsk) To UBound(bitFileAsk)
txtFileAsk = txtFileAsk & bitFileAsk(intCount) & " "
Next intCount
'ќтображаем в Unicode
bitFileUnc = strFile
For intCount = LBound(bitFileUnc) To UBound(bitFileUnc)
txtFileUnc = txtFileUnc & bitFileUnc(intCount) & " "
Next intCount
' закрываем поток
tstrOpen.Close
End Sub

Private Sub scrHeightFile_Change()
If (Len(strFile) - scrHeightFile.Value) > 0 Then txtFile = Right(strFile, (Len(strFile) - scrHeightFile.Value))
End Sub
