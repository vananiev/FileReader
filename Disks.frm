VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Disks 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "—осто€ние дисков"
   ClientHeight    =   3960
   ClientLeft      =   2760
   ClientTop       =   2325
   ClientWidth     =   7755
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   7755
   Begin RichTextLib.RichTextBox txtData 
      Height          =   3975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   7011
      _Version        =   393217
      ScrollBars      =   2
      TextRTF         =   $"Disks.frx":0000
   End
End
Attribute VB_Name = "Disks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' создаем экземпл€р объекта FileSystemObject
Dim mfsysObject As New Scripting.FileSystemObject
' объ€вл€ем объект Drive
Dim drvItem As Drive
Private Sub Form_Load()
' добавл€ем заголовки в текстовое поле
txtData.Text = "Drives    " & "Free space" & vbCrLf & vbCrLf
' измен€ем форму курсора мыши на песочные часы
MousePointer = vbHourglass
' провер€ем каждое дисковое устройство
'ƒл€ системных дисков
'If drvltem.DriveType = Fixed Then
' здесь провер€ем размер свободного пространства...
'End If
For Each drvItem In mfsysObject.Drives
' обновл€ем текстовое поле
DoEvents
' если диск готов к работе, можно вы€снить
' размер свободного места на нем
If drvItem.IsReady Then
txtData.Text = txtData.Text & drvItem.DriveLetter & ":\         " & Round(drvItem.FreeSpace / 10 ^ 6, 2) & " Mb" & vbCrLf
' иначе сообщаем, что диск не готов
Else
txtData.Text = txtData.Text & drvItem.DriveLetter & ":\         " & "Not Ready." & vbCrLf
End If
Next drvItem
' восстанавливаем исходную форму курсора мыши
MousePointer = vbDefault
Folder
End Sub
Sub Folder()
Dim fldObject As Folder
' выводим информацию о папках
txtData.Text = txtData.Text & vbCrLf & "Windows folder: " & mfsysObject.GetSpecialFolder(WindowsFolder) & vbCrLf & "System folder: " & mfsysObject.GetSpecialFolder(SystemFolder) & vbCrLf & "Temporary folder: " & mfsysObject.GetSpecialFolder(TemporaryFolder) & vbCrLf & "Current folder: " & CurDir & vbCrLf
' получаем объект текущей папки...
Set fldObject = mfsysObject.GetFolder(CurDir)
' и выводим кое-какую информацию о нем
txtData.Text = txtData.Text & "Current directory contains: " & fldObject.Size & " bytes."
End Sub

