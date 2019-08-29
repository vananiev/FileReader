VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Disks 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "��������� ������"
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
' ������� ��������� ������� FileSystemObject
Dim mfsysObject As New Scripting.FileSystemObject
' ��������� ������ Drive
Dim drvItem As Drive
Private Sub Form_Load()
' ��������� ��������� � ��������� ����
txtData.Text = "Drives    " & "Free space" & vbCrLf & vbCrLf
' �������� ����� ������� ���� �� �������� ����
MousePointer = vbHourglass
' ��������� ������ �������� ����������
'��� ��������� ������
'If drvltem.DriveType = Fixed Then
' ����� ��������� ������ ���������� ������������...
'End If
For Each drvItem In mfsysObject.Drives
' ��������� ��������� ����
DoEvents
' ���� ���� ����� � ������, ����� ��������
' ������ ���������� ����� �� ���
If drvItem.IsReady Then
txtData.Text = txtData.Text & drvItem.DriveLetter & ":\         " & Round(drvItem.FreeSpace / 10 ^ 6, 2) & " Mb" & vbCrLf
' ����� ��������, ��� ���� �� �����
Else
txtData.Text = txtData.Text & drvItem.DriveLetter & ":\         " & "Not Ready." & vbCrLf
End If
Next drvItem
' ��������������� �������� ����� ������� ����
MousePointer = vbDefault
Folder
End Sub
Sub Folder()
Dim fldObject As Folder
' ������� ���������� � ������
txtData.Text = txtData.Text & vbCrLf & "Windows folder: " & mfsysObject.GetSpecialFolder(WindowsFolder) & vbCrLf & "System folder: " & mfsysObject.GetSpecialFolder(SystemFolder) & vbCrLf & "Temporary folder: " & mfsysObject.GetSpecialFolder(TemporaryFolder) & vbCrLf & "Current folder: " & CurDir & vbCrLf
' �������� ������ ������� �����...
Set fldObject = mfsysObject.GetFolder(CurDir)
' � ������� ���-����� ���������� � ���
txtData.Text = txtData.Text & "Current directory contains: " & fldObject.Size & " bytes."
End Sub

