VERSION 5.00
Begin VB.Form DskMgr 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   10680
   ClientLeft      =   -60
   ClientTop       =   540
   ClientWidth     =   15270
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10680
   ScaleWidth      =   15270
   Begin VB.CommandButton Command1 
      Caption         =   "Dicks"
      Default         =   -1  'True
      Height          =   495
      Left            =   14040
      TabIndex        =   2
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label txtdata4 
      Height          =   3495
      Left            =   5280
      TabIndex        =   7
      Top             =   6360
      Width           =   9975
   End
   Begin VB.Label Label3 
      Caption         =   "In Unicode:"
      Height          =   255
      Left            =   5280
      TabIndex        =   6
      Top             =   6000
      Width           =   9975
   End
   Begin VB.Label Label2 
      Caption         =   "In ASKILL:"
      Height          =   255
      Left            =   4920
      TabIndex        =   5
      Top             =   3120
      Width           =   4815
   End
   Begin VB.Label Label1 
      Caption         =   "������ ������:"
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   3120
      Width           =   1815
   End
   Begin VB.Label txtdata3 
      Height          =   2415
      Left            =   5280
      TabIndex        =   3
      Top             =   3480
      Width           =   9975
   End
   Begin VB.Label txtData2 
      Height          =   6855
      Left            =   120
      TabIndex        =   1
      Top             =   3600
      Width           =   5055
   End
   Begin VB.Label txtData 
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   7935
   End
   Begin VB.Menu Fail 
      Caption         =   "����"
   End
   Begin VB.Menu Prav 
      Caption         =   "������"
      Begin VB.Menu Disks 
         Caption         =   "�������� ��������� ����������"
      End
   End
   Begin VB.Menu Help 
      Caption         =   "������"
   End
End
Attribute VB_Name = "Dskmgr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' ������� ��������� ������� FileSystemObject
Dim mfsysObject As New Scripting.FileSystemObject

Private Sub Command1_Click()
' ��������� ������ Drive
Dim drvltem As Drive
' ��������� ��������� � ��������� ����
txtData = "Drive" & "    " & "Free space" & vbCrLf & vbCrLf
' �������� ����� ������� ���� �� �������� ����
MousePointer = vbHourglass
' ��������� ������ �������� ����������
'��� ��������� ������
'If drvltem.DriveType = Fixed Then
' ����� ��������� ������ ���������� ������������...
'End If
For Each drvltem In mfsysObject.Drives
' ��������� ��������� ����
DoEvents
' ���� ���� ����� � ������, ����� ��������
' ������ ���������� ����� �� ���
If drvltem.IsReady Then
txtData = txtData & drvltem.DriveLetter & ":\       " & Round(drvltem.FreeSpace / 10 ^ 6, 2) & " Mb" & vbCrLf
' ����� ��������, ��� ���� �� �����
Else
txtData = txtData & drvltem.DriveLetter & ":\       " & "Not Ready." & vbCrLf
End If
Next drvltem
' ��������������� �������� ����� ������� ����
MousePointer = vbDefault
folder
End Sub

Sub folder()
Dim fldObject As folder
' ������� ���������� � ������
txtData = txtData & vbCrLf & "Windows folder: " & mfsysObject.GetSpecialFolder(WindowsFolder) & vbCrLf & "System folder: " & mfsysObject.GetSpecialFolder(SystemFolder) & vbCrLf & "Temporary folder: " & mfsysObject.GetSpecialFolder(TemporaryFolder) & vbCrLf & "Current folder: " & CurDir & vbCrLf
' �������� ������ ������� �����...
Set fldObject = mfsysObject.GetFolder(CurDir)
' � ������� ���-����� ���������� � ���
txtData = txtData & "Current directory contains: " & fldObject.Size & " bytes."
End Sub
Sub cmdOpenFile_Click()
txtData2 = ""
txtdata3 = ""
txtdata4 = ""
' ��������� ������ ���������� ������
Dim tstrOpen As TextStream
Dim strFileName As String
' ��������� ����������� ���������� ����
dlgfile.ShowOpen
strFileName = dlgfile.FileName
' ���������, ���� �� ������� ��� �����
If strFileName = "" Then Exit Sub
' ���������, ��� �� ��� ������ �����
If Not mfsysObject.FileExists(strFileName) Then
Dim intCreate As Integer
intCreate = MsgBox("File not found. Create it?", vbYesNo)
If intCreate = vbNo Then
Exit Sub
End If
End If
' ��������� ��������� �����
Set tstrOpen = mfsysObject.OpenTextFile(strFileName, ForReading, True)
' ���������, �� ������� �� ����� � ������� �����
If tstrOpen.AtEndOfStream Then
' ������� ��������� ����, �� ������ �� ���������,
' ��� ��� � ����� ������� �����
txtData2 = ""
Else
' ��������� � ���������� ��������� �����
txtData2 = tstrOpen.ReadAll
End If
Dim bitB() As Byte
Dim intn As Integer
bitB() = txtData2
For intn = LBound(bitB) To UBound(bitB)
txtdata4 = txtdata4 & bitB(intn) & " "
Next intn
Print
bitB() = StrConv(txtData2, vbFromUnicode)
For intn = LBound(bitB) To UBound(bitB)
txtdata3 = txtdata3 & bitB(intn) & " "
Next intn


' ��������� �����
tstrOpen.Close
End Sub

