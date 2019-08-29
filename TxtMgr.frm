VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form TxtMgr 
   Caption         =   "FineReade"
   ClientHeight    =   8070
   ClientLeft      =   2670
   ClientTop       =   2235
   ClientWidth     =   10170
   LinkTopic       =   "Form1"
   ScaleHeight     =   8070
   ScaleWidth      =   10170
   Begin VB.Timer tmrLoad 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   3120
      Top             =   0
   End
   Begin MSComctlLib.ProgressBar prbInfor 
      Height          =   270
      Left            =   8160
      TabIndex        =   7
      Top             =   7800
      Width           =   2010
      _ExtentX        =   3545
      _ExtentY        =   476
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin MSComctlLib.StatusBar stbInfor 
      Align           =   2  'Align Bottom
      DragMode        =   1  'Automatic
      Height          =   360
      Left            =   0
      TabIndex        =   6
      ToolTipText     =   "Status Bar"
      Top             =   7710
      Width           =   10170
      _ExtentX        =   17939
      _ExtentY        =   635
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin RichTextLib.RichTextBox txtFileAsk 
      Height          =   3495
      Left            =   5160
      TabIndex        =   5
      Top             =   4080
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   6165
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"TxtMgr.frx":0000
   End
   Begin RichTextLib.RichTextBox txtFileUnc 
      Height          =   3495
      Left            =   120
      TabIndex        =   4
      Top             =   4080
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   6165
      _Version        =   393217
      ScrollBars      =   2
      MousePointer    =   99
      TextRTF         =   $"TxtMgr.frx":0084
   End
   Begin RichTextLib.RichTextBox txtFile 
      Height          =   3135
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   5530
      _Version        =   393217
      ScrollBars      =   2
      TextRTF         =   $"TxtMgr.frx":0108
   End
   Begin MSComDlg.CommonDialog dlgfile 
      Left            =   2640
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      FilterIndex     =   1
      Flags           =   4100
      MaxFileSize     =   32000
   End
   Begin VB.Label Label1 
      Caption         =   "   ������ ������:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "   In ASKILL:"
      Height          =   255
      Left            =   5160
      TabIndex        =   1
      Top             =   3720
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "   In Unicode:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   3720
      Width           =   975
   End
   Begin VB.Menu File 
      Caption         =   "����"
      Begin VB.Menu New 
         Caption         =   "New"
      End
      Begin VB.Menu Razd 
         Caption         =   "-"
      End
      Begin VB.Menu Open 
         Caption         =   "�������"
      End
      Begin VB.Menu Save 
         Caption         =   "��������� ���"
      End
      Begin VB.Menu Razd1 
         Caption         =   "-"
      End
   End
   Begin VB.Menu Pravka 
      Caption         =   "������"
      Begin VB.Menu Disk 
         Caption         =   "��������� ������"
      End
      Begin VB.Menu Razd2 
         Caption         =   "-"
      End
      Begin VB.Menu Unic 
         Caption         =   "��������� � Unicode"
      End
      Begin VB.Menu Askill 
         Caption         =   "��������� � ASKILL"
      End
      Begin VB.Menu Dict 
         Caption         =   "�������������� �������"
         Begin VB.Menu VklDict 
            Caption         =   "��������"
         End
         Begin VB.Menu VklcDict 
            Caption         =   "���������"
         End
      End
      Begin VB.Menu Razd3 
         Caption         =   "-"
      End
   End
   Begin VB.Menu Help 
      Caption         =   "������"
   End
End
Attribute VB_Name = "TxtMgr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mfsysObject As New Scripting.FileSystemObject
Dim strFile As String
Dim bitFileAsk() As Byte
Dim bitFileUnc() As Byte
Dim intCount As Long
Dim blnLoad As Boolean '���� ��������
Dim blnDict As Boolean '�������������� �������
Dim strFileName As String '��� �������� �����

Private Sub Askill_Click()
txtFileAsk.Text = ""
 ' ���������� � ASKILL
    bitFileAsk = StrConv(txtFile.Text, vbFromUnicode)
    For intCount = LBound(bitFileAsk) To UBound(bitFileAsk)
    txtFileAsk.Text = txtFileAsk.Text & bitFileAsk(intCount) & " "
    Next intCount
End Sub

Private Sub Disk_Click()
Disks.Show vbModal
End Sub
Private Sub Form_Resize()
txtFile.Height = TxtMgr.Height / 2 - 1100
txtFile.Width = TxtMgr.Width - 400
txtFileAsk.Height = txtFile.Height
txtFileUnc.Height = txtFile.Height
txtFileAsk.Width = TxtMgr.Width / 2 - 300
txtFileUnc.Width = txtFileAsk.Width
txtFileUnc.Top = txtFile.Height + 1000
txtFileAsk.Top = txtFile.Height + 1000
txtFileAsk.Left = txtFileUnc.Width + 300
Label1.Left = 100
Label2.Left = 100
Label3.Left = txtFileUnc.Width + 300
Label1.Top = 120
Label2.Top = txtFile.Height + 620
Label3.Top = txtFile.Height + 620
prbInfor.Left = TxtMgr.Width - 2100
prbInfor.Top = TxtMgr.Height - 1110
End Sub

Private Sub Open_Click()
txtFile = ""
txtFileAsk = ""
txtFileUnc = ""
' ��������� ������ ���������� ������
Dim tstrOpen As TextStream
' ��������� ����������� ���������� ����
dlgfile.Filter = "Text files(*.txt)|*.txt|Files(*)|*.*|Binary files(*.*)|*.*|Cipher files(*.cph)|*.cph|All files(*.*)|*.*"
dlgfile.DialogTitle = "�������"
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
'����������������� ��������
blnLoad = True
stbInfor.SimpleText = "��������"
tmrLoad.Enabled = True
' ��������� ��������� �����
Set tstrOpen = mfsysObject.OpenTextFile(strFileName, ForReading, True)
DoEvents
' ���������, �� ������� �� ����� � ������� �����
If tstrOpen.AtEndOfStream Then
   ' ������� ��������� ����, �� ������ �� ���������,
   ' ��� ��� � ����� ������� �����
   strFile = ""
Else
   Select Case dlgfile.FilterIndex
   Case 4
       '��������� ������������� ����
       Dim strKey As String
       strFile = tstrOpen.ReadAll
       '����
       strKey = Left(strFile, 8)
       '������
       strFile = Right(strFile, Len(strFile) - 8)
       Dim cipherTest As New Cipher
       cipherTest.KeyString = strKey
       cipherTest.Text = strFile
       cipherTest.DoXor
       strFile = Left(cipherTest.Text, Len(cipherTest.Text) - 2)
       txtFile.Text = strFile
    Case 3
       Open strFileName For Binary As #1
       Get #1, , strFile
       Close #1
       txtFile.Text = strFile
    Case 1
         ' ��������� � ���������� ��������� �����
        strFile = tstrOpen.ReadAll
        txtFile.Text = strFile
    Case 2
        Open strFileName For Input As #1
        Line Input #1, strFile
        Close #1
        txtFile.Text = strFile
    Case 5
         ' ��������� � ���������� ��������� �����
        strFile = tstrOpen.ReadAll
        txtFile.Text = strFile
    End Select
End If
' ��������� �����
tstrOpen.Close
'����������������� ��������
blnLoad = False
stbInfor.SimpleText = ""
tmrLoad.Enabled = False
prbInfor.Value = 0
End Sub
Private Sub New_Click()
txtFileAsk.Text = ""
txtFileUnc.Text = ""
txtFile.Text = ""
End Sub

Private Sub Save_Click()
' ��������� ������ ���������� ������
Dim tstrSave As TextStream
' ��������� ����������� ���������� ����
dlgfile.Filter = "Text files(*.txt)|*.txt|Files(*)|*.*|Binary files(*.*)|*.*|Cipher files(*.cph)|*.cph|All files(*.*)|*.*"
dlgfile.DialogTitle = "���������"
dlgfile.ShowSave
strFileName = dlgfile.FileName
' ���������, ���� �� ������� ��� �����
If strFileName = "" Then MsgBox "�������� ��� �����", vbOKOnly, "Information": Exit Sub
' ���������, ��� �� ��� ������ �����
If mfsysObject.FileExists(strFileName) Then
Dim intOverwrite As Integer
' ����������� ������������� �� ���������� ������������� �����
intOverwrite = MsgBox("File already exists. " & "Overwrite it?", vbYesNo)
'���� ������������ �������� No, ������� �� ���� ���������
If intOverwrite = vbNo Then
Exit Sub
End If
End If
'����� ��������� ������
Select Case dlgfile.FilterIndex
Case 5
   ' ��������� ��������� �����...
   Set tstrSave = mfsysObject.OpenTextFile(strFileName, ForWriting, True)
   ' ���������...
   tstrSave.Write txtFile.Text
   ' � ���������
   tstrSave.Close
Case 1
   ' ��������� ��������� �����...
   Set tstrSave = mfsysObject.OpenTextFile(strFileName, ForWriting, True)
   ' ���������...
   tstrSave.Write txtFile.Text
   ' � ���������
   tstrSave.Close
Case 2
   Open strFileName For Output As #1
   Print #1, txtFile.Text
   Close #1
Case 3
   '��������� �������� ����
   Open strFileName For Binary As #1
   Put #1, , txtFile.Text
   Close #1
Case 4
   Dim cipherTest As New Cipher
   '������ �����
   EnterKey.Show vbModal
   '��������
   cipherTest.KeyString = EnterKey.txtKey
   cipherTest.Text = txtFile.Text
   cipherTest.DoXor
   '������ ������ 8 �������� - ����
   Open strFileName For Output As #1
   Print #1, EnterKey.txtKey & cipherTest.Text
   Close #1
End Select
End Sub

Private Sub tmrLoad_Timer()
DoEvents
prbInfor.Value = Len(strFile) / FileLen(strFileName) * 100
End Sub

Private Sub txtFile_Change()
If blnDict Then
     txtFileUnc.Text = ""
     txtFileAsk.Text = ""
 ' ���������� � ASKILL
    bitFileAsk = StrConv(txtFile.Text, vbFromUnicode)
    For intCount = LBound(bitFileAsk) To UBound(bitFileAsk)
    txtFileAsk.Text = txtFileAsk.Text & bitFileAsk(intCount) & " "
    Next intCount
    '���������� � Unicode
    bitFileUnc = txtFile.Text
    For intCount = LBound(bitFileUnc) To UBound(bitFileUnc)
    txtFileUnc.Text = txtFileUnc.Text & bitFileUnc(intCount) & " "
    Next intCount
Else
    Exit Sub
    End If
End Sub

Private Sub Unic_Click()
    txtFileUnc.Text = ""
 '���������� � Unicode
    bitFileUnc = txtFile.Text
    For intCount = LBound(bitFileUnc) To UBound(bitFileUnc)
    txtFileUnc.Text = txtFileUnc.Text & bitFileUnc(intCount) & " "
    Next intCount
End Sub

Private Sub VklcDict_Click()
blnDict = False
End Sub

Private Sub VklDict_Click()
blnDict = True
End Sub
