VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   4380
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6000
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  '����������
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCancel_Click() 'ȡ����ť�¼�����
    If ListBox1.ListIndex >= 0 Then
        ' ��ȡѡ���������
        Dim index As Integer
        index = openaiapi.GetHistoryCount() - ListBox1.ListIndex
        
        ' ɾ��ѡ�еļ�¼
        openaiapi.DeleteHistory index
        
        ' ���¼����б�
        ListBox1.Clear
        Dim i As Integer
        For i = openaiapi.GetHistoryCount() To 1 Step -1
            ListBox1.AddItem openaiapi.GetHistoryItem(i)
        Next
    End If
End Sub

Private Sub cmdOK_Click()
    ' ȷ����ť
    UserForm1.Tag = TextBox1.Text
    Me.Hide
End Sub

Private Sub ListBox1_Click()
    ' ����б���ʱ����ѡ���������ı���
    If ListBox1.ListIndex <> -1 Then
        TextBox1.Text = ListBox1.List(ListBox1.ListIndex)
    End If
End Sub

Private Sub UserForm_Initialize()
    ' ��ʼ������
    Me.Caption = "AI��ʽ������"
    
    ' ����������֧������
    Me.Font.Name = "����"
    TextBox1.Font.Name = "����"
    ListBox1.Font.Name = "����"
    
    ' ȷ���ı���Ϊ��
    TextBox1.Text = ""
    TextBox1.SetFocus
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Me.Tag = ""
    Me.Hide
End Sub
