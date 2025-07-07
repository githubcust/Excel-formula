VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   4380
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6000
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCancel_Click() '取消按钮事件处理
    If ListBox1.ListIndex >= 0 Then
        ' 获取选中项的索引
        Dim index As Integer
        index = openaiapi.GetHistoryCount() - ListBox1.ListIndex
        
        ' 删除选中的记录
        openaiapi.DeleteHistory index
        
        ' 重新加载列表
        ListBox1.Clear
        Dim i As Integer
        For i = openaiapi.GetHistoryCount() To 1 Step -1
            ListBox1.AddItem openaiapi.GetHistoryItem(i)
        Next
    End If
End Sub

Private Sub cmdOK_Click()
    ' 确定按钮
    UserForm1.Tag = TextBox1.Text
    Me.Hide
End Sub

Private Sub ListBox1_Click()
    ' 点击列表项时，将选中项填入文本框
    If ListBox1.ListIndex <> -1 Then
        TextBox1.Text = ListBox1.List(ListBox1.ListIndex)
    End If
End Sub

Private Sub UserForm_Initialize()
    ' 初始化窗体
    Me.Caption = "AI公式生成器"
    
    ' 设置字体以支持中文
    Me.Font.Name = "宋体"
    TextBox1.Font.Name = "宋体"
    ListBox1.Font.Name = "宋体"
    
    ' 确保文本框为空
    TextBox1.Text = ""
    TextBox1.SetFocus
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Me.Tag = ""
    Me.Hide
End Sub
