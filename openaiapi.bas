Attribute VB_Name = "openaiapi"

' 声明API：用于读取INI配置文件
Private Declare PtrSafe Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" ( _
    ByVal lpApplicationName As String, _
    ByVal lpKeyName As String, _
    ByVal lpDefault As String, _
    ByVal lpReturnedString As String, _
    ByVal nSize As Long, _
    ByVal lpFileName As String) As Long

Option Explicit

' =========================
' 全局变量
' =========================
Private HISTORY_FILE As String         ' 历史记录文件路径
Private formulaHistory As Collection   ' 历史公式集合

' =========================
' 主入口：生成Excel公式
' =========================
Sub GenerateExcelFormula()
    On Error GoTo ErrorHandler

    ' 初始化历史文件路径
    HISTORY_FILE = Environ("USERPROFILE") & "\default.ini"

    Dim userInput As String
    Dim response As String
    Dim selectedCell As Range

    ' 初始化历史记录
    InitHistory

    ' 检查是否有选中的单元格
    If TypeName(Selection) <> "Range" Then
        MsgBox "请先选择一个单元格!", vbExclamation
        Exit Sub
    End If

    Set selectedCell = Selection

    ' 显示自定义输入窗体，获取用户输入
    userInput = ShowInputForm("请描述您想要生成的Excel公式:")

    ' 检查用户是否取消或输入为空
    If userInput = "" Then Exit Sub

    ' 保存到历史记录
    SaveToHistory userInput

    ' 调用OpenAI API生成公式
    response = CallOpenAI(userInput)

    ' 检查响应是否包含错误
    If Left(response, 6) = "Error:" Then
        MsgBox response, vbExclamation
        Exit Sub
    End If

    ' 将生成的公式写入选中的单元格
    selectedCell.Formula = response

    ' 显示AI返回内容
    MsgBox "AI返回内容：" & vbCrLf & response

    Exit Sub

ErrorHandler:
    MsgBox "发生错误: " & Err.Description, vbCritical
End Sub

' =========================
' 显示自定义输入窗体，返回用户输入
' =========================
Function ShowInputForm(promptText As String) As String
    ' 加载窗体
    Load UserForm1

    ' 设置标签文本
    UserForm1.Controls("Label1").Caption = promptText

    ' 填充历史记录列表
    Dim i As Integer
    UserForm1.ListBox1.clear
    For i = formulaHistory.Count To 1 Step -1
        UserForm1.ListBox1.AddItem formulaHistory(i)
    Next

    ' 显示窗体并获取输入
    UserForm1.Show
    ShowInputForm = Trim(UserForm1.Tag)

    ' 卸载窗体
    Unload UserForm1
End Function

' =========================
' 初始化历史记录集合
' =========================
Sub InitHistory()
    Set formulaHistory = New Collection
    LoadHistoryFromFile
End Sub

' =========================
' 从UTF-8编码的INI文件加载历史记录
' =========================
Sub LoadHistoryFromFile()
    On Error Resume Next

    ' 检查文件是否存在
    If Dir(HISTORY_FILE) = "" Then Exit Sub

    ' 使用ADODB.Stream读取UTF-8文件
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")

    With stream
        .Charset = "UTF-8"
        .Type = 2 ' adTypeText
        .Open
        .LoadFromFile HISTORY_FILE

        ' 读取全部内容并按行分割
        Dim content As String
        content = .ReadText
        .Close

        Dim lines As Variant
        lines = Split(content, vbCrLf)

        ' 添加到集合
        Dim i As Long
        For i = LBound(lines) To UBound(lines)
            If Trim(lines(i)) <> "" Then
                formulaHistory.Add lines(i)
            End If
        Next
    End With
End Sub

' =========================
' 保存输入到历史记录，并写入文件
' =========================
Sub SaveToHistory(inputText As String)
    On Error Resume Next
    ' 检查是否已存在
    Dim i As Integer
    For i = 1 To formulaHistory.Count
        If formulaHistory(i) = inputText Then Exit Sub
    Next

    ' 添加到集合
    formulaHistory.Add inputText

    ' 保持最多50条记录
    If formulaHistory.Count > 50 Then
        For i = 1 To formulaHistory.Count - 50
            formulaHistory.Remove 1
        Next
    End If

    ' 保存到文件
    SaveHistoryToFile
End Sub

' =========================
' 保存历史记录到UTF-8编码的文件
' =========================
Sub SaveHistoryToFile()
    On Error Resume Next

    ' 使用ADODB.Stream写入UTF-8文件
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")

    With stream
        .Charset = "UTF-8"
        .Type = 2 ' adTypeText
        .Open

        ' 写入每一行
        Dim i As Long
        For i = 1 To formulaHistory.Count
            If i > 1 Then .WriteText vbCrLf
            .WriteText formulaHistory(i)
        Next

        ' 保存到文件
        .SaveToFile HISTORY_FILE, 2 ' 2 = adSaveCreateOverWrite
        .Close
    End With
End Sub

' =========================
' 获取历史记录预览（最多显示3条）
' =========================
Function GetHistoryPreview() As String
    Dim result As String
    Dim i As Integer
    Dim showCount As Integer

    showCount = formulaHistory.Count
    If showCount > 3 Then showCount = 3

    For i = formulaHistory.Count To formulaHistory.Count - showCount + 1 Step -1
        result = result & "? " & Left(formulaHistory(i), 30) & vbNewLine
    Next

    GetHistoryPreview = result
End Function

' =========================
' 删除指定的历史记录
' =========================
Public Sub DeleteHistory(ByVal index As Integer)
    On Error Resume Next
    If index > 0 And index <= formulaHistory.Count Then
        formulaHistory.Remove index
        SaveHistoryToFile
    End If
End Sub

' =========================
' 返回历史公式列表中的项数
' =========================
Public Function GetHistoryCount() As Integer
    GetHistoryCount = formulaHistory.Count
End Function

' =========================
' 根据索引返回历史公式，索引无效则返回空字符串
' =========================
Public Function GetHistoryItem(ByVal index As Integer) As String
    On Error Resume Next
    If index > 0 And index <= formulaHistory.Count Then
        GetHistoryItem = formulaHistory(index)
    Else
        GetHistoryItem = ""
    End If
End Function

' =========================
' 读取配置文件中的值
' =========================
Private Function GetConfigValue(section As String, key As String, defaultValue As String) As String
    Dim iniPath As String
    iniPath = Environ("USERPROFILE") & "\config.ini"
    Dim ret As String * 1024
    Dim length As Long
    length = GetPrivateProfileString(section, key, defaultValue, ret, 1024, iniPath)
    GetConfigValue = Left(ret, length)
End Function

' =========================
' 调用OpenAI API，返回生成的公式
' =========================
Private Function CallOpenAI(prompt As String) As String
    On Error GoTo ErrHandler

    Dim httpObj As Object
    Dim url As String
    Dim apiKey As String
    Dim model As String
    Dim requestBody As String
    Dim responseText As String
    Dim systemPrompt As String
    Dim json As Object

    ' 从config.ini读取API配置
    url = GetConfigValue("openai", "url", "")
    apiKey = GetConfigValue("openai", "apikey", "")
    model = GetConfigValue("openai", "model", "")
    systemPrompt = "你是一个excel专家，能够根据描述生成excel公式，注意：回复中仅包含excel公式，不要有其他任何内容。例如：描述为求和A1:A10，则回复为=SUM(A1:A10)。"

    ' 构造请求体
    requestBody = "{""model"":""" & model & """,""messages"":[
    requestBody = requestBody & "{""role"":""system"",""content"":""" & Replace(systemPrompt, """", "\""") & """},"
    requestBody = requestBody & "{""role"":""user"",""content"":""" & Replace(prompt, """", "\""") & """}"
    requestBody = requestBody & "],""max_tokens"":128}"

    ' 发送POST请求
    Set httpObj = CreateObject("WinHttp.WinHttpRequest.5.1")
    httpObj.Open "POST", url, False
    httpObj.SetRequestHeader "Content-Type", "application/json"
    httpObj.SetRequestHeader "Authorization", "Bearer " & apiKey
    httpObj.Send requestBody

    ' 防止中文乱码，使用ADODB.Stream读取响应
    Dim bytes() As Byte
    bytes = httpObj.responseBody

    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    With stream
        .Type = 1 ' adTypeBinary
        .Open
        .Write bytes
        .Position = 0
        .Type = 2 ' adTypeText
        .Charset = "utf-8"
        responseText = .ReadText
        .Close
    End With

    ' 解析返回的JSON，提取content字段
    Set json = Nothing
    On Error Resume Next
    Set json = JsonConverter.ParseJson(responseText)
    On Error GoTo 0

    If Not json Is Nothing Then
        ' 只取第一行且以=开头，确保为合法Excel公式
        Dim lines As Variant
        Dim formulaText As String
        lines = Split(Trim(json("choices")(1)("message")("content")), vbLf)
        formulaText = Trim(lines(0))
        If Left(formulaText, 1) = "=" Then
            CallOpenAI = formulaText
        Else
            CallOpenAI = "Error: AI返回内容不是有效的Excel公式：" & formulaText
        End If
    Else
        CallOpenAI = "Error: 无法解析OpenAI响应"
    End If
    Exit Function

ErrHandler:
    CallOpenAI = "Error: " & Err.Description
End Function
