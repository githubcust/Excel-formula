Attribute VB_Name = "openaiapi"

' ����API�����ڶ�ȡINI�����ļ�
Private Declare PtrSafe Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" ( _
    ByVal lpApplicationName As String, _
    ByVal lpKeyName As String, _
    ByVal lpDefault As String, _
    ByVal lpReturnedString As String, _
    ByVal nSize As Long, _
    ByVal lpFileName As String) As Long

Option Explicit

' =========================
' ȫ�ֱ���
' =========================
Private HISTORY_FILE As String         ' ��ʷ��¼�ļ�·��
Private formulaHistory As Collection   ' ��ʷ��ʽ����

' =========================
' ����ڣ�����Excel��ʽ
' =========================
Sub GenerateExcelFormula()
    On Error GoTo ErrorHandler

    ' ��ʼ����ʷ�ļ�·��
    HISTORY_FILE = Environ("USERPROFILE") & "\default.ini"

    Dim userInput As String
    Dim response As String
    Dim selectedCell As Range

    ' ��ʼ����ʷ��¼
    InitHistory

    ' ����Ƿ���ѡ�еĵ�Ԫ��
    If TypeName(Selection) <> "Range" Then
        MsgBox "����ѡ��һ����Ԫ��!", vbExclamation
        Exit Sub
    End If

    Set selectedCell = Selection

    ' ��ʾ�Զ������봰�壬��ȡ�û�����
    userInput = ShowInputForm("����������Ҫ���ɵ�Excel��ʽ:")

    ' ����û��Ƿ�ȡ��������Ϊ��
    If userInput = "" Then Exit Sub

    ' ���浽��ʷ��¼
    SaveToHistory userInput

    ' ����OpenAI API���ɹ�ʽ
    response = CallOpenAI(userInput)

    ' �����Ӧ�Ƿ��������
    If Left(response, 6) = "Error:" Then
        MsgBox response, vbExclamation
        Exit Sub
    End If

    ' �����ɵĹ�ʽд��ѡ�еĵ�Ԫ��
    selectedCell.Formula = response

    ' ��ʾAI��������
    MsgBox "AI�������ݣ�" & vbCrLf & response

    Exit Sub

ErrorHandler:
    MsgBox "��������: " & Err.Description, vbCritical
End Sub

' =========================
' ��ʾ�Զ������봰�壬�����û�����
' =========================
Function ShowInputForm(promptText As String) As String
    ' ���ش���
    Load UserForm1

    ' ���ñ�ǩ�ı�
    UserForm1.Controls("Label1").Caption = promptText

    ' �����ʷ��¼�б�
    Dim i As Integer
    UserForm1.ListBox1.clear
    For i = formulaHistory.Count To 1 Step -1
        UserForm1.ListBox1.AddItem formulaHistory(i)
    Next

    ' ��ʾ���岢��ȡ����
    UserForm1.Show
    ShowInputForm = Trim(UserForm1.Tag)

    ' ж�ش���
    Unload UserForm1
End Function

' =========================
' ��ʼ����ʷ��¼����
' =========================
Sub InitHistory()
    Set formulaHistory = New Collection
    LoadHistoryFromFile
End Sub

' =========================
' ��UTF-8�����INI�ļ�������ʷ��¼
' =========================
Sub LoadHistoryFromFile()
    On Error Resume Next

    ' ����ļ��Ƿ����
    If Dir(HISTORY_FILE) = "" Then Exit Sub

    ' ʹ��ADODB.Stream��ȡUTF-8�ļ�
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")

    With stream
        .Charset = "UTF-8"
        .Type = 2 ' adTypeText
        .Open
        .LoadFromFile HISTORY_FILE

        ' ��ȡȫ�����ݲ����зָ�
        Dim content As String
        content = .ReadText
        .Close

        Dim lines As Variant
        lines = Split(content, vbCrLf)

        ' ��ӵ�����
        Dim i As Long
        For i = LBound(lines) To UBound(lines)
            If Trim(lines(i)) <> "" Then
                formulaHistory.Add lines(i)
            End If
        Next
    End With
End Sub

' =========================
' �������뵽��ʷ��¼����д���ļ�
' =========================
Sub SaveToHistory(inputText As String)
    On Error Resume Next
    ' ����Ƿ��Ѵ���
    Dim i As Integer
    For i = 1 To formulaHistory.Count
        If formulaHistory(i) = inputText Then Exit Sub
    Next

    ' ��ӵ�����
    formulaHistory.Add inputText

    ' �������50����¼
    If formulaHistory.Count > 50 Then
        For i = 1 To formulaHistory.Count - 50
            formulaHistory.Remove 1
        Next
    End If

    ' ���浽�ļ�
    SaveHistoryToFile
End Sub

' =========================
' ������ʷ��¼��UTF-8������ļ�
' =========================
Sub SaveHistoryToFile()
    On Error Resume Next

    ' ʹ��ADODB.Streamд��UTF-8�ļ�
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")

    With stream
        .Charset = "UTF-8"
        .Type = 2 ' adTypeText
        .Open

        ' д��ÿһ��
        Dim i As Long
        For i = 1 To formulaHistory.Count
            If i > 1 Then .WriteText vbCrLf
            .WriteText formulaHistory(i)
        Next

        ' ���浽�ļ�
        .SaveToFile HISTORY_FILE, 2 ' 2 = adSaveCreateOverWrite
        .Close
    End With
End Sub

' =========================
' ��ȡ��ʷ��¼Ԥ���������ʾ3����
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
' ɾ��ָ������ʷ��¼
' =========================
Public Sub DeleteHistory(ByVal index As Integer)
    On Error Resume Next
    If index > 0 And index <= formulaHistory.Count Then
        formulaHistory.Remove index
        SaveHistoryToFile
    End If
End Sub

' =========================
' ������ʷ��ʽ�б��е�����
' =========================
Public Function GetHistoryCount() As Integer
    GetHistoryCount = formulaHistory.Count
End Function

' =========================
' ��������������ʷ��ʽ��������Ч�򷵻ؿ��ַ���
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
' ��ȡ�����ļ��е�ֵ
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
' ����OpenAI API���������ɵĹ�ʽ
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

    ' ��config.ini��ȡAPI����
    url = GetConfigValue("openai", "url", "")
    apiKey = GetConfigValue("openai", "apikey", "")
    model = GetConfigValue("openai", "model", "")
    systemPrompt = "����һ��excelר�ң��ܹ�������������excel��ʽ��ע�⣺�ظ��н�����excel��ʽ����Ҫ�������κ����ݡ����磺����Ϊ���A1:A10����ظ�Ϊ=SUM(A1:A10)��"

    ' ����������
    requestBody = "{""model"":""" & model & """,""messages"":[
    requestBody = requestBody & "{""role"":""system"",""content"":""" & Replace(systemPrompt, """", "\""") & """},"
    requestBody = requestBody & "{""role"":""user"",""content"":""" & Replace(prompt, """", "\""") & """}"
    requestBody = requestBody & "],""max_tokens"":128}"

    ' ����POST����
    Set httpObj = CreateObject("WinHttp.WinHttpRequest.5.1")
    httpObj.Open "POST", url, False
    httpObj.SetRequestHeader "Content-Type", "application/json"
    httpObj.SetRequestHeader "Authorization", "Bearer " & apiKey
    httpObj.Send requestBody

    ' ��ֹ�������룬ʹ��ADODB.Stream��ȡ��Ӧ
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

    ' �������ص�JSON����ȡcontent�ֶ�
    Set json = Nothing
    On Error Resume Next
    Set json = JsonConverter.ParseJson(responseText)
    On Error GoTo 0

    If Not json Is Nothing Then
        ' ֻȡ��һ������=��ͷ��ȷ��Ϊ�Ϸ�Excel��ʽ
        Dim lines As Variant
        Dim formulaText As String
        lines = Split(Trim(json("choices")(1)("message")("content")), vbLf)
        formulaText = Trim(lines(0))
        If Left(formulaText, 1) = "=" Then
            CallOpenAI = formulaText
        Else
            CallOpenAI = "Error: AI�������ݲ�����Ч��Excel��ʽ��" & formulaText
        End If
    Else
        CallOpenAI = "Error: �޷�����OpenAI��Ӧ"
    End If
    Exit Function

ErrHandler:
    CallOpenAI = "Error: " & Err.Description
End Function
