VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StrBuff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' 类成员
Private cFile As TextStream      ' fso TextStream
Private cFileNumber As Integer  ' 供Open filename For Input As #filein，保存文件号(filenumber)
Private cBuff As String      ' 使用字符串作为buff，
Private cRowCount As Long   ' cBuff中行计数
Private cRowMax As Long     ' cBuff最大行数
Private cFileOpenType As Integer ' 文件打开方式: 0未知; 1由fso OpenTextFile 打开; 2 由Open filename For Input As #filein 方式打开
Private cFirstWrite As Boolean  '是不是第一次向文件写入内容,供fso 打开方式使用，防止在文件最前加入空白行

Private Const fsoOpenTextFile As Integer = 1
Private Const vbOpen As Integer = 2
Private Const otherOpen As Integer = 0

''' 初始化
Private Sub class_initialize()
  cBuff = ""
  cRowCount = 0
  cFileNumber = 0
  cRowMax = 1024  '最大1024行，可以调整
  cFileOpenType = otherOpen
  cFirstWrite = True
End Sub

''' file 属性 r w
Public Property Get File() As Object
    Set File = cFile
End Property

Public Property Set File(f As Object)
  Set cFile = f
End Property

'''FileNumber属性  rw
Public Property Get FileNumber() As Integer
  FileNumber = cFileNumber
End Property

Public Property Let FileNumber(ByVal fn As Integer)
    cFileNumber = fn
End Property

'''buff属性 r
Public Property Get Buff() As String
  Buff = cBuff
End Property

'Public Property Let Buff(buf As String)
'    cBuff = buf
'End Property

'''RowCount属性 r
Public Property Get RowCount() As Long
  RowCount = cRowCount
End Property

'Public Property Let RowCount(byval n As Long)
'  cRowCount = n
'End Property

''' RowMax 属性 r w
Public Property Get RowMax() As Long
  RowMax = cRowMax
End Property

Public Property Let RowMax(ByVal n As Long)
  cRowMax = n
End Property


''''''''''''''''''''''''''''''''''''
''' 公有成员函数

Public Sub AddLine(ByVal line$, Optional ByVal line_ending$ = vbCrLf)
''' 向 buff中添加一行
    '' 参数 line: string 要添加的字符串
    '' 参数lineending : string 换行符
    
    If cBuff = "" Then cBuff = line Else cBuff = cBuff & line_ending & line
    
    cRowCount = cRowCount + 1
    If cRowCount = cRowMax Then WritetoFile line_ending
End Sub

Public Sub WriteUnixFile(Optional line_ending$ = vbLf)
    ''' 写入前检验
    ' unix 文件换行符是 vbLf 所以只能用 textstream.write函数写入
    If cRowCount = 0 Then Exit Sub
    
    Dim tp As Integer
    tp = GetOpenType()     ' 缺少错误处理机制，tp为其他类型要提醒报错

    If tp = fsoOpenTextFile Then
        If cFirstWrite Then
            cFile.Write cBuff
            cFirstWrite = False
        Else
            cFile.Write (line_ending & cBuff)
        End If
    End If

    cBuff = ""
    cRowCount = 0
End Sub

Public Sub WritetoFile()
''' 将buff中内容写入文件
''' fso.write 不会写入换行符，但是 write #1, str 会自动写入换行vbCrLf
    
    ''' 写入前检验
    If cRowCount = 0 Then Exit Sub
    
    Select Case GetOpenType
        Case fsoOpenTextFile
            cFile.WriteLine cBuff
        Case vbOpen
            Print #cFileNumber, cBuff   '' Write 函数会 把 字符串的双引号也写入，但是Print 不会
    End Select
    
    cBuff = ""
    cRowCount = 0
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''私有成员函数
Private Function GetOpenType() As Integer
''' 判断文件打开方式类型 fso.opentextfile 为1，vb 的 Open 为 2 ，其他为0 （出错，要进行错误处理）
    If cFileOpenType = 0 Then
        If Not cFile Is Nothing Then
            cFileOpenType = fsoOpenTextFile
        ElseIf cFileNumber > 0 And cFileNumber < 512 Then
            cFileOpenType = vbOpen
        Else
            cFileOpenType = otherOpen
        End If
    End If
    GetOpenType = cFileOpenType
End Function


''' terminate
Private Sub class_terminate()
    If Not cFile Is Nothing Then Set cFile = Nothing
End Sub
