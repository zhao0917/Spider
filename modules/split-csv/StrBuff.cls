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

' 类成员''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


' cFile is a Textstream object when open with fso OpenTextFile();
' an Integer when open with vb Open filename For... as #num
Private cFile As variant     

' 文件打开方式: 0未知; 1由fso OpenTextFile 打开; 
' 2 由Open filename For Input As #filein 方式打开 
Private cOpenType As Integer 
Private Const otFsoOpen As Integer = 1
Private Const otVbOpen As Integer = 2
Private Const otOtherOpen As Integer = 0

Private cBuff As String      ' 使用字符串作为buff，
Private cRowCount As Long   ' cBuff中行计数
Private cRowMax As Long     ' cBuff最大行数

'是不是第一次向文件写入内容
Private cFirstWrite As Boolean  


''' 初始化
Private Sub class_initialize()
  cBuff = ""
  cRowCount = 0
  cRowMax = 1024  '最大1024行，可以调整
  cOpenType = otOtherOpen
  cFirstWrite = True
End Sub


'''buff属性 r
Public Property Get Buff() As String
  Buff = cBuff
End Property


'''RowCount属性 r
Public Property Get RowCount() As Long
  RowCount = cRowCount
End Property


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
    If cRowCount = cRowMax Then WritetoFile
End Sub

Public Sub WriteUnixFile(Optional line_ending$ = vbLf)
    ''' 写入前检验
    ' unix 文件换行符是 vbLf 所以只能用 textstream.write函数写入
    If cRowCount = 0 Then Exit Sub
    

    If cOpenType = otFsoOpen Then
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
    
    Select Case cOpenType
        Case otFsoOpen
            cFile.WriteLine cBuff
        Case otVbOpen
            '' Write 函数会 把 字符串的双引号也写入，但是Print 不会
            Print #cFile, cBuff   
    End Select
    
    cBuff = ""
    cRowCount = 0
End Sub

public Function GetFile()as variant
    select case cOpenType
        case otFsoOpen
            set GetFile = cFile
        case otVbOpen
            GetFile = cFile
    end select
end function


public sub SetFile(f as variant)
    if isobject(f) then
        if not f is nothing then
            set cFile = f
            cOpenType = otFsoOpen
            exit sub
        end if
    end if

    if isNumeric(f) then
        if f = fix(f) and f > 0 and f < 512 then
            cFile = f
            cOpenType = otVbOpen
            exit sub
        end if
    end if

    msgbox "Function SetFile 参数f错误!"
    ' raise error
end sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''私有成员函数

''' terminate
Private Sub class_terminate()
    if isobject(cFile) then
        If Not cFile Is Nothing Then Set cFile = Nothing
    end if
End Sub
