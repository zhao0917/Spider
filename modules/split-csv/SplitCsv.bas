Attribute VB_Name = "SplitCsv"
Private Function get_Tristate(ByVal encoding As String)
        Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0
        
        Select Case LCase(encoding)
        'Case "utf8"                   'fso OpenTextFile不支持utf8编码
            'get_Tristate = TristateTrue
        'Case "utf-8"
            'get_Tristate = TristateTrue
        Case "unicode"
            get_Tristate = TristateTrue
        Case "ascii"
            get_Tristate = TristateFalse
        Case Else
            get_Tristate = TristateUseDefault
    End Select
    
End Function


''' 根据字段拆分
Public Function split_csv(ByVal file_in$, Optional ByVal save_dir$ = ".", _
    Optional ByVal with_title As Boolean = True, Optional ByVal sep$ = ",", _
    Optional ByVal col& = 0, Optional encoding$ = "acsii")
'参数说明：
    ' file_in:string 要处理文件的全路径
    ' save_dir : string 拆分后文件的保存目录。 可选，默认和 file_in在同一目录下
    ' with_title : bool 可选，输入文件的第一行是不是标题。默认True表示第一行是标题
    ' sep : string optional，字段之间的分隔符，默认逗号","
    ' col : long optional 以哪一列为基准进行分割，默认是0表示第一列
    ' encoding : string option 文件编码,默认"acsii"，还支持unicode编码，注意不支持
        ' utf8编码，可以用记事本的另存为转unicode编码
    
    'todo:完成后延时返回，保证数据写入（未完成）

    Const ForReading = 1, ForWriting = 2, ForAppending = 8
    
    Dim Tristate As Integer
    Dim fs As Object
    Dim dict As Object
    Dim obj_filein As Object
    Dim title As String
    Dim basedir As String

    Dim buf As StrBuff
    Dim bufs As New Collection
    
    Set dict = CreateObject("Scripting.Dictionary")
    Set fs = CreateObject("Scripting.FileSystemObject")


    If Not fs.FileExists(file_in) Then Exit Function '文件存在检查
    
    ''' 设置拆分后文件保存目录
    If save_dir = "." Then
        basedir = fs.GetParentFolderName(file_in) & "\"
    Else
        If fs.FolderExists(save_dir) Then
            basedir = save_dir
        Else
            MsgBox "你输入的保存目录: save_dir" & save_dir & " 有错。程序退出！"
            Exit Function
        End If
    End If
    
    ''' 打开输入文件
    Tristate = get_Tristate(encoding)
    Set obj_filein = fs.OpenTextFile(file_in, ForReading, False, Tristate)

    If with_title Then title = obj_filein.readline()   ''' title 提取
    
    i = 1
    Do While Not obj_filein.AtEndOfStream

        row = obj_filein.readline()
        filed = Split(row, sep)(col)
        
        ''' 以前未出现的字段，新键一个buf来保存
        If Not dict.exists(filed) Then
            Set buf = New StrBuff
            Set buf.File = fs.OpenTextFile(basedir & filed & ".csv", _
                ForWriting, True, Tristate)
            
            If with_title Then buf.AddLine title  ' 添加 title
            
            dict.Add filed, i
            bufs.Add buf
            i = i + 1
        Else
            bufs(dict(filed)).AddLine row
        End If
    Loop
    
    ''' 最后写入
    For Each buf In bufs
        buf.WritetoFile
        buf.File.Close
    Next
    
    obj_filein.Close
    Set dict = Nothing
    Set fs = Nothing

    For Each buf In bufs
        Set buf = Nothing
    Next

End Function


Public Function split_csv2(ByVal file_in$, _
    Optional ByVal with_title As Boolean = True, _
    Optional ByVal sep$ = ",", _
    Optional ByVal col& = 0)
    
'参数说明：
    ' file_in:string 要处理文件的全路径
    ' with_title : bool 可选，输入文件的第一行是不是标题。默认True表示第一行是标题
    ' sep : string optional，字段之间的分隔符，默认逗号","
    ' col : long optional 以哪一列为基准进行分割，默认是0表示第一列
    
    Dim filenum!, filein!, title$
    filein = FreeFile
    Set dict = CreateObject("Scripting.Dictionary")
    
    ''' 文件保存目录,设置为和输入文件同一个目录
    filedir = ""
    temp = Split(file_in, "\")
    For i = LBound(temp) To UBound(temp) - 1 Step 1
        filedir = filedir & "\" & temp(i)
    Next
    filedir = filedir & "\"
    filedir = Mid(filedir, 2, Len(filedir) - 1)
    
    title = ""
    Dim buf As StrBuff
    Dim bufs As New Collection

    Open file_in For Input As #filein
        If with_title Then Line Input #filein, title  ''' 文件第一行是否是标题
        
        i = 1
        Do While Not EOF(filein)
            Line Input #filein, row
            filed = Split(row, sep)(col) '获取字段名

            If Not dict.exists(filed) Then  ' 文件未打开
                Set buf = New StrBuff
                filenum = FreeFile
                
                Open filedir & filed & ".csv" For Output As #filenum
                buf.FileNumber = filenum
                
                If with_title Then buf.AddLine title
                
                dict.Add filed, i
                bufs.Add buf
                
                i = i + 1
            Else
                bufs(dict(filed)).AddLine row
            End If
        Loop
    
    For Each buf In bufs
        buf.WritetoFile
        Close #buf.FileNumber
    Next

    Close #filein
    
    Set dict = Nothing
    Set buf = Nothing
    Set bufs = Nothing
End Function

