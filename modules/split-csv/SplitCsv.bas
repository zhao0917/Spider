Attribute VB_Name = "SplitCsv"
Private Function get_Tristate(ByVal encoding As String)
        Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0
        
        Select Case LCase(encoding)
        'Case "utf8"                   'fso OpenTextFile��֧��utf8����
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


''' �����ֶβ��
Public Function split_csv(ByVal file_in$, Optional ByVal save_dir$ = ".", _
    Optional ByVal with_title As Boolean = True, Optional ByVal sep$ = ",", _
    Optional ByVal col& = 0, Optional encoding$ = "acsii")
'����˵����
    ' file_in:string Ҫ�����ļ���ȫ·��
    ' save_dir : string ��ֺ��ļ��ı���Ŀ¼�� ��ѡ��Ĭ�Ϻ� file_in��ͬһĿ¼��
    ' with_title : bool ��ѡ�������ļ��ĵ�һ���ǲ��Ǳ��⡣Ĭ��True��ʾ��һ���Ǳ���
    ' sep : string optional���ֶ�֮��ķָ�����Ĭ�϶���","
    ' col : long optional ����һ��Ϊ��׼���зָĬ����0��ʾ��һ��
    ' encoding : string option �ļ�����,Ĭ��"acsii"����֧��unicode���룬ע�ⲻ֧��
        ' utf8���룬�����ü��±������Ϊתunicode����
    
    'todo:��ɺ���ʱ���أ���֤����д�루δ��ɣ�

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


    If Not fs.FileExists(file_in) Then Exit Function '�ļ����ڼ��
    
    ''' ���ò�ֺ��ļ�����Ŀ¼
    If save_dir = "." Then
        basedir = fs.GetParentFolderName(file_in) & "\"
    Else
        If fs.FolderExists(save_dir) Then
            basedir = save_dir
        Else
            MsgBox "������ı���Ŀ¼: save_dir" & save_dir & " �д������˳���"
            Exit Function
        End If
    End If
    
    ''' �������ļ�
    Tristate = get_Tristate(encoding)
    Set obj_filein = fs.OpenTextFile(file_in, ForReading, False, Tristate)

    If with_title Then title = obj_filein.readline()   ''' title ��ȡ
    
    i = 1
    Do While Not obj_filein.AtEndOfStream

        row = obj_filein.readline()
        filed = Split(row, sep)(col)
        
        ''' ��ǰδ���ֵ��ֶΣ��¼�һ��buf������
        If Not dict.exists(filed) Then
            Set buf = New StrBuff
            Set buf.File = fs.OpenTextFile(basedir & filed & ".csv", _
                ForWriting, True, Tristate)
            
            If with_title Then buf.AddLine title  ' ��� title
            
            dict.Add filed, i
            bufs.Add buf
            i = i + 1
        Else
            bufs(dict(filed)).AddLine row
        End If
    Loop
    
    ''' ���д��
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
    
'����˵����
    ' file_in:string Ҫ�����ļ���ȫ·��
    ' with_title : bool ��ѡ�������ļ��ĵ�һ���ǲ��Ǳ��⡣Ĭ��True��ʾ��һ���Ǳ���
    ' sep : string optional���ֶ�֮��ķָ�����Ĭ�϶���","
    ' col : long optional ����һ��Ϊ��׼���зָĬ����0��ʾ��һ��
    
    Dim filenum!, filein!, title$
    filein = FreeFile
    Set dict = CreateObject("Scripting.Dictionary")
    
    ''' �ļ�����Ŀ¼,����Ϊ�������ļ�ͬһ��Ŀ¼
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
        If with_title Then Line Input #filein, title  ''' �ļ���һ���Ƿ��Ǳ���
        
        i = 1
        Do While Not EOF(filein)
            Line Input #filein, row
            filed = Split(row, sep)(col) '��ȡ�ֶ���

            If Not dict.exists(filed) Then  ' �ļ�δ��
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

