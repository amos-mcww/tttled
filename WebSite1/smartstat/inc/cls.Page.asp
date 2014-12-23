<%
'˵������:�ⲿ�ɷ��ʺ���:
'��GetPerPage��                ����            ���ص�ǰҳ
'��GetTotalPage��              ����            ������ҳ��
'��GetStartEndInfo��           ����            ������β��
'��GetStart��                  ����            ���ز�ѯ��俪ʼȡ��
'��GetEnd��                    ����            ���ز�ѯ������ȡ��
'��GetPerNextInfo��            ����            ��������ҳ����
'��GetPerNextMoreInfo��        ����            �����ϼ�ҳ�¼�ҳ����
'��GetPerNext��                ����            ���ط�װ������ҳ
'��GetPerNextMore��            ����            ���ط�װ���ϼ�ҳ�¼�ҳ
'��GetPageHeader��             �ַ���          ���ط�װ��ҳͷ��Ϣ
'��GetPageBody��               �ַ���          ���ط�װ�ķ�ҳ������Ϣ
'��GetPageEnd��                �ַ���          ���ط�װ�ķ�ҳβ����Ϣ
'��PageInfo��                  �ַ���          ���ط�װ��������ҳ��Ϣ
'��PageList��                                  ��װ�ķ�ҳ����ִ�к���

'ʹ��˵��:
'����ʹ�õ��ļ���Include()���ļ�,ʹ��New�������´���һ���������´�������ʱ�����ù��캯���������
'����˵�����¹��캯��˵������

Class PageList
'-------------------------���������õ���ȫ�ֱ���----------------------
    Private PerPage                 '��ǰҳ��
    Private PerLimit                '��ǰÿҳ��ʾ����
    Private PerPageLimit            '��ǰÿҳ��ʾҳ��
    Private TotalNums               '��ǰ��ҳ�е�������
    Private TotalPage               '��ǰ��ҳ����ҳ��
    Private PageUrl                 '���嵱ǰ��ҳ·��
    Private PageStart               '���嵱ǰ��ʼĩID
    Private PageEnd                 '���嵱ǰ����ID
    Private PageStyle               '������ǰ��ҳ��ʾ��ʽ
    Private PageHeader              '������ʾͷ����
    Private PageBody                '������ʾ���岿��
    Private PageBottom              '������ʾβ����
    Private DisplayInfo

'-----------------------------------��ʼ�����й̶�����-------------------------------
'��ڲ���:       
'pageParameter(0�� :��Ϣ������
'pageParameter(1��:ÿҳ��ʾ����
'pageParameter(2��:ÿҳ��ʾҳ��
'styleParameter(0):����̬���trueΪ��̬ falseΪ��̬ noΪ����ʾ
'styleParameter(1):ǰ��ҳ��ҳ,��һҳ��һҳ���.trueΪǰ��ҳ falseΪ��һҳ noΪ����ʾ
'styleParameter(2):������ת����trueΪ���� falseΪsubmit��input noΪ����ʾ
'styleParameter(3):�����Ƿ���ʾ���岿��

	'��ʼ��������
    Private Sub Class_Initialize()
        On Error Resume Next
        Set PageStyle = CreateObject("Scripting.Dictionary")
            PageStyle.Add "DS","false"
            PageStyle.Add "PN","true"
            PageStyle.Add "JP","true"
            PageStyle.Add "BY","true"
            PerPage=cint(Request.QueryString("page"))
    End Sub

    '�ڴ�˵�����������Ϊ��̬�ġ���һҳ��һҳ�����Ϊ�в���̬���������Ϊ��̬��ֻ��Ϊ�ϼ�ҳ���¼�ҳ
    '�������Ϊ��̬�ġ��ϼ�ҳ�¼�ҳ�����Ϊ������̬
    Public Property Let InitPara(PLRs)
        If isArray(PLRs) Then 
            TotalNums=PLRs(0)
            PerLimit=PLRs(1)
            PerPageLimit=PLRs(2)
        Else
            TotalNums=1
            PerLimit=1
            PerPageLimit=1
        End If
    End Property
    Public Property Let DisplayStyle(styleParameter)
            If isArray(styleParameter) Then
            PageStyle.Item("DS")=styleParameter(0)        '��̬��ʽ
            PageStyle.Item("PN")=styleParameter(1)        '����ҳ��ʽ
            PageStyle.Item("JP")=styleParameter(2)        '����ʽ���
            PageStyle.Item("BY")=styleParameter(3)        '�������岿�ַ��
        End If
    End Property
    '-------------����õ�ǰ̨��ʾ��������---------------
    Public Property Get PageInfo()
        PageInfo=DisplayInfo
    End Property
    '-------------����ȡ�õ�ǰ��ҳ�еĵ�ǰҳ����------------------------------
    Public ProPerty Get GetPerPage()
            GetPerPage=PerPage
    End ProPerty 
    '-------------����ȡ�õ�ǰ��ҳ������ҳ������------------------------------
    Public ProPerty Get GetTotalPage()
        GetTotalPage=TotalPage
    End ProPerty

    Public ProPerty Get GetStart()
        GetStart=(PerPage-1)*PerLimit+1
    End ProPerty
    '-----------------����ȡ�õ�ǰҳ����Ϣ��β�ź���---------------------------
    Public Property Get GetEnd()
        Dim Ends
        Ends=GetStart()+PerLimit-1
        If Ends>TotalNums Then Ends=TotalNums
        GetEnd=Ends
    End Property
    '--------------���崴��������ʾ����------------------------------------
    Public Sub PageList()
        Call SetDefaultStyle()
        Call SetPageUrl()
        Call SetToTalPage()
        Call SetPerPage()
        Call SetStaticDynamic()
        Call SetDisplayPageInfo()
    End Sub


    Public Sub PageListUrl(urlTmp)
        Call SetDefaultStyle()
        Call SetPageUrl_single(urlTmp)
        Call SetToTalPage()
        Call SetPerPage()
        Call SetStaticDynamic()
        Call SetDisplayPageInfo()
    End Sub

    '------------��������ǿ�Ʒ����----------------------------------------
    Private Sub SetDefaultStyle()
        If PageStyle("DS")="false" Then
            If PageStyle("PN")<>"no" Then PageStyle("PN")="true"
        End If
        If PageStyle("BY")="no" Then PageStyle("PN")="false"
    End Sub
    
    Private Sub SetPageUrl_single(urlTmp)
        PageUrl = urlTmp & "&"
    End Sub

    '-------------����ȡ�õ�ǰ��ҳ·������----------------------------------
    Private Sub SetPageUrl()
        On Error Resume Next
        Dim queryString
        Dim tmp_str
        Dim tmp_key,tmp_array
        queryString =Request.ServerVariables("QUERY_STRING")
        queryString = Split(Request.ServerVariables("QUERY_STRING"), "&")
        For Each tmp_key In queryString
            tmp_array = Split(tmp_key,"=")
            If tmp_array(0)<>"page" Then
                tmp_str = tmp_str & tmp_array(0) & "=" & tmp_array(1) & "&"
            End If
        Next 
        PageUrl="http://" & Request.ServerVariables("HTTP_HOST") & Request.ServerVariables("SCRIPT_NAME")&"?"&tmp_str
    End Sub
    '-------------�������õ�ǰ��ҳ������ҳ������------------------------------
    Private Sub SetTotalPage()
        TotalPage=Abs(Int(-TotalNums/PerLimit))
    End Sub
    '--------------�������õ�ǰ��ҳ���е�ǰҳ��------------------------------
    Private Sub SetPerPage()
        If PerPage<=0 Then
            PerPage=1
        elseIf PerPage>TotalPage Then
            PerPage=TotalPage
        End IF
    End Sub
    '----------����ȡ���Ƿ�������ҳ����-----------------------------------
    Public Function GetPerNextInfo()
        Dim per,nexts
        If PerPage>1 Then
            per=PerPage-1
        Else
            per=false
        End If
        If PerPage<TotalPage Then
            nexts=PerPage+1
        Else 
            nexts=false
        End If
        GetPerNextInfo=array(per,nexts)
    End Function
    '-----------------����ȡ���Ƿ����ϼ�ҳ���¼�ҳ����-----------------
    Public Function GetPerNextMoreInfo()
        Dim start,ends
        If PerPage>PerPageLimit Then
            start=PageStart-1
        Else
            start=false
        End If
        If PageEnd<TotalPage Then
            ends=PageEnd+1
        Else
            ends=false
        End If
        GetPerNextMoreInfo=array(start,ends)
    End Function
    '-----------------����ȡ�õ�ǰ������̬��ҳ��Ϣ����----------------------------
    Private Sub SetDynamicPageInfo()
        PageStart=PerPage
        If PageStart<=0 Then PageStart=1
        PageEnd=PerPageLimit+PageStart-1
        If PageEnd>TotalPage Then PageEnd=TotalPage
    End Sub
    '-----------------����ȡ�õ�ǰ�м䶯̬��ҳ��Ϣ����----------------------------
    Private Sub SetMIDDynamicPageInfo()
        Dim add
        add=fix(PerPageLimit/2)
        PageStart=PerPage-add
        If PageStart<=0 Then PageStart=1
        PageEnd=PerPageLimit+PageStart-1
        If PageEnd>TotalPage Then PageEnd=TotalPage
    End Sub
    '-------------����ȡ�õ�ǰ��̬��ҳ��Ϣ����-------------------------------
    Private Sub SetStaticPageInfo()
        starts=PerPage/PerPageLimit
        If starts-fix(starts)>0 Then
            starts=fix(starts)
        Else
            starts=starts-1
        End If
        PageStart=starts*PerPageLimit+1
        PageEnd=PageStart+PerPageLimit-1
        If PageEnd>TotalPage Then PageEnd=TotalPage
    End Sub

    '------------����ͨ���ж��Ǿ���̬��������β��----------------
    Private Sub SetStaticDynamic()
        If PageStyle("DS")="false" Then
            SetStaticPageInfo()
        ElseIF PageStyle("DS")="true" Then
            If PageStyle("PN")="true" Then
                SetDynamicPageInfo()
            Else
                SetMIDDynamicPageInfo()    
            End If
        Else
            SetStaticPageInfo()        
        End If
    End Sub

    '-------------������ʾǰ��ҳ�ͺ�ҳ����----------------------------
    Public Function GetPerNextMore(mark)
        Dim tmp_str
        If PageStyle("PN")="false" or PageStyle("PN")="no" Then    
            PageBody=PageBody&""
            GetPerNextMore=""
            Exit Function
        End If
        IfPerNextArray=GetPerNextMoreInfo()
        If PageStyle("DS")="true" and PageStyle("DS")="true" Then
            IfPerNextArray(0)=IfPerNextArray(0)+1-PerPageLimit
            If IfPerNextArray(0)<=0 Then IfPerNextArray(0)=1    
            If PerPage=1 Then IfPerNextArray(0)=false
        End If
        If IfPerNextArray(0) and mark="true" Then tmp_str="<a href="&PageUrl&"page="&IfPerNextArray(0)&" title=ǰ"&PerPageLimit&"ҳ style='margin:5px;'> < </a>"&chr(10)
        If IfPerNextArray(1) and mark<>"true" Then tmp_str="<a href="&PageUrl&"page="&IfPerNextArray(1)&" title=��"&PerPageLimit&"ҳ style='margin:5px;'> > </a>"&chr(10)
        PageBody=PageBody&tmp_str
        GetPerNextMore=tmp_str
    End Function

        '-------------������ʾǰһҳ�ͺ�һҳ����----------------------------
    Public Function GetPerNext(mark)
        Dim tmp_str
        If PageStyle("PN")<>"false" Then 
            GetPerNext=""
            Exit Function
        End If
        IfPerNextArray=GetPerNextInfo()
        If IfPerNextArray(0) and mark="true" Then tmp_str="<a href="&PageUrl&"page="&IfPerNextArray(0)&" title= ǰһҳ> < </a>"
        If IfPerNextArray(1) and mark<>"true" Then tmp_str="<a href="&PageUrl&"page="&IfPerNextArray(1)&" title=��һҳ> > </a>"
        PageBody=PageBody&tmp_str
        GetPerNext=tmp_str
    End Function

    '-------------------������ʾ��ҳͷ����-------------------
    Public Function GetPageHeader()
        Dim tmp_str
        If PageStyle("DS")="no" Then
            GetPageHeader=""
            Exit Function
        End If
        tmp_str="��"&PerPage&"ҳ/��"&GetTotalPage()&"ҳ ��"&TotalNums&"��"
        PageHeader="<div style='float:right;display:inline;'>"&tmp_str&""&Chr(10)
        GetPageHeader=tmp_str
    End Function

    '-------------------������ʾ��ҳ���庯��-------------------
    Public Function GetPageBody()
        If PageStyle("BY")="no" Then
            GetPageBody=""
            Exit Function
        End IF
        For PageStart=PageStart to PageEnd
            If PerPage=PageStart Then
                PageBody=PageBody&"<a href="&PageUrl&"page="&PageStart&" title=��"&PageStart&"ҳ>["&PageStart&"]</a>"&Chr(10)
            Else
                PageBody=PageBody&"<a href="&PageUrl&"page="&PageStart&" title=��"&PageStart&"ҳ>["&PageStart&"]</a>"&Chr(10)
            End If
        Next
        GetPageBody=PageBody
    End Function

    '-------------------������ʾ��ҳ���庯��-------------------
    Private Sub SetPageBody()
        GetPerNextMore("true")
        GetPerNext("true")
        GetPageBody()        
        GetPerNextMore("false")
        GetPerNext("false")
        PageBody="Page: "&PageBody&""&Chr(10)
    End Sub

    '-------------------������ʾ��ҳβ����-------------------
    Public Function GetPageEnd()
        Dim tmp_str
        If PageStyle("JP")="no" Then
            GetPageEnd=""
            PageBottom=""
            Exit Function
        ElseIF PageStyle("JP")="true" Then
            tmp_str="&nbsp;&nbsp;����<select name='pageSelect' onChange='document.location=this.value'>"
            for i=1 to GetTotalPage()
                If i=PerPage Then
                    tmp_str=tmp_str&"<option value="&PageUrl&"page="&i&" selected>"&i&"</option>"
                Else
                    tmp_str=tmp_str&"<option value="&PageUrl&"page="&i&">"&i&"</option>"
                End If
            Next
            tmp_str=tmp_str&"</select>ҳ"
        ElseIF PageStyle("JP")="false" Then
            tmp_str="<input type='text' name='pageSelect' ID='pageSelect' size='3' maxlength='5' value='"&PerPage&"'>"
            tmp_str=tmp_str&"&nbsp;<input type='button' value='GO' onClick=""document.location='"&PageUrl&"page='+ pageSelect.value"">"
        End If
        PageBottom=""&tmp_str&"</div>"&Chr(10)
        GetPageEnd=tmp_str
    End Function

    '-------------����õ�ǰ̨��ʾ��������---------------
    Private Sub SetDisplayPageInfo()
        GetPageHeader()
        SetPageBody()
        GetPageEnd()
        DisplayInfo="<div style='margin:5px;'>"&chr(10)&PageHeader&PageBody&PageBottom&"</div>"
    End Sub
End Class
%>