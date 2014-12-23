<%
'说明部分:外部可访问函数:
'【GetPerPage】                整型            返回当前页
'【GetTotalPage】              整型            返回总页数
'【GetStartEndInfo】           数组            返回首尾号
'【GetStart】                  整型            返回查询语句开始取数
'【GetEnd】                    整型            返回查询语句结束取数
'【GetPerNextInfo】            数组            返回上下页数字
'【GetPerNextMoreInfo】        数组            返回上几页下几页数字
'【GetPerNext】                数组            返回封装的上下页
'【GetPerNextMore】            数组            返回封装的上几页下几页
'【GetPageHeader】             字符串          返回封装分页头信息
'【GetPageBody】               字符串          返回封装的分页主体信息
'【GetPageEnd】                字符串          返回封装的分页尾部信息
'【PageInfo】                  字符串          返回封装的完整分页信息
'【PageList】                                  封装的分页主体执行函数

'使用说明:
'在所使用的文件中Include()本文件,使用New操作符新创建一个本对象。新创建对象时需设置构造函数所需参数
'参数说明如下构造函数说明处！

Class PageList
'-------------------------定义类中用到的全局变量----------------------
    Private PerPage                 '当前页数
    Private PerLimit                '当前每页显示条数
    Private PerPageLimit            '当前每页显示页数
    Private TotalNums               '当前分页中的总条数
    Private TotalPage               '当前分页中总页数
    Private PageUrl                 '定义当前网页路径
    Private PageStart               '定义当前开始末ID
    Private PageEnd                 '定义当前结束ID
    Private PageStyle               '定主当前分页显示样式
    Private PageHeader              '定义显示头部分
    Private PageBody                '定义显示主体部分
    Private PageBottom              '定义显示尾部分
    Private DisplayInfo

'-----------------------------------初始化所有固定变量-------------------------------
'入口参数:       
'pageParameter(0） :信息总条数
'pageParameter(1）:每页显示条数
'pageParameter(2）:每页显示页数
'styleParameter(0):动静态风格，true为动态 false为静态 no为不显示
'styleParameter(1):前几页后几页,上一页下一页风格.true为前几页 false为上一页 no为不显示
'styleParameter(2):设置跳转函数true为下拉 false为submit加input no为不显示
'styleParameter(3):设置是否显示主体部分

	'初始化对像函数
    Private Sub Class_Initialize()
        On Error Resume Next
        Set PageStyle = CreateObject("Scripting.Dictionary")
            PageStyle.Add "DS","false"
            PageStyle.Add "PN","true"
            PageStyle.Add "JP","true"
            PageStyle.Add "BY","true"
            PerPage=cint(Request.QueryString("page"))
    End Sub

    '在此说明：如果设置为动态的、上一页下一页风格则为中部动态。如果设置为静态则只能为上几页和下几页
    '如果设置为动态的、上几页下几页风格则为步进动态
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
            PageStyle.Item("DS")=styleParameter(0)        '动态样式
            PageStyle.Item("PN")=styleParameter(1)        '上下页样式
            PageStyle.Item("JP")=styleParameter(2)        '表单样式风格
            PageStyle.Item("BY")=styleParameter(3)        '设置主体部分风格
        End If
    End Property
    '-------------定义得到前台显示变量函数---------------
    Public Property Get PageInfo()
        PageInfo=DisplayInfo
    End Property
    '-------------定义取得当前分页中的当前页函数------------------------------
    Public ProPerty Get GetPerPage()
            GetPerPage=PerPage
    End ProPerty 
    '-------------定义取得当前分页类中总页数函数------------------------------
    Public ProPerty Get GetTotalPage()
        GetTotalPage=TotalPage
    End ProPerty

    Public ProPerty Get GetStart()
        GetStart=(PerPage-1)*PerLimit+1
    End ProPerty
    '-----------------定义取得当前页中信息首尾号函数---------------------------
    Public Property Get GetEnd()
        Dim Ends
        Ends=GetStart()+PerLimit-1
        If Ends>TotalNums Then Ends=TotalNums
        GetEnd=Ends
    End Property
    '--------------定义创建分面显示内容------------------------------------
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

    '------------定义设置强制风格函数----------------------------------------
    Private Sub SetDefaultStyle()
        If PageStyle("DS")="false" Then
            If PageStyle("PN")<>"no" Then PageStyle("PN")="true"
        End If
        If PageStyle("BY")="no" Then PageStyle("PN")="false"
    End Sub
    
    Private Sub SetPageUrl_single(urlTmp)
        PageUrl = urlTmp & "&"
    End Sub

    '-------------定认取得当前网页路径函数----------------------------------
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
    '-------------定义设置当前分页类中总页数函数------------------------------
    Private Sub SetTotalPage()
        TotalPage=Abs(Int(-TotalNums/PerLimit))
    End Sub
    '--------------定义设置当前分页类中当前页数------------------------------
    Private Sub SetPerPage()
        If PerPage<=0 Then
            PerPage=1
        elseIf PerPage>TotalPage Then
            PerPage=TotalPage
        End IF
    End Sub
    '----------定义取得是否有上下页函数-----------------------------------
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
    '-----------------定义取得是否有上几页和下几页函数-----------------
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
    '-----------------定义取得当前步进动态分页信息函数----------------------------
    Private Sub SetDynamicPageInfo()
        PageStart=PerPage
        If PageStart<=0 Then PageStart=1
        PageEnd=PerPageLimit+PageStart-1
        If PageEnd>TotalPage Then PageEnd=TotalPage
    End Sub
    '-----------------定义取得当前中间动态分页信息函数----------------------------
    Private Sub SetMIDDynamicPageInfo()
        Dim add
        add=fix(PerPageLimit/2)
        PageStart=PerPage-add
        If PageStart<=0 Then PageStart=1
        PageEnd=PerPageLimit+PageStart-1
        If PageEnd>TotalPage Then PageEnd=TotalPage
    End Sub
    '-------------定义取得当前静态分页信息函数-------------------------------
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

    '------------定义通过判断是静动态来设置首尾号----------------
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

    '-------------定义显示前几页和后几页函数----------------------------
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
        If IfPerNextArray(0) and mark="true" Then tmp_str="<a href="&PageUrl&"page="&IfPerNextArray(0)&" title=前"&PerPageLimit&"页 style='margin:5px;'> < </a>"&chr(10)
        If IfPerNextArray(1) and mark<>"true" Then tmp_str="<a href="&PageUrl&"page="&IfPerNextArray(1)&" title=后"&PerPageLimit&"页 style='margin:5px;'> > </a>"&chr(10)
        PageBody=PageBody&tmp_str
        GetPerNextMore=tmp_str
    End Function

        '-------------定义显示前一页和后一页函数----------------------------
    Public Function GetPerNext(mark)
        Dim tmp_str
        If PageStyle("PN")<>"false" Then 
            GetPerNext=""
            Exit Function
        End If
        IfPerNextArray=GetPerNextInfo()
        If IfPerNextArray(0) and mark="true" Then tmp_str="<a href="&PageUrl&"page="&IfPerNextArray(0)&" title= 前一页> < </a>"
        If IfPerNextArray(1) and mark<>"true" Then tmp_str="<a href="&PageUrl&"page="&IfPerNextArray(1)&" title=后一页> > </a>"
        PageBody=PageBody&tmp_str
        GetPerNext=tmp_str
    End Function

    '-------------------定义显示分页头函数-------------------
    Public Function GetPageHeader()
        Dim tmp_str
        If PageStyle("DS")="no" Then
            GetPageHeader=""
            Exit Function
        End If
        tmp_str="第"&PerPage&"页/共"&GetTotalPage()&"页 共"&TotalNums&"条"
        PageHeader="<div style='float:right;display:inline;'>"&tmp_str&""&Chr(10)
        GetPageHeader=tmp_str
    End Function

    '-------------------定义显示分页主体函数-------------------
    Public Function GetPageBody()
        If PageStyle("BY")="no" Then
            GetPageBody=""
            Exit Function
        End IF
        For PageStart=PageStart to PageEnd
            If PerPage=PageStart Then
                PageBody=PageBody&"<a href="&PageUrl&"page="&PageStart&" title=第"&PageStart&"页>["&PageStart&"]</a>"&Chr(10)
            Else
                PageBody=PageBody&"<a href="&PageUrl&"page="&PageStart&" title=第"&PageStart&"页>["&PageStart&"]</a>"&Chr(10)
            End If
        Next
        GetPageBody=PageBody
    End Function

    '-------------------定义显示分页主体函数-------------------
    Private Sub SetPageBody()
        GetPerNextMore("true")
        GetPerNext("true")
        GetPageBody()        
        GetPerNextMore("false")
        GetPerNext("false")
        PageBody="Page: "&PageBody&""&Chr(10)
    End Sub

    '-------------------定义显示分页尾函数-------------------
    Public Function GetPageEnd()
        Dim tmp_str
        If PageStyle("JP")="no" Then
            GetPageEnd=""
            PageBottom=""
            Exit Function
        ElseIF PageStyle("JP")="true" Then
            tmp_str="&nbsp;&nbsp;跳到<select name='pageSelect' onChange='document.location=this.value'>"
            for i=1 to GetTotalPage()
                If i=PerPage Then
                    tmp_str=tmp_str&"<option value="&PageUrl&"page="&i&" selected>"&i&"</option>"
                Else
                    tmp_str=tmp_str&"<option value="&PageUrl&"page="&i&">"&i&"</option>"
                End If
            Next
            tmp_str=tmp_str&"</select>页"
        ElseIF PageStyle("JP")="false" Then
            tmp_str="<input type='text' name='pageSelect' ID='pageSelect' size='3' maxlength='5' value='"&PerPage&"'>"
            tmp_str=tmp_str&"&nbsp;<input type='button' value='GO' onClick=""document.location='"&PageUrl&"page='+ pageSelect.value"">"
        End If
        PageBottom=""&tmp_str&"</div>"&Chr(10)
        GetPageEnd=tmp_str
    End Function

    '-------------定义得到前台显示变量函数---------------
    Private Sub SetDisplayPageInfo()
        GetPageHeader()
        SetPageBody()
        GetPageEnd()
        DisplayInfo="<div style='margin:5px;'>"&chr(10)&PageHeader&PageBody&PageBottom&"</div>"
    End Sub
End Class
%>