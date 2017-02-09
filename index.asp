<%@ CODEPAGE = "936" %>
<%
'####################################
'#                                  #
'#        阿江ASP留言本 V1.0        #
'#                                  #
'#  阿江守候 http://www.ajiang.net  #
'#      电子邮件 zjyfc@263.net      #
'#                                  #
'#    转载本程序时请保留这些信息    #
'#                                  #
'####################################
%>
<%
'下面这一行设置默认每页显示条数
thesize=10			'默认为10

pagesize=Request("pagesize")
keyword=request("keyword")
if request("page")="" then
  	curpage = 1
else
	curpage = cint(request("page"))
end if
%>
<html>
<head>
<meta http-equiv="Copyright" content="Ajiang http://www.ajiang.net">
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>留言本</title>
<link rel="stylesheet" type="text/css" href="css.css">
</head>

<body>

<!--需修改的部分 开始-->
<p><!--需修改的部分 结束-->
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber3">
  <tr>
    <td width="140" valign="top" align=center>

<!--需修改的部分 开始-->
    <p><!--需修改的部分 结束-->
    </td>
    <td><img border="0" src="guestbook.jpg"><p style="margin: 0 20">
    <font color="#FF0000">注意：</font></p>
    <p style="text-indent: 25; margin: 0 20">
    1、把留言本的公告写在这里。</p>
    <p style="text-indent: 25; margin: 0 20">2、请勿在此发布与本站无关的话题及违法的内容。</p>
    <%if Request("msg")<> "" then%><p style="text-indent: 25; margin: 0 20"><font color=red><%=Request("msg")%></font></p><%end if%><p style="text-indent: 25; margin: 0 20">　</p>
    <div align="center">
      <center>
      <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#3F8805" width="90%" id="AutoNumber4">
  <form method=post action="add.asp?keyword=<%=keyword%>&page=<%=curpage%>&pagesize=<%=pagesize%>" id=form1 name=form1>
        <tr><td colspan="2" bgcolor="#eefee0">
          &nbsp;<img src="gb-add.gif" align="absmiddle"> :::: 请 您 留 言 ::::</td>
        </tr>
        <tr>
          <td width="35%" align="center">
          <p style="margin-top: 3; margin-bottom: 3">&nbsp;姓名：<input class="input" name="name" size="22"></p>
          <p style="margin-top: 3; margin-bottom: 3">&nbsp;主页：<input class="input" name="url" size="22" value="http://"></p>
          <p style="margin-top: 3; margin-bottom: 3">&nbsp;主题：<input class="input" name="title" size="22"></p>
          <p style="margin-top: 3; margin-bottom: 3">&nbsp;Email:<input class="input" name="mail" size="22"></td>
          <td width="65%" height="130">
            <p align="center"><textarea class="input" rows="6" name="content" cols="50"></textarea><br>
&nbsp;<input class="backc" type="submit" value="提 交" name="B1">
            <input class="backc" type="reset" value="重 填" name="B1"></td>
        </tr>
  </form>
        </table>
      </center>
    </div>
    </td>
  </tr>
</table>
<%
if pagesize="0" or pagesize="" then pagesize=thesize

set conn=server.createobject("adodb.connection")
DBPath = Server.MapPath("gb.mdb")
conn.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" & DBPath
Set rs = Server.CreateObject("ADODB.Recordset")

dim rs, sql
set rs = server.createobject("adodb.recordset")
dim curpage, strcate

if keyword <> "" then 
keyword = replace(keyword,"'","")               '过滤关键字
keyword = replace(keyword,"[","")
keyword = trim(keyword)
wherestr=" where name like '%" & trim(keyword) & "%' or content like '%" & trim(keyword) & "%' or title like '%" & trim(keyword) & "%'"
end if

sql = "SELECT * FROM guestbook " & wherestr & " ORDER BY id DESC"
rs.open sql, conn, 1, 1

	if rs.bof and rs.eof then
		rs.close
		response.write "<br><center>还没有符合条件的留言呢！</center>"
	else
		dim i
		rs.pagesize = pagesize
		if rs.pagecount < curpage then
			rs.absolutepage=rs.pagecount
			curpage=rs.pagecount
		else
			rs.absolutepage = curpage
		end if

for i = 1 to rs.pagesize

%>
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber3">
  <tr>
    <td width="140" valign="top">
    </td>
    <td>
    <br>
    <div align="center">
      <center>
      <table  style="word-break:break-all" border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#3F8805" width="90%" id="AutoNumber4">
        <tr>
          <td height="20" bgcolor="#eefee0" width="75%">&nbsp;<img src="gb-index2.gif" align="absmiddle"<%if Session.Contents("thegbmaster")="yes" then%> title="IP:<%=rs("ip")%>"<%end if%>"> <%=rs("name")%> <font color="#3F8805">
          <%=formatdatetime(rs("addtime"),2) & " " & formatdatetime(rs("addtime"),4)%> 说</font>&nbsp;<%=rs("title")%></td>
          <td height="20" bgcolor="#eefee0" width="25%">
          <p align="right" style="margin-right: 10">
          <%if Session.Contents("thegbmaster")="yes" then%>
          <a title="删除" href="del.asp?id=<%=rs("id")%>&page=<%=page%>&pagesize=<%=pagesize%>&keyword=<%=keyword%>"><img border="0" src="gb-del.gif" align="absmiddle"></a>
          <a title="回复" href="reply.asp?id=<%=rs("id")%>&page=<%=page%>&pagesize=<%=pagesize%>&keyword=<%=keyword%>"><img border="0" src="gb-reply.gif" align="absmiddle"></a>
<%
end if

if rs("mail")="" or isnull(rs("mail")) then
Response.Write " <img border='0'  src=gb-mail.gif align='absmiddle'>"
else
Response.Write "<a href=mailto:" & rs("mail") & "><img border='0' alt='信箱 " & rs("mail") & "' src=gb-mail.gif align='absmiddle'></a>"
end if
Response.Write "&nbsp;"
if rs("url")="" or isnull(rs("url")) then
Response.Write "<img border='0'  src=gb-url.gif align='absmiddle'>"
else
Response.Write "<a href=" & rs("url") & " target=_blank><img border='0' alt='主页 " & rs("url") & "' src=gb-url.gif align='absmiddle'></a>"
end if
%>

		  </td>
        </tr>
        <tr>
          <td colspan="2" height="120" style="WORD-WRAP: break-word">
            <p style="line-height: 140%; margin-left: 15; margin-right: 10; margin-top: 10; margin-bottom: 5" align=left>
    <%
    Response.Write changechr(cstr(rs("content")))
    if rs("reply")<> "" then
    %>
<p style="line-height: 140%; margin-left: 15; margin-right: 10; margin-top: 10; margin-bottom: 5" align=left><font color=Orange>■ 版主回复：</font><br><font class="fonts"><%=changechr(rs("reply"))%></font>
	<%end if%>
    </p>
</td>
        </tr>
</table>
    </div>

    </td>
  </tr>
<%
  rs.movenext
    if rs.eof then
      i = i + 1
      exit for
    end if
  next
%>
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber5">
  <tr>
    <td width="140" valign="top">
    </td>
    <td>
    <br>
    <div align="center">
      <center>
      <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#3F8805" width="90%" id="AutoNumber6">
        <tr>
          <td colspan="2" height="50">
            &nbsp;<%for i=1 to rs.pagecount
if curpage <> i then%>
    <a class=a1 href="<%=Request.ServerVariables("SCRIPT_NAME")%>?page=<%=i%>&pagesize=<%=pagesize%>&keyword=<%=keyword%>">&lt;<%=i%>&gt;</a>
<%else%>
    <font style="color:#000000">&lt;<%=i%>&gt;</font>
<%end if
next
      if Session.Contents("thegbmaster")="yes" then%>
		　[<a href="logout.asp?page=<%=cstr(curpage)%>&pagesize=<%=pagesize%>&keyword=<%=keyword%>">退出管理</a>]
	  <%end if%></td>
        </tr>
        </table>
      </center>
    </div>

    </td>
  </tr>
</table>
<%
rs.close
end if

set rs=nothing
conn.Close 
set conn=nothing
%>
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber5">
<form method=post action="index.asp?page=<%=curpage%>&pagesize=<%=pagesize%>" id=form2 name=form2>

  <tr>
    <td width="140" valign="top">
    </td>
    <td>
    <br>
    <div align="center">
      <center>
      <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#3F8805" width="90%" id="AutoNumber6">
        <tr>
          <td colspan="2" height="60">
&nbsp;<img src="gb-search.gif" align=absmiddle> 留 言 搜 索：<input type=text name='keyword' size=16 class=input> <input type=submit value='搜索' class=backc id=submit1 name=submit1>
	</td>
        </tr>
        </table>
      </center>
    </div>
    </td>
  </tr>
</form>
</table>
<p align=center>

</body>
</html>
<%
function changechr(str)   
    changechr=replace(replace(replace(replace(str,"<","&lt;"),">","&gt;"),chr(13),"<br>")," ","&nbsp;")   
end function
%>