<%@ CODEPAGE = "936" %>
<%
'####################################
'#                                  #
'#        阿江ASP留言本 V1.0        #
'#                                  #
'#  阿江守候 http://www.ajiang.net  #
'#      电子邮件 zjyfc@371.net      #
'#                                  #
'#    转载本程序时请保留这些信息    #
'#                                  #
'####################################
%>
<%
'管理员用户名及密码设置开始

'注意要保留双引号：
adminname="awt"	'用户名
adminpassword="20031231"	'密码

'管理员用户名及密码设置结束

keyword=Request("keyword")
page=Request("page")
pagesize=Request("pagesize")

if Request.Form("manageid")<> "" then

if Request.Form("manageid") = adminname and Request.Form("managepassword")=adminpassword then

	Session.Contents("thegbmaster")="yes"
	response.redirect("index.asp?keyword=" & keyword & "&page=" & page & "&pagesize=" & pagesize)
else
%>
登录失败。
<%
end if
else
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>留言板</title>
<style>
<!--#include file="css.css"-->
</style>
</head>
<body topmargin=0 rightmargin=0 leftmargin=0 bottommargin=0>
<CENTER>
<br><br><br>
<form method=post action="login.asp?page=<%=cstr(curpage)%>&pagesize=<%=pagesize%>&keyword=<%=keyword%>">
<table width=335 cellspacing=1 cellpadding=5>
<tr><td width=100%align=center>
<span style="font-size: 10.5pt;line-height: 13pt">- 留 言 本 管 理 签 入 -</span><BR>
</td></tr>
<tr><td align=center>用户名: <input type=text class=input3 name='manageid'></td></tr>
<tr><td align=center>密&nbsp; 码: <input class=input3 type=password name='managepassword'></td></tr>
<tr><td align=center><INPUT type="submit" value="确 定" class="backc"></td></tr>
</table>
</form>
<br><br><br>
</body></html>
<%end if%>