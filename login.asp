<%@ CODEPAGE = "936" %>
<%
'####################################
'#                                  #
'#        ����ASP���Ա� V1.0        #
'#                                  #
'#  �����غ� http://www.ajiang.net  #
'#      �����ʼ� zjyfc@371.net      #
'#                                  #
'#    ת�ر�����ʱ�뱣����Щ��Ϣ    #
'#                                  #
'####################################
%>
<%
'����Ա�û������������ÿ�ʼ

'ע��Ҫ����˫���ţ�
adminname="awt"	'�û���
adminpassword="20031231"	'����

'����Ա�û������������ý���

keyword=Request("keyword")
page=Request("page")
pagesize=Request("pagesize")

if Request.Form("manageid")<> "" then

if Request.Form("manageid") = adminname and Request.Form("managepassword")=adminpassword then

	Session.Contents("thegbmaster")="yes"
	response.redirect("index.asp?keyword=" & keyword & "&page=" & page & "&pagesize=" & pagesize)
else
%>
��¼ʧ�ܡ�
<%
end if
else
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>���԰�</title>
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
<span style="font-size: 10.5pt;line-height: 13pt">- �� �� �� �� �� ǩ �� -</span><BR>
</td></tr>
<tr><td align=center>�û���: <input type=text class=input3 name='manageid'></td></tr>
<tr><td align=center>��&nbsp; ��: <input class=input3 type=password name='managepassword'></td></tr>
<tr><td align=center><INPUT type="submit" value="ȷ ��" class="backc"></td></tr>
</table>
</form>
<br><br><br>
</body></html>
<%end if%>