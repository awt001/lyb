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

if Session.Contents("thegbmaster")="yes" then

keyword=Request("keyword")
page=Request("page")
pagesize=Request("pagesize")

	if Request("reply")<> "" then

		set conn=server.createobject("adodb.connection")
		DBPath = Server.MapPath("gb.mdb")
		conn.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" & DBPath
		Set rs = Server.CreateObject("ADODB.Recordset")
		sql="select * from guestbook where id=" & request("id")
		rs.Open sql,conn,3,2

		if not rs.EOF then
		rs("reply")=Request("reply")
		rs.Update
		msg="�ظ��ɹ���"
		else
		msg="û���������ԡ�"
		end if
		rs.Close
		set rs=nothing
		conn.Close
		set conn=nothing
		Response.Redirect "index.asp?pagesize=" & pagesize & "&keyword=" & keyword & "&page=" & page & "&msg=" & msg
	else%>
<html>
<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>���Ա�</title>
<link rel="stylesheet" type="text/css" href="css.css">
</head>

<body>
<form action=reply.asp method=post>
<input type="hidden" name="keyword" value="<%=keyword%>">
<input type="hidden" name="page" value="<%=page%>">
<input type="hidden" name="pagesize" value="<%=pagesize%>">
<input type="hidden" name="id" value="<%=Request("id")%>">
<TEXTAREA class="input" rows=5 cols=40 name="reply"></TEXTAREA>
<INPUT type="submit" class="backc" value="ȷ��" id=submit1 name=submit1>
</form>
</body>

</html>
<%
	end if
	
else
msg="�����ǹ���Ա�������޸����ԡ�"
Response.Redirect "index.asp?pagesize=" & pagesize & "&keyword=" & keyword & "&page=" & page & "&msg=" & msg
end if
%>