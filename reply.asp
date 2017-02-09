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
		msg="回复成功。"
		else
		msg="没有这条留言。"
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
<title>留言本</title>
<link rel="stylesheet" type="text/css" href="css.css">
</head>

<body>
<form action=reply.asp method=post>
<input type="hidden" name="keyword" value="<%=keyword%>">
<input type="hidden" name="page" value="<%=page%>">
<input type="hidden" name="pagesize" value="<%=pagesize%>">
<input type="hidden" name="id" value="<%=Request("id")%>">
<TEXTAREA class="input" rows=5 cols=40 name="reply"></TEXTAREA>
<INPUT type="submit" class="backc" value="确定" id=submit1 name=submit1>
</form>
</body>

</html>
<%
	end if
	
else
msg="您不是管理员，不能修改留言。"
Response.Redirect "index.asp?pagesize=" & pagesize & "&keyword=" & keyword & "&page=" & page & "&msg=" & msg
end if
%>