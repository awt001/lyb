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
keyword=Request("keyword")
page=Request("page")
pagesize=Request("pagesize")

'检查填写是否完整
if trim(request("name"))="" or trim(request("name"))="恶毒反对" or trim(request("content"))="" then
msg="您没有填写姓名或者没有填写留言内容或者您的名字被拒绝！"
else

'加入数据
set conn=server.createobject("adodb.connection")
DBPath = Server.MapPath("gb.mdb")
conn.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" & DBPath
Set rs = Server.CreateObject("ADODB.Recordset")
sql="select top 1 * from guestbook"
rs.Open sql,conn,3,3

url=Request("url")
if url="http://" then url=""
if len(url)>7 and left(url,7) <> "http://" then url="http://" & url

can=false

if rs.eof then
  can=true
else
  if trim(rs("content")) <> trim(Request("content")) then
  can=true
  else
  can=false
  end if
end if

if can=true then
rs.AddNew 
rs("addtime")=now
rs("ip")=Request.ServerVariables("Remote_Addr")
rs("name")=server.HTMLEncode(Request("name"))
rs("mail")=server.HTMLEncode(request("mail"))
rs("title")=server.HTMLEncode(request("title"))
rs("url")=server.HTMLEncode(url)
rs("content")=Request("content")
rs.Update 
rs.Close
msg="留言成功！"

set rs=nothing
conn.Close
set conn=nothing

end if
end if

	Response.Redirect "index.asp?pagesize=" & pagesize & "&keyword=" & keyword & "&page=1&msg=" & msg
%>