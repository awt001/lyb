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

'加入数据
set conn=server.createobject("adodb.connection")
DBPath = Server.MapPath("gb.mdb")
conn.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" & DBPath
Set rs = Server.CreateObject("ADODB.Recordset")
sql="select  * from guestbook where id=" & request("id")
rs.Open sql,conn,3,2

if not rs.EOF then rs.Delete

msg="删除成功。"

rs.Close
set rs=nothing
conn.Close
set conn=nothing

if error<> 0 then msg="删除过程中遇到了错误，请检查删除是否成功。"
else
msg="您不是管理员，不能删除留言。"
end if
Response.Redirect "index.asp?pagesize=" & pagesize & "&keyword=" & keyword & "&page=" & page & "&msg=" & msg
%>