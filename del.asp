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

'��������
set conn=server.createobject("adodb.connection")
DBPath = Server.MapPath("gb.mdb")
conn.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" & DBPath
Set rs = Server.CreateObject("ADODB.Recordset")
sql="select  * from guestbook where id=" & request("id")
rs.Open sql,conn,3,2

if not rs.EOF then rs.Delete

msg="ɾ���ɹ���"

rs.Close
set rs=nothing
conn.Close
set conn=nothing

if error<> 0 then msg="ɾ�������������˴�������ɾ���Ƿ�ɹ���"
else
msg="�����ǹ���Ա������ɾ�����ԡ�"
end if
Response.Redirect "index.asp?pagesize=" & pagesize & "&keyword=" & keyword & "&page=" & page & "&msg=" & msg
%>