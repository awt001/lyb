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
keyword=Request("keyword")
page=Request("page")
pagesize=Request("pagesize")

'�����д�Ƿ�����
if trim(request("name"))="" or trim(request("name"))="�񶾷���" or trim(request("content"))="" then
msg="��û����д��������û����д�������ݻ����������ֱ��ܾ���"
else

'��������
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
msg="���Գɹ���"

set rs=nothing
conn.Close
set conn=nothing

end if
end if

	Response.Redirect "index.asp?pagesize=" & pagesize & "&keyword=" & keyword & "&page=1&msg=" & msg
%>