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

Session.Contents("thegbmaster")=""

Response.Redirect "index.asp?pagesize=" & pagesize & "&keyword=" & keyword & "&page=" & page

%>
