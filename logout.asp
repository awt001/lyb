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

Session.Contents("thegbmaster")=""

Response.Redirect "index.asp?pagesize=" & pagesize & "&keyword=" & keyword & "&page=" & page

%>
