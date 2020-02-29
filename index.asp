<%@LANGUAGE="JAVASCRIPT" CODEPAGE="65001"%>
<%
' *** Logout the current user.
MM_Logout = CStr(Request.ServerVariables("URL")) & "?MM_Logoutnow=1"
If (CStr(Request("MM_Logoutnow")) = "1") Then
  Session.Contents.Remove("MM_Username")
  Session.Contents.Remove("MM_UserAuthorization")
  MM_logoutRedirectPage = "index.asp"
  ' redirect with URL parameters (remove the "MM_Logoutnow" query param).
  if (MM_logoutRedirectPage = "") Then MM_logoutRedirectPage = CStr(Request.ServerVariables("URL"))
  If (InStr(1, UC_redirectPage, "?", vbTextCompare) = 0 And Request.QueryString <> "") Then
    MM_newQS = "?"
    For Each Item In Request.QueryString
      If (Item <> "MM_Logoutnow") Then
        If (Len(MM_newQS) > 1) Then MM_newQS = MM_newQS & "&"
        MM_newQS = MM_newQS & Item & "=" & Server.URLencode(Request.QueryString(Item))
      End If
    Next
    if (Len(MM_newQS) > 1) Then MM_logoutRedirectPage = MM_logoutRedirectPage & MM_newQS
  End If
  Response.Redirect(MM_logoutRedirectPage)
End If
%>
<!--#include file="Connections/leaveword.asp" -->
<%
Dim Recordset1
Dim Recordset1_cmd
Dim Recordset1_numRows

Set Recordset1_cmd = Server.CreateObject ("ADODB.Command")
Recordset1_cmd.ActiveConnection = MM_leaveword_STRING
Recordset1_cmd.CommandText = "SELECT * FROM main" 
Recordset1_cmd.Prepared = true

Set Recordset1 = Recordset1_cmd.Execute
Recordset1_numRows = 0
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>无标题文档</title>
</head>

<body>
				<p>
					<span>
						用户名
					</span>
					<input  type="text" value="<%=(Recordset1.Fields.Item("Name").Value)%>" />
				</p>
                
                				<p>
					<span>
						QQ
					</span>
					<input  type="text" value="<%=(Recordset1.Fields.Item("QQ").Value)%>" />
				</p>				<p>
					<span>
						Email
					</span>
					<input  type="text" value="<%=(Recordset1.Fields.Item("Email").Value)%>" />
				</p>				
				<p>
					<span>
						Homepage
					</span>
					<input  type="text" value="<%=(Recordset1.Fields.Item("Homepage").Value)%>" />
				</p>
                
                <p>
					<span>
						内容
					</span>
					<textarea rows="7"><%=(Recordset1.Fields.Item("content").Value)%></textarea>
                </p>
                <p><a href="new.asp">发表新留言</a></p><p><a href="login.asp">管理员入口</a></p> 
                  <a href="<%= MM_Logout %>">注销</a>
                </body>
</html>
<%
Recordset1.Close()
Set Recordset1 = Nothing
%>
