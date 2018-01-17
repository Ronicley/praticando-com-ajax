<!--#include file ="conexao.asp"-->
<%
response.write("<meta charset=""""utf-8"""">")
Set rs = conn.Execute("SELECT Nome, Matricula from CadFunc where Nome like '%"&ucase(request.querystring("q"))&"%'")
If Not rs.Eof Then
	While Not rs.Eof
		if  IsNull(rs("Nome"))then 
			response.Write(null)
		else
			response.write("<p>Nome: "&rs("Nome")&", Matricula: "&rs("matricula")&"</p><br>")
	  End If
	rs.movenext
	Wend
End If
rs.Close
Set rs = Nothing
%>
