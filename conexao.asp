
<%
Response.CharSet = "utf-8"
Set conn = Server.CreateObject( "ADODB.Connection" )
Set rs = Server.CreateObject( "ADODB.Recordset" )
Set rs2 = Server.CreateObject( "ADODB.Recordset" )
Set rs3 = Server.CreateObject( "ADODB.Recordset" )
Set rs4 = Server.CreateObject( "ADODB.Recordset" )
Set rs5 = Server.CreateObject( "ADODB.Recordset" )
Set rs6 = Server.CreateObject( "ADODB.Recordset" )
Set rs7 = Server.CreateObject( "ADODB.Recordset" )
Set rs8 = Server.CreateObject( "ADODB.Recordset" )
conn.Open "Driver={SQL Server};Server=ADAPECSERV08;Database=Adapec; uid=sa; password=4d4p3ct0.2"
%>