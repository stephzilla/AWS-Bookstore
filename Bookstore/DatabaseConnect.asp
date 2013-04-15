<%
     dim objConn
     'Create a DSN-less database connection (M&A p. 512)
     Set objConn = Server.CreateObject("ADODB.Connection")
     objConn.ConnectionString="DRIVER={Microsoft Access Driver (*.mdb)};" & _
     "DBQ=\sullivs8\database\bookstore\"
     objConn.Open
%>
