<% option explicit %>
<!--#include file="DatabaseConnect.asp"-->
<!--#include virtual="/06spring/sullivs8/database/adovbs.inc"-->
<!--#include file="ListAuthors.asp"-->

<html>
<head>
<title>BamBooks.com -- Order History</title>
</head>
<body>
<!--#include file="includes/header.html"-->



<center><font face='copperplate gothic bold' color='#000000'>Order History</font><br><br>
    
<%
Dim strSQL, objRS, strEmail

strEmail = request.querystring("email")

    strSQL = "SELECT tblBookDescription.ISBN, tblBookDescription.strTitle, tblBookDescription.dblPrice, tblOrders.dtOrderDate, tblOrders.strEmail " & _
             "FROM tblBookDescription INNER JOIN tblOrders ON tblBookDescription.ISBN = tblOrders.ISBN " & _
             "WHERE (((tblOrders.strEmail)='" & strEmail & "'));"

                  Set objRS = Server.CreateObject("ADODB.Recordset")
                      objRS.Open strSQL, objConn, 1
%>
		<center><font face="copperplate gothic light">You have ordered <% =objRS.recordcount %> books.</font></center><br>


		
		<div align="center">
  		<center>
  		<table border="0" cellpadding="5" cellspacing="0" width="600">
    	<tr>
      		<td><font face="Times New Roman"><u>Order Date</u></font></td>
      		<td><font face="Times new Roman"><u>Title</u></font></td>
      		<td>&nbsp;</td>
      	</tr>
<%
Do While NOT objRS.EOF
%>
    	
        <tr>
      		<td><%= objRS("dtOrderDate") %></td>
      		<td><font face="copperplate gothic light" color="#FF0000"><% =objRS("strTitle") %> </font> 
                    <font size="-1">
                       by <a href='SearchBrowse.asp?action=search&strSearch='><% =funListAuthors(objRS("ISBN"))%></a>
                    </font>
                    <br>
                    <FONT face=arial,verdana,helvetica>
                        Your price only:<font color=#000000> <b> <% =objRS("dblPrice") %></b></font>
                </td>
      		<td>
      			<a href="ProductPage.asp?isbn=<% =objRS("ISBN") %>
        		<img height="97" hspace="7" src="/sandvig/bookstore/images/<% =objRS("ISBN") %>.01.20TLZZZZ.gif" width="67" align="left" vspace="3" border="0">
        		</a>
		</td>
    	</tr>

<%
objRS.MoveNext
Loop
%>
  	</table>
  	<br>
  	<a href="default.asp"><img border="0" src="/06Spring/sullivs8/bookstore/images/continue.png" width="121" height="19"></a><br>
 
  	</center>
	</div>
     
<!--#include file="includes/footer.html"-->  
 


</body>
