<% Option Explicit %>
<!--#include virtual="/sullivs8/bookstore/database/adovbs.inc"-->
<!--#include file="/sullivs8/bookstore/DatabaseConnect.asp"-->


<html>
<head>
    <title>BamBooks.com</title>
    <link rel="stylesheet" href="Bookstore.css" type="text/css">
</head>
<body>
     
<%  Dim objRS, strSQL, ibookID, ibookID2, ibookID3
               randomize
               iBookID = int (11 * rnd(1)) +1 
               iBookID2 = ibookID +1 
               iBookID3 = ibookID2 +1 


       strSQL = "SELECT tblBookDescription.ISBN, tblBookDescription.BookID, tblBookDescription.strTitle, tblBookDescription.strDescription " & _
                "FROM tblBookDescription " & _
                "WHERE (((tblBookDescription.BookID) = " & ibookID & ")) OR (((tblBookDescription.BookID) = " & ibookID2 & ")) OR (((tblBookDescription.BookID) = " & ibookID3 & ")); "

    Set objRS=Server.CreateObject("ADODB.Recordset")
        
        objRS.Open strSQL, ObjConn  


%>

<!--#include file="/sullivs8/bookstore/includes/header.html"-->

    <div align="center">
        <center>
            <table border="0" cellpadding="5" cellspacing="0" width="800">
                <tr>
                    <td width="184" valign="top">
                          
  
<!--#include file="includes/menu.html"-->
<% dim iItemCount
    iItemcount= request.cookies("ShoppingCart")("ItemCount") 
%>
<br>
<br>
<center>You have<a href="shoppingcart.asp"> <% =iItemCount %> </a> item<% If iItemCount <> 1 then response.write "s" %> in your cart.



                    </td>
                    <td valign="top">
                        <font face="copperplate gothic bold" color="#000000" size="4">Today's Featured Books</font>

<% Do while not objRS.eof %>         
               
                  <table border="0">
                         <tr>
                                <td>
                                    <p>
                                        <font face="Times New Roman" color="#000000"><u><%= response.write(objRS("strTitle")) %></u></font>
                                        <br>
                                        <a href="ProductPage.asp?isbn=<%= response.write(objRS("ISBN")) %>">
                                            <img height="97" width="69" hspace="7" vspace="3" src="/sandvig/bookstore/images/<%= response.write(objRS("ISBN")) %>.01.20TLZZZZ.gif"
                                                align="left" border="0">
                                        </a>
                                        <%= response.write(Left(objRS("strDescription"), 300)) %>
                                        <a href="ProductPage.asp?isbn=<%= response.write(objRS("ISBN")) %>">more...</a>
                                </td>
                            </tr>
<% objrs.movenext
   loop
   objRS.close
   set objrs = nothing
   objconn.close
   set objconn = nothing
%>
                        </table>
                    </td>
                </tr>
            </table>
        </center>
    </div>
      

 <!--#include file="includes/footer.html"--> 
 
</body>
</html>
