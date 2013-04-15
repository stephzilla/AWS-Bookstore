<% Option Explicit %>
<!--#include virtual="/06Spring/sullivs8/database/adovbs.inc"-->
<!--#include file="DatabaseConnect.asp"-->
<!--#include file="ListAuthors.asp"-->

<html>
<head>
    <title>BamBooks.com -- Product Page</title>
    <link rel="stylesheet" href="BookStore.css" type="text/css">
</head>
<body>
 
<%  Dim objRS, strSQL, ISBN
          ISBN = request.querystring("ISBN")

        strSQL= "SELECT tblBookDescription.ISBN, tblBookDescription.strTitle, tblBookDescription.strDescription, tblBookDescription.dblPrice, tblBookDescription.strPublisher " & _
                "FROM tblBookDescription " & _ 
                "WHERE (((tblBookDescription.ISBN)='" & ISBN & "')) ;"


    Set objRS=Server.CreateObject("ADODB.Recordset")
        
        objRS.Open strSQL, ObjConn  


%>    
 
<!--#include file="includes/header.html"-->

    <div align="center">
        <center>
            <table border="0" cellpadding="5" cellspacing="0" width="800">
                <tr>
                    <td width="184" valign="top">
                        <p>
                            &nbsp;
                              
  
<!--#include file="includes/menu.html"-->
 
  
                    </td>
                    <td width="417" valign="top">
                        
                           <font face="copperplate gothic bold" color="#000000" size="4"><%= response.write(objrs("strTitle")) %></font>
                        <br>
                           <font size="-1"> by <%= funListAuthors(objRS("ISBN")) %><a href= "SearchBrowse.asp?action=search&strSearch="></a></font>
                        <br>
                        <a href="/sandvig/bookstore/images/<%= response.write(objrs("ISBN")) %>.01.LZZZZZZZ.jpg">
                            <img height="140" hspace="17" src="/sandvig/bookstore/images/<%= response.write(objrs("ISBN")) %>.01.MZZZZZZZ.jpg"
                                width="95" align="left" vspace="13" border="0">
                        </a>
                        <br>
                        <font face="arial,verdana,helvetica"><b>List Price: <font color="red"><strike>
                            <%= response.write(formatcurrency(objrs("dblPrice"))) %>
                        </strike></font>
                            <br>
                            <font face="arial,verdana,helvetica">Our Price: <font color="red">
                                <%= response.write(formatcurrency(objrs("dblPrice")- objrs("dblPrice")*.20)) %>
                            </font>
                                <br>
                                <font face="arial,verdana,helvetica">You Save: <font color="red">
                                    <%= response.write(formatcurrency(objrs("dblPrice")*.20)) %></b></font><br>
                        <br>
                        <font face="verdana,arial,helvetica" size="-1"><b>Availability:</b> Usually ships within
                            24 hours.<br>
                            ISBN:
                            <%= response.write(objrs("ISBN")) %>
                            <br>
                            Publisher:
                            <%= response.write(objrs("strPublisher")) %>
                        </font>
                        <br>
                        <a href="ShoppingCart.asp?isbn=<%= response.write(objrs("ISBN")) %>&amp;action=add">
                            <img border="0" src="/06Spring/sullivs8/Bookstore/images/cart.png" align="right"></a>
                        <br>
                        <br>
                        <br>
                            <%= response.write(objrs("strDescription")) %>
                          <br>
                    </td>
                    <td width="24">
                    </td>
                </tr>
                <tr>
                    <td width="184">
                    </td>
                    <td width="417">
                        <a href="ShoppingCart.asp?isbn=<%= response.write(objrs("ISBN")) %>&amp;action=add">
                            <img border="0" src="/06Spring/sullivs8/Bookstore/images/cart.png" align="right"></a></td>
                    <td width="24">
                    </td>
                </tr>
            </table>
        </center>
    </div>
    <br>
    <br>
      
 <!--#include file="includes/footer.html"--> 
  
 
</body>

</html>
