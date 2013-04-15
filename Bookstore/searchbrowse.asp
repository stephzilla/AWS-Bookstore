<% Option Explicit %>
<!--#include virtual="/06Spring/sullivs8/database/adovbs.inc"-->
<!--#include file="DatabaseConnect.asp"-->
<!--#include file="ListAuthors.asp"-->


<html>
<head>
    <title>BamBooks.com -- Search/Browse </title>
    <link rel="stylesheet" href="Bookstore.css" type="text/css">
</head>
<body>
     

<!--#include file="includes/header.html"-->


    <div align="center">
        <center>
            <table border="0" cellpadding="5" cellspacing="0" width="800">
                <tr>
                    <td width="184" valign="top">
                          
  
<!--#include file="includes/menu.html"-->

                    <!--</td>-->

<%  Dim objRS, strSearch, bolpostback, strSQL, strSQLSearch, strBrowse
               strSearch= Replace(Request.querystring("strSearch"), "'", "''")

    IF strSearch > " " THEN
      
      strSQLSearch=  "SELECT distinctrow tblBookDescription.strTitle, tblBookDescription.strDescription, tblBookDescription.ISBN "&_
                     "FROM tblAuthors INNER JOIN (tblBookDescription INNER JOIN tblAuthorsBooks ON tblBookDescription.ISBN = tblAuthorsBooks.ISBN) ON tblAuthors.AuthorID = tblAuthorsBooks.AuthorID "&_
                     "WHERE ((tblAuthors.strLastName Like '%"&strSearch&"%') "& _
                     "OR (tblAuthors.strFirstName Like '%"&strSearch&"%') "&_
                     "OR (tblBookDescription.strTitle Like '%"&strSearch&"%') "&_
                     "OR (tblBookDescription.strDescription Like '% "&strSearch&" %') "& _
                     "OR (tblBookDescription.strPublisher Like '%"&strSearch&"%')) "&_
                     "ORDER BY tblBookDescription.strTitle;"
    
              Set objRS=Server.CreateObject("ADODB.Recordset")
                  objRS.Open strSQLSearch, ObjConn, 1 
      
                  IF objrs.eof THEN
                      response.write "<td>Sorry, there were no matches to your search. Try again.</td>"
                   ELSE

IF objRS.recordcount = 1 THEN
     response.redirect ("ProductPage.asp?isbn=" & objRS("ISBN"))   
ELSE        
%>
       </td><td valign="top">We found <% =objRS.RecordCount %>  Book<% If objRS.RecordCount > 1 then response.write "s" %> for '<font color="#FF0000"><% =strSearch %></font>' <br><br>

<% 
END IF
Do while not objrs.EOF %>

                    
                   
                  <table border="0" width ="417" valign="top">
                  <font face="Times new roman" color="#000000"><u><% response.write(objrs("strTitle")) %></u></font>
                        <br>
                   <font size="-1" by <%= funListAuthors(objRS("ISBN")) %><a href= "SearchBrowse.asp?action=search&strSearch="></a></font>
                        <br>
                   <a href="ProductPage.asp?isbn=<% response.write(objrs("ISBN")) %>&amp;action=add">
                     <img src="/06Spring/sullivs8/bookstore/images/cart2.png" border="0" align="right"></a>
                   <br>
                   <br>
                   <a href="ProductPage.asp?isbn=<% response.write(objrs("ISBN")) %>&amp;action=add">
                     <img src="/sandvig/bookstore/images/<% response.write(objrs("ISBN")) %>.01.20TLZZZZ.gif" border="0" align="left" height="97" hspace="7" width="67" vspace="3"></a>
              <% response.write(Left(objrs("strDescription"), 300)) %>
                   <a href="ProductPage.asp?isbn=<% response.write(objrs("ISBN")) %>">more...</a>
                  <br>
                  <br>
                    <p>
                                </td>
                            


<% objrs.movenext
   loop

    END IF
END IF

         strBrowse=request.querystring("strBrowse")

     
           IF strBrowse > " " THEN
              'response.write("the browse is for: " & strBrowse)

   strSQL= "SELECT tblBookDescription.ISBN, tblBookDescription.strTitle, tblBookDescription.strDescription, tblCategories.strCategory " & _
           "FROM tblBookDescription INNER JOIN tblCategories ON tblBookDescription.ISBN = tblCategories.ISBN " & _
           "WHERE (((tblCategories.strCategory) = '" & strBrowse & "')); "

      set objRS= server.createobject("ADODB.recordset")
          objRS.open strSQL, objConn, 1

IF objRS.recordcount = 1 THEN
     response.redirect ("ProductPage.asp?isbn=" & objRS("ISBN"))   
ELSE   
%>       

</td><td valign="top">We found <% =objRS.RecordCount %>  Book<% If objRS.RecordCount > 1 then response.write "s" %> for  '<font color="#FF0000"><% =strBrowse %></font>' <br><br>

<% END IF
Do while not objRS.EOF %>

                  <table border="0" width ="417" valign="top">
                  <font face="Times new Roman" color="#000000"><u><%= response.write(objrs("strTitle")) %></u></font>
                        <br>
                   <font size="-1"> by <%= funListAuthors(objRS("ISBN")) %><a href= "SearchBrowse.asp?action=search&strSearch="></a></font>
                        <br>
                   <a href="ProductPage.asp?isbn=<%= response.write(objrs("ISBN")) %>&amp;action=add">
                     <img src="/06Spring/sullivs8/bookstore/images/cart2.png" border="0" align="right"></a>
                   <br>
                   <br>
                   <a href="ProductPage.asp?isbn=<%= response.write(objrs("ISBN")) %>&amp;action=add">
                     <img src="/sandvig/bookstore/images/<%= response.write(objrs("ISBN")) %>.01.20TLZZZZ.gif" border="0" align="left" height="97" hspace="7" width="67" vspace="3"></a>
              <%= response.write(Left(objrs("strDescription"), 300)) %>
                   <a href="ProductPage.asp?isbn=<%= response.write(objrs("ISBN")) %>">more...</a>
                  <br>
                  <br>
                    <p>
                    </td>
                 
                 


<%
   objRS.movenext
   loop

 END IF

   objRS.close
   set objRS = nothing
   objConn.close
   set objConn = nothing
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
