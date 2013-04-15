<% option explicit %>
<!--#include file="DatabaseConnect.asp"-->
<!--#include virtual="/06spring/sullivs8/database/adovbs.inc"-->
<!--#include file="ListAuthors.asp"-->

<html>
<head><title>Checkout</title></head>
<body>
     
<!--#include file="includes/Header.html"-->

<% dim strEmail
       strEmail = request("email")

IF len(stremail) < 6 THEN
%>
   <center>
        <font face='Comic Sans MS' color='#FF0000'>Order Information<br></font><br>
        <font face='Comic Sans MS' color='#FF0000'>Please provide your e-mail address.</font><br><br><INPUT TYPE='button' VALUE='< Go Back'onClick='history.back()'><br>

<% 
    response.end
END IF

   Dim objRS, strName, strAddress, strCity, strState, strZip, strSQL

          StrSQL = "SELECT tblCustomers.strName, tblCustomers.strAddress, tblCustomers.strCity, tblCustomers.strState, tblCustomers.strZip, tblCustomers.strEmail " & _
                   "FROM tblCustomers " & _
                   "WHERE (((tblCustomers.strEmail)='" & strEmail & "'));"

                      Set objRS = Server.CreateObject("ADODB.Recordset")
                          objRS.Open strSQL, ObjConn

      IF NOT objRS.EOF THEN 
           strName = objRS("strName")
           strAddress = objRS("strAddress")
           strCity = objRS("strCity")
           strState = objRS("strstate")
           strZip = objRS("strZip")
%>

<center>
        <font face='copperplate gothic Bold' color='#FF0000'>Order Information<br></font>
        <br>
        <font face='copperplate gothic Bold' color='#FF0000'>Returning Customer - Please confirm your mailing and e-mail addresses.</font><br><br>
<% ELSE %>
    <center>
        <font face='copperplate gothic Bold' color='#FF0000'>Order Information<br></font>
        <br>
        <font face='copperplate gothic Bold' color='#FF0000'>New Customer - Please provide your shipping address.</font><br><br>
<% END IF %>
        <div align="center">
            <center>
                <table border="0" cellpadding="0" cellspacing="0" width="697">
                    <tr>
                        <td valign="top" width="93"><font face="copperplate gothic Bold">Ship to:</font></td>
                        <td width="600">
                            <table border="0" cellpadding="0" cellspacing="0" width="600">
                                <tr>
                                    <td width="214">Name:</td>
                                    <td><form method="GET" action="checkout03.asp">
                                                   <input type="text" name="Name" size="21" value="<% =strName %>"></td>
                                </tr>
                                <tr>
                                    <td width="214"> Street address:</td>
                                    <td width="844"><input type="text" name="Street" size="21" value="<% =strAddress %>"></td>
                                </tr>
                                <tr>
                                    <td width="214">City:</td>
                                    <td width="844"><input type="text" name="City" size="21" value="<% =strCity %>"></td>
                                </tr>
                                <tr>
                                    <td width="214">State:</td>
                                    <td width="844"><input type="text" name="State" size="6" value="<% =strState %>"></td>
                                </tr>
                                <tr>
                                    <td width="214">Zip:</td>
                                    <td width="844"><input type="text" name="Zip" size="10" value="<% =strZip %>"></td>
                                </tr>
                                <tr>
                                    <td width="214" valign="top">E-mail:</td>
                                    <td width="844"><input type="text" name="NewEmail" size="36" value="<% =strEmail%>">
                                                    <input type="hidden" name="email" value="<% =strEmail%>">
                                        <br>
                                        <br>
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>
                    <tr>
                        <td valign="top" width="93">
                            <font face="copperplate gothic Bold">Items:</font>
                        </td>
                        <td>
                            

<% Dim ISBN, strSQLOrder, iItemcount, icount
                          iItemCount = request.cookies("shoppingcart")("itemcount")

               'List each item in cart
               For iCount = 1 to iItemCount

                   ISBN = Request.Cookies("ShoppingCart")("ISBN"&iCount)

        strSQLOrder= "SELECT tblBookDescription.ISBN, tblBookDescription.strTitle, tblBookDescription.strDescription, tblBookDescription.dblPrice, tblBookDescription.strPublisher " & _
                     "FROM tblBookDescription " & _
                     "WHERE (((tblBookDescription.ISBN)= '" & ISBN & "')); "
                     
                         set objRS = server.createobject("ADODB.recordset")
                             ObjRS.open strSQLOrder, ObjConn    


%>
                                <table>
                                <tr>
                                    <td valign="top" width="20"><% =icount %>)</td>
                                    <td valign="top" width="350">
                                        <font color="#5B78A4"><% =objRS("strTitle") %></font>
                                        <br>
                                        <font size="-1">by <% =funListAuthors(objRS("ISBN"))%></font>
                                    </td>
                                    <td valign="top">
                                        Price: <% =formatcurrency(objrs("dblPrice")*.8) %>
                                    </td>
                                </tr>
                                 
                            </table>
<% NEXT %>

                            <br>
                    </tr>

<% Dim strGrandTotal 
       strGrandTotal = request.cookies("ShoppingCart")("strGrandTotal") 
%>

                    <tr>
                        <td valign="top" width="93">
                            <font face="copperplate gothic Bold">Total:</font></td>
                        <td width="600"><font face="arial,verdana,helvetica"><% =formatcurrency(strGrandTotal) %></font></td>
                    </tr>
                </table>
            </center>
            <p align="center">
                <input type="image" src="/06Spring/sullivs8/Bookstore/images/placeorder.png">
                </form>
            </p>
            
            <p align='center'>
                <a href="OrderHistory.asp?email=<% =stremail%> ">View Order History for <% =stremail%></a>
</body>
</html>
