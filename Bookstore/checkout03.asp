<% option explicit %>
<!--#include file="DatabaseConnect.asp"-->
<!--#include virtual="/06spring/sullivs8/database/adovbs.inc"-->
<!--#include file="ListAuthors.asp"-->

<html>
<head><title>BamBooks.com -- Checkout</title></head>
<body>
<center>     
<!--#include file="includes/Header.html"-->
<% dim objRS, strSQL, strEmail, strName, strAddress, strCity, strState, strZip, itemcount, i
       strName = request.querystring("Name")
       strAddress = request.querystring("Street")
       strCity = request.querystring("City")
       strState = request.querystring("State")
       strZip = request.querystring("Zip")
       strEmail = request.querystring("NewEmail")
IF len(strName) < 5 OR _
   len(strEmail) < 6 OR _
   len(strAddress) < 6 OR _
   len(strCity) < 2 OR _
   len(strState) < 2 OR _
   len(strZip) < 5 THEN

    response.write("Please complete all fields." & "<br><br><INPUT TYPE='button' VALUE='< Go Back'onClick='history.back()'><br>")
    response.end
END IF



          StrSQL = "SELECT tblCustomers.strName, tblCustomers.strAddress, tblCustomers.strCity, tblCustomers.strState, tblCustomers.strZip, tblCustomers.strEmail " & _
                   "FROM tblCustomers " & _
                   "WHERE (((tblCustomers.strEmail)='" & strEmail & "'));"

                      Set objRS = Server.CreateObject("ADODB.Recordset")
                          objRS.Open strSQL, ObjConn,, 3

      IF objRS.EOF THEN
         objrs.AddNew
      END IF
 
           objRS("strName") = strName 
           objRS("strAddress") = strAddress
           objRS("strCity") = strCity
           objRS("strstate") = strState
           objRS("strZip") = strZip
           objRS("strEmail") = request.querystring("Newemail")
           objRS.Update
%>


        <font face='copperplate gothic Bold' color='#FF0000'>Order Confirmation</font><br>
        <br>
        
        <div align="center">
            <center>
                <table border="0" cellpadding="0" cellspacing="0" width="600">
                    <tr>
                        <td width="159" valign="top"><font face="copperplate gothic light">Books Shipped:</font></td>
                        <td width="437">

<% Dim iItemCount, strSQLOrder, strBookTitle
       iItemCount = request.cookies("shoppingcart")("itemCount")

          StrSQL = "SELECT tblOrders.dtOrderDate, tblOrders.strEmail, tblOrders.ISBN " & _
                   "FROM tblOrders;"

                      Set objRS = Server.CreateObject("ADODB.Recordset")
                          objRS.Open strSQL, ObjConn,, 3

   FOR i = 1 to iItemCount
       objRS.AddNew
       objRS("strEmail") = strEmail
       objRS("ISBN") = request.cookies("shoppingcart")("ISBN" & i)
       objRS("dtOrderDate") = now()
       objRS.Update
   NEXT

   objRS.Close

   FOR i = 1 to iItemCount

        strSQLOrder= "SELECT tblBookDescription.ISBN, tblBookDescription.strTitle, tblBookDescription.strDescription, tblBookDescription.dblPrice, tblBookDescription.strPublisher " & _
                     "FROM tblBookDescription " & _
                     "WHERE (((tblBookDescription.ISBN)= '" & request.cookies("ShoppingCart")("ISBN" & i) & "')); "

         objRS.Open strSQLOrder, objConn

         strBookTitle = (strBookTitle & i & ". " & objRS("strTitle") & "<br>" )
   
         objRS.Close
   NEXT


%>
                            <% =strBookTitle %><br></td>
                    </tr>
                    <tr>
                        <td width="159" valign="top"><font face="copperplate gothic light">Shipping Address:</font></td>
                        <td width="437">
                            <% =strName %><br><% =strAddress %><br><% =strCity %>, <% =strState %> &nbsp; <% =strZip %><br><br>
                        </td>
                    </tr>
                    <tr>
                        <td width="159" valign="top">
<% Dim strFinalTotal 
       strFinalTotal = request.cookies("ShoppingCart")("strGrandTotal") 

%>
                            <font face="copperplate gothic light">Total:</font></td>
                        <td width="437"><% =formatcurrency(strFinalTotal) %>
                </table>
                <br>
                <font face="copperplate gothic Bold">A confirmation has been sent to your email address.<br>Thank you for shopping with GeekBooks.com.</font>
                <br>
                <br>
                <br>
                <a href="default.asp">
                    <img border="0" src="/06Spring/sullivs8/bookstore/images/continue.png"></a><br>
                <br>
                <br>
                <br>
                <p align='center'>
                    <a href="OrderHistory.asp?email=<% =strEmail %>">View Order History for <% =strEmail %></a>
                </center>
        </div>
<%

Dim strHTMLMessage
strHTMLMessage =("<a href='http://yorktown.cbe.wwu.edu/06Spring/sullivs8/bookstore/default.asp'> " &_
                 "   <img src='http://yorktown.cbe.wwu.edu/06Spring/sullivs8/bookstore/images/bambooks.png' " &_
                 "   border='0' alt='BamBooks Logo'>" &_
                 "</a><br>" &_
                 "<h3 align='center'>Order Confirmation</h3>" & _
                 "<i><tr><td width='159' valign='top'><font face='Comic Sans MS'>Shipping Address:<br></font></td><td width='437'>" &_
                strName & "<br>" & strAddress & "<br>" & strCity & "," & strState & "&nbsp;" & strZip & "<br><br>" &_
                 "</td></tr><tr><td width='159' valign='top'>" &_
                 "<font face='Comic Sans MS'>Total:</font></td>" &_
                 "<td width='437'>" &_
                formatcurrency(strFinalTotal) &_
                 "</table><br>" &_ 
                 "<font face='Comic Sans MS'>Thank you for shopping with BamBooks.com.</font><br><br><br>")

'Create the mail object 
Dim objMail
Set objMail      = Server.CreateObject("CDO.Message")

'Set mail object properties 
objMail.To       = request.querystring("email")
objMail.From     = "sullivs8@cc.wwu.edu"
objMail.Subject  = "Order Confirmation from GeekBooks"

objMail.HTMLBody = strHTMLMessage 'your HTML email message

'Optional properties
'objMail.TextBody = "Your email message in plain text for email clients that do not support HTML"
'objMail.CC       = "csandvig@wwu.edu"
'objMail.BCC      = "chris.sandvig@wwu.edu"

'Optional importance
'objMail.Fields("urn:schemas:httpmail:importance").Value = 2
'objMail.Fields.Update()

'Send the email 
objMail.Send 

'Clean-up 
Set objMail = Nothing      

Response.Cookies("ShoppingCart").expires = date -1

%>
    </center>
</body>
</html>
