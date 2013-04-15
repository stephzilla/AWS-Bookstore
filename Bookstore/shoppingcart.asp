<% Option Explicit %>
<!--#include virtual="/06Spring/sullivs8/database/adovbs.inc"-->
<!--#include file="DatabaseConnect.asp"-->
<!--#include file="ListAuthors.asp"-->

<html>
<head>
   <title>BamBooks.com -- Shopping cart</title>
</head>
<body>
<!--#include file="includes/header.html"-->
<%
'This shopping cart uses a cookie named "ShoppingCart" 
'to store the ISBNs of the items in the cart.
'The cookie puts each ISBN value into a separate key named ISBN1, ISBN2, etc.
'The current number of items in the cart is stored in a key named "ItemCount"

dim ISBN, iItemCount, iItem, iCount, bFound
ISBN = trim(Request.QueryString("isbn"))

'Initialize iItemCount with number of books in the cart
If request.cookies("ShoppingCart").HasKeys then
     iItemCount = request.cookies("ShoppingCart")("ItemCount")
else
     'No cookie yet
     iItemCount = 0
end if

'Add items to cart
If Request("action") = "add" and len(ISBN)>5 then
    'add new item to ShoppingCart
     iItemCount = iItemCount + 1
     Response.cookies("ShoppingCart")("ISBN"&iItemCount) = ISBN
end if

'Delete items from Cart
If Request.QueryString("action") = "del" and len(ISBN)>5 then
     'find the isbn
     bFound = "False"
     iItem = 1
     Do while NOT bFound and iItem <= iItemCount
          If Request.Cookies("ShoppingCart")("ISBN"&iItem) = ISBN then
               bFound = "True"
          else
               iItem = iItem + 1
          End If
     Loop
     
     If bFound then
         'replace the deleted item with the item above it
          iItemCount=iItemCount - 1
          For iCount = iItem to iItemCount
               Response.Cookies("ShoppingCart")("ISBN" &iCount)= _ 
                Request.Cookies("ShoppingCart")("ISBN"&iCount+1)
          Next
     else
          Response.write "Error: Could not match ISBN to delete."
     end if
End If
     
'Update iItemCount
Response.Cookies("ShoppingCart")("ItemCount")=iItemCount
Response.Cookies("ShoppingCart").Expires = Date + 30

' ***********************************************************************************
' You do not need to modify anything above this line.
' You will need to modify the table below to list the items
' in your cart, calculate, total costs, etc.
' ***********************************************************************************

Dim strSQL, strTitle, dblPrice, strSubTotal, objRS
 
    
  


' list items in shopping cart
If iItemCount = 0 then
     response.write "<center><font face='copperplate gothic light' color='#FF0000'>" & _
                    "Your Shopping Cart is empty.</font><br><br>"
Else
%>

  
     <div align="center">
      <center>
          <font face='Comic Sans MS' color='#FF0000'>You have <% =iItemCount %>
               book<% If iItemCount > 1 then response.write "s" %> in your shopping cart.
          </font><br><br>

          <table cellpadding="4">
             <tr>
                <td>Item</td>
                <td>ISBN</td>
                <td>Price</td>
                <td>  </td>
             </tr>
<% 

               'List each item in cart
               FOR iCount = 1 to iItemCount
                     ISBN = Request.Cookies("ShoppingCart")("ISBN" & iCount)
       

        strSQL= "SELECT tblBookDescription.ISBN, tblBookDescription.strTitle, tblBookDescription.strDescription, tblBookDescription.dblPrice , tblBookDescription.strPublisher " & _
                "FROM tblBookDescription " & _
                "WHERE (((tblBookDescription.ISBN)= '" & ISBN & "')); "

              Set objRS=Server.CreateObject("ADODB.Recordset")
                  objRS.Open strSQL, ObjConn, 1 

              

%>

                     <tr>
                          <td><% =iCount %>.</td>
                          <td><a href="ProductPage.asp?ISBN=<% =objrs("ISBN") %>"><% =objRS ("strTitle") %></a>
                              <br>by <% =funlistAuthors(ISBN) %></td>
                          <td>List Price:<% =FormatCurrency(objRS("dblPrice")) %><br>
                              Our Price: <% =FormatCurrency((objRS("dblPrice"))*.8) %><br>
                              You save: <% =FormatCurrency((objRS("dblPrice"))-((objRS("dblPrice"))*.8)) %>
                              </td>
                          <td><a href="ShoppingCart.asp?ISBN=<% =ISBN %>&action=del">Remove</a></td>
                     </tr>
                    
            <% strSubTotal = ((objRS("dblPrice"))*.8) + strSubTotal

NEXT 
                 IF icount >0 THEN
                Dim strShipping, strGrandTotal
                    strShipping = (((icount-2)*.5)+ 3.00)
                    strGrandTotal = (strSubTotal + strShipping)
                END IF
  %> 
       <tr>
          <td></td>
          <td></td>
          <td>
             Totals:<br>
             Subtotal: <% =FormatCurrency(strSubTotal)  %> <br>
             Shipping: <% =FormatCurrency(strShipping)  %> <br>
             Total: <% =FormatCurrency(strGrandTotal)  %> 
          </td>
       </tr>
            
 <% response.cookies("shoppingcart")("strGrandTotal") = strGrandTotal %>                    
       </table>
     </center>
    <br>
    <div align="center">
        <center>
            <table border="0" cellpadding="5" cellspacing="0" width="400">
                <tr>
                    <td>
                        <a href="default.asp">
                            <img border="0" src="/06Spring/sullivs8/bookstore/images/continue.png" width="150" height="21"></a><br>
                    </td>
                    <td>
                        <p align="right">
                            <a href="checkout01.asp">
                                <img border="0" src="/06Spring/sullivs8/bookstore/images/checkout.png" width="150" height="21"></a>
                    </td>
                </tr>
            </table>
            <table border="0" cellpadding="5" cellspacing="0" width="400">
                <tr>
                    <td>
                        <br>
                        <p align="center">
                            Shipping is $3.00 for the first book and $.50 for each additional book. To assure
                        reliable delivery and to keep your costs low we send all books via UPS ground.
                    </td>
                </tr>
            </table>
    </div>
<% 

END IF

   objRS.close
   set objRS = nothing
   objConn.close
   set objConn = nothing
%> 
</body>
</html>
