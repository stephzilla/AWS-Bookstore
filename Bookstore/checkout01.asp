<% Option Explicit %>
<!--#include virtual="/06Spring/sullivs8/database/adovbs.inc"-->
<!--#include file="DatabaseConnect.asp"-->

<html>

<head>
<title>BamBooks.com -- Checkout</title>
</head>

<body>

<% dim iItemCount
    iItemcount= request.cookies("ShoppingCart")("ItemCount") 

%>
<!--#include file="includes/header.html"-->



<div align="center">
  <font face="copperplate gothic Bold" color="#FF0000">Your Account</font><center><br>

       <font face="copperplate gothic Bold" size="2" color="#000000">You have 
       <%= iItemcount %>
       items in your cart.<br>
      
       </font><br>

  </div>
</center>
<div align="center">
  <center>
  <table border="0" cellpadding="0" cellspacing="0" width="600">
    <tr>
      <td>
      
        <p align="center"><font face="copperplate gothic light" color="#000000">Buying
        on-line is quick and easy!</font></p>
        <p>&nbsp;&nbsp;
        
      </td>
    </tr>
    

    <tr>
      <td>
      
        <div align="center">
          <table border="0" cellpadding="0" cellspacing="0" width="560" >
            <tr>
              <td width="6">&nbsp;</td>
              <td width="197"><font face="Times new Roman">Please enter your e-mail:</font>
              </td>
              <td width="351"> 
                <form method="GET" action="checkout02.asp">
                  <input type="text" name="email" size="36">
                  <input type="submit" value="Submit">
              </td>
              </form>
            </tr>
          </table>
        </div>
        
      </td>
    </tr>
    

  </table>
  <br><br>
      <br>
 
  </center>
</div>

 
</body>
</html>
