<html>

<head>
<title>BamBooks.com -- Site Features</title>
</head>

<body>

 
 
<!--#include file="includes/header.html"-->


  <center>
  <table border="0" cellpadding="5" cellspacing="0" width="800">
    <tr>
      <td valign="top" >  
  
<!--#include file="includes/menu.html"-->
 
  </td>
      <td><font face="copperplate gothic Bold" color="#FF0000">Site Features</font>
        <ul>
          <li>Site created by Stephanie Sullivan as a class project for <a href="http://yorktown.cbe.wwu.edu/sandvig/mis314/">MIS
            314</a> at Western Washington University. </li>
          <li>All product information is dynamically generated using Active
            Server Pages.</li>
          <li>Book, customer and order information is stored in an Access database.</li>
          <li>Server-side includes are used for all components that are used more
            than once (such as the search/browse menu, ListAuthor function, 
            header and footer). </li>
          <br><br>
          <li><b>Home Page</b>
            <ul>
              <li>&quot;Today's Features&quot; are randomly selected from the
            database.</li>
              <li>The browse menu is dynamically generated from the database using a SQL query that shows
            only the current book categories.</li>
              <li>Book descriptions are truncated at 300 characters.</li>
            </ul>
          <li><b>Search/Browse response page</b>
                <ul>
                  <li>The search function searches book titles, descriptions and
                    categories fields in the database.</li>
                  <li>The RecordCount property of the RecordSet object is used
                    to count the number of books found by the search.</li>
                  <li>Searches that have no matches respond gracefully</li>
                </ul>
          </li>
          <li><b>Shopping cart page</b>
                <ul>
                  <li>Uses a cookie to store the ISBNs of items in the
                    cart.</li>
                </ul>
          </li>
          <li><b>Checkout pages</b>
                <ul>
                  <li>Searches the database for email addresses of existing
                    customer accounts and writes their shipping information in
                    the form on the order confirmation page.</li>
                </ul>
          </li>
          <li><b>Order Confirmation Page</b>
                <ul>
                  <li>Checks for shopping cart and prompts user if cart is
                    empty.</li>
                  <li>All fields are checked to make sure that they contain
                    information.</li>
                  <li>Checks email address in database and prompts user to try
                    again user if address not found.</li>
                  <li>Modifications made to customer information are updated in
                    the database.</li>
                  <li>Order information are written to the database.</li>
                  <li>An email is sent to the customer with the order
                    information.</li>
                    <li>The shopping cart is emptied by setting ItemCount to zero in the ShoppingCart cookie. .</li>

                </ul>
          </li>
          <li><b>Order History Page</b>
                <ul>
                  <li>Searches the database for all orders associated with
                    e-mail address</li>
                  <li>If no matching email address is found user is prompted to
                    try again.</li>
                </ul>
                </li>
          <li><b>Enhancements</b>
                <ul>
                  <li><b>Update Graphical Appearance:</b> Logo, buttons and Search/Browse menu images are created by me.</li>
                  <li><b>Bomb Proofing:</b> " ' " is now equal to " '' "</li>
                  <li><b>Shopping Cart Status box:</b> Located right below the Search/Browse menu.</li>
                  <li><b>Product Page Redirect:</b> If only one item is returned in the search/browse menu then it goes straight to the product page.</li>
                </ul>
                <p>&nbsp;</li>
          <li>Thanks to Amazon.com for the use of its book images and book descriptions.</li>
      </td>
    </tr>
    <tr>
      <td></td>
      <td></td>
    </tr>
  </table>
  </center>
<br>
  
<!--#include file="includes/footer.html"-->  
 

</body>

</html>
