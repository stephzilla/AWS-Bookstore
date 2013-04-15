<%
   'List Authors Function:
   ' This function requires ISBN as an input and returns a string of author names formatted as 
   ' hyperlinks. To use the function:
   ' 1. Include this function in your page using <!--#include file="ListAuthors.asp"-->
   '    (do this only once per page)
   ' 2. To list a book's authors, call the function and pass it the ISBN of the book using:
   '    funListAuthors(ISBN)
   '    where ISBN is a variable containing the book's ISBN.
   '    If you are retrieving the ISBN directly from a database the syntax will look like:
   '    funListAuthors(objRS("ISBN"))
   ' It requires that you have an open database connection object named "objConn"
   ' You are welcome to modify this script.

   function funListAuthors(ISBN2) 
        dim objAuthors, sSQL, strAuthorList

        sSQL = "SELECT tblAuthors.strLastName, tblAuthors.strFirstName "&_
               "FROM tblAuthors INNER JOIN tblAuthorsBooks ON tblAuthors.AuthorID = tblAuthorsBooks.AuthorID "&_
               "WHERE (((tblAuthorsBooks.ISBN)='"& isbn2 &"')) "&_
               "ORDER BY tblAuthors.strLastName;"

        'Create the recordset object for book authors
        Set objAuthors = server.CreateObject("ADODB.Recordset")
        objAuthors.open sSQL, objConn

        do while not objAuthors.EOF
           strAuthorList = strAuthorList & "<a href='SearchBrowse.asp?action=search&strSearch=" & _ 
                           objAuthors("strLastName") & "'>" & objAuthors("strFirstName") & _
                           " " & objAuthors("strLastName") & "</a>"
           objAuthors.MoveNext
           if not objAuthors.EOF then strAuthorList = strAuthorList & ", "
        Loop 

        objAuthors.close
        Set objAuthors = Nothing

        'return the author list
        funListAuthors = strAuthorList
   End Function
%>
