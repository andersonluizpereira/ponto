<!-- #include file = "../includes/Request.asp" -->


<%

dim itemSession

for each itemSession in Session.Contents
         
'limpando a session.
Session(itemSession) = empty
            
next

Response.Redirect getBaseLink("/login/login.asp")

%>
