<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Nebraska Medicaid Planning: Submission Confirmation</title>
<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
}
.style5 {
	color: #000000;
	font-size: 12px;
}
-->
</style>
<script type="text/JavaScript">
<!--
function MM_preloadImages() { //v3.0
  var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
    var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
    if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}
//-->
</script>
</head>

<body>
<table width="614" height="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td width="614"><table width="600" height="100%" border="0" cellpadding="0" cellspacing="0">
      <tr>
        <td width="600" height="155%" valign="top"><body bgproperties="fixed">
  <p align="center"><%@ LANGUAGE="VBSCRIPT" %>
    
    
  <%

' Description Transfer


%>
  <p align="center">Nebraska Medicaid Planning</p>
  <p align="center">Thank you for your submission. We will contact you shortly. </p>
  <p align="center">You can now close this window. </p>
  <p><%



' Set up variables to hold the text strings

Dim DestinationEmail

Dim OriginatingEmail

Dim Subject

Dim Body




' assign text to the variable

OriginatingEmail = "webform@nebraskamedicaidplanning.com"

DestinationEmail = "dan@dunn-stockmannlaw.com"

Subject = "Nebraska Medicade Planning Dot Com"



' The body contains all the form data separated by semicolons

Body = "Name = " & Request.form("LName") & ", " & Request.form("FName") & vbcrlf _

& "Relationship to person in Nursing Home = " & Request.form("relation") &  vbcrlf _

& "City = " & Request.form("City") & vbcrlf _

&  "State = " & Request.form("State") & vbcrlf _

& "Phone = " & Request.form("Phone") & vbcrlf _

& "EMail Address = " & Request.form("EMail") & vbcrlf _

& "Type Person = " & Request.form("type") & vbcrlf _

& "Home Owner = " & Request.form("homeowner")& vbcrlf _

& "Assets = " & Request.form("assets") & vbcrlf _

& "Comments and Questions = " & Request.form("cq") 



Set Mailer = Server.CreateObject("Persits.MailSender")

Mailer.From = OriginatingEmail

Mailer.Host = "mail1.omni-tech.net"

Mailer.AddAddress DestinationEmail

Mailer.Subject = Subject

Mailer.Body = Body

Mailer.Send

If Err <> 0 Then

  Response.Write "Mail send failure. Error was " & Err.Description

end if



%> </p>
  </p></td>
        </tr>
    </table></td>
  </tr>
  
  <tr>
    <td height="5" bgcolor="#FFFFFF"><div align="right" class="style5"> Copyright Nebraskamedicaidplanning.com</div></td>
  </tr>
  <tr>
    <td>&nbsp;</td>
  </tr>
</table>
</body>
</html>
