<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"  "http://www.w3.org/TR/html4/loose.dtd">

<%@ Page language="VB"
         title="Kaртка інінціатора"
         MaintainScrollPositionOnPostback=true
         Culture="en-Us"
         UICulture="En"
%>

<%@ Import Namespace = "System.Data.Odbc" %>
<%@ Import Namespace = "System"  %>
<%@ Import Namespace = "System.IO"  %>
<%@ Import Namespace = "System.Data"  %>
<%@ Import Namespace = "System.IO.DirectoryInfo"  %>
<%@ Import Namespace = "System.IO.Directory"  %>
<%@ Import Namespace = "System.IO.File"  %>
<%@ Import Namespace = "System.Web.HttpServerUtility"  %>
<%@ Import Namespace = "System.Data"  %>
<%@ Import Namespace = "System.Web"  %>
<%@ Import Namespace = "System.Web.Sessionstate"  %>
<%@ Import Namespace = "System.Web.Configuration"  %>
<%@ Import Namespace = "System.Web.HttpApplication"  %>
<%@ Import Namespace = "System.Data.SqlClient" %>
<%@ Import Namespace = "System.Collections.Specialized" %>
<%@ Import Namespace = "System.Configuration" %>
<%@ Import Namespace = "System.Collections.ObjectModel" %>
<%@ Import Namespace = "System.Collections" %>
<%@ Import Namespace = "System.Text" %>
<%@ Import Namespace = "System.Web.UI.Page" %>
<%@ Import Namespace = "System.Web.HttpRequest" %>
<%@ Import Namespace = "System.Threading" %>
<%@ Import Namespace = "System.Globalization" %>
<%@ Import Namespace = "Microsoft.SharePoint.WebControls"  %>
<%@ Import Namespace = "System.Data.SqlClient" %>

<HTML xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office">

<HEAD>
<META name="WebPartPageExpansion" content="full">
<script type="text/javascript" >

/*
* Browser.js простейший анализатор клиента
*/

function poss() {

  var browser={version:     parseInt(navigator.appVersion),
               isNetscape:  navigator.appName.indexOf("Netscape") !=-1,
                isMicrosoft: navigator.appName.indexOf("Microsoft") !=-1,
               author:      "Savchenko Arthur",
               version:     "001.01 version 01.02",
               descript:    "MDE Programm Navigation",
               url:         navigator.appName,
               loc:         navigator.Location,
               agent:       navigator.userAgent,
               online:      navigator.onLine,
               codeapp:     navigator.appCodeName,
               iscookie:    navigator.cookieEnabled,
               platform:    navigator.platform,
               host:        location.host,
               port:        location.port,
               path:        location.pathname ,
               search:      location.search
               };

        alert(browser.host);

     //alert(document.getElementById('dd0')[0].attributes);

}

</script>

<style type="text/css">p{color:#2a2a2a;margin-top:0;margin-bottom:0;padding-bottom:15px;line-height:18px;}.topic a:link{text-decoration:none;color:#1364c4;}.topic a:link{text-decoration:none;color:#00709f;}.topic a{text-decoration:none;color:#1364c4;}.topic a{text-decoration:none;color:#00709f;}a:link{text-decoration:none;}div.alert p{margin:0;}table p{padding-bottom:0;}</style>

</HEAD>

<BODY>
  <form id="form1" runat="server">
  <input name="Button1" type="button" value="button" onclick="poss()" />
  <input id="dd0" type="text"/>

<hr>
<%
Dim ip As String
ip=request.ServerVariables("REMOTE_ADDR")

'if ip<>"194.248.333.500" then
'  response.Status="401 Unauthorized"
'  response.Write(response.Status)
'  response.End
'end if

response.Write("<br>")
response.Write(response.Status & "<br>")
response.Write(ip)
%>
<hr>
<span >Просомтр коннекта</span>
<%
If response.IsClientConnected=true then
   response.write("The user is still connected!")
else
   response.write("The user is not connected!")
end if
%>

<hr>
<% response.write(now().tostring("dd MM - yyyy")) %>

<hr>
<% response.write(FormatDateTime(now(),vbshortdate)) %>
<hr>
<%  Response.Write(Request.QueryString("id")) %>
<hr>
<%
    dim numvisits
    response.cookies("NumVisits").Expires = DateTime.Now.AddYears(1)
    numvisits=request.cookies("NumVisits")

    if request.Querystring("c")=1 then
       sss_company.SelectCommand    = "SELECT * FROM [s_company] where id= " & request.Querystring("c")
    else
       sss_company.SelectCommand    = "SELECT * FROM [s_company] where id= 3"
    end if


       if request.Querystring("c")="" then
          viewstate("dd")=22
          sss_company.SelectCommand    = "SELECT * FROM [s_company]"
       else
          viewstate("dd")=23
          sss_company.SelectCommand = "SELECT * FROM [s_company] where id= " & request.Querystring("c")
          sss_comp.SelectCommand    = "SELECT * FROM [s_company] where id = 3"
          scrollyes.visible="false"
       end if
%>

<%  =viewstate("dd") %>

     <!--Company-->
          <asp:SqlDataSource id          = "sss_company"
                        runat            = "server"
                        ConnectionString = "<%$ Resources:Art_Global,Aps_connection %>"
                        ProviderName     = "<%$ Resources:Art_Global,Aps_proviser %>"

     SelectCommand    = "SELECT * FROM [s_company] where id= 1"/>

     <!--Company-->
          <asp:SqlDataSource id          = "sss_comp"
                        runat            = "server"
                        ConnectionString = "<%$ Resources:Art_Global,Aps_connection %>"
                        ProviderName     = "<%$ Resources:Art_Global,Aps_proviser %>"

     SelectCommand    = "SELECT * FROM [s_company] where id= 1"/>

<br>
<asp:GridView runat="server" id="GridView1" DataSourceID="sss_company">                    </asp:GridView>
<br>
<asp:GridView runat="server" id="GridView2" DataSourceID="sss_comp">                    </asp:GridView>
</form>

<div runat="server" id="scrollyes">
   <p>Внимание щшибка.</p>
</div>
<div class="alert">
</div>
</BODY>
</HTML>
