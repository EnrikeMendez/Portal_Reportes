posYmenu = 0;
bgcolor='#006699';
bgcolor2='#FFCC00';
needcenter =false;
if(!needcenter)
	posXmenu = 170;
else
	{	if (document.all)
		posXmenu = (document.body.clientWidth/2)-(200/2);
	else
		posXmenu = (window.innerWidth/2)-(200/2); 
	}
document.write('<style type="text/css">');
document.write('.popper { POSITION: absolute; VISIBILITY: hidden; z-index:100; }')
document.write('#topgauche { position:absolute; top:'+posYmenu+'px; left:'+posXmenu+'px; z-index:110; }')
document.write('A:hover.ejsmenu {color:#FFFFFF; text-decoration:none;}')
document.write('A.ejsmenu {color:#FFFFFF; text-decoration:none;}')
document.write('</style>')
document.write('<DIV class=popper id=topdeck></DIV>');
/*
SCRIPT EDITE SUR L'EDITEUR JAVACSRIPT
http://www.editeurjavascript.com
*/

/*
LIENS
*/
zlien = new Array;
zlien[0] = new Array;
zlien[1] = new Array;
zlien[0][0] = '<A HREF="smqry.asp" CLASS=ejsmenu>Simple Query</A>';
zlien[1][0] = '<A HREF="state-account.asp" CLASS=ejsmenu>Statement of Account</A>';
zlien[1][1] = '<A HREF="payment-search.asp?concept=I01" CLASS=ejsmenu>Advanced Payments</A>';
zlien[1][2] = '<A HREF="payment-search.asp?concept=I02" CLASS=ejsmenu>Invoice Payments</A>';
zlien[1][3] = '<A HREF="payment-search.asp" CLASS=ejsmenu>General Payments</A>';
zlien[1][4] = '<A HREF="invoice-search.asp?inv_type=a" CLASS=ejsmenu>American Invoices</A>';
zlien[1][5] = '<A HREF="invoice-search.asp?inv_type=m" CLASS=ejsmenu>Mexican Invoices</A>';
var nava = (document.layers);
var dom = (document.getElementById);
var iex = (document.all);
if (nava) { skn = document.topdeck }
else if (dom) { skn = document.getElementById("topdeck").style }
else if (iex) { skn = topdeck.style }
skn.top = posYmenu+24;

function pop(msg,pos)
{
skn.visibility = "hidden";
a=true
skn.left = posXmenu+pos;
var content ="<TABLE BORDER=0 CELLPADDING=0 CELLSPACING=0 BGCOLOR=#000000 WIDTH=150><TR><TD><TABLE WIDTH=100% BORDER=0 CELLPADDING=0 CELLSPACING=1>";
pass = 0
while (pass < msg.length)
	{
	content += "<TR><TD BGCOLOR="+bgcolor+" onMouseOver=\"this.style.background='"+bgcolor2+"'\" onMouseOut=\"this.style.background='"+bgcolor+"'\" HEIGHT=20><FONT SIZE=1 FACE=\"Arial\">&nbsp;&nbsp;"+msg[pass]+"</FONT></TD></TR>";
	pass++;
	}
content += "</TABLE></TD></TR></TABLE>";
if (nava)
  {
    skn.document.write(content);
	  skn.document.close();
	  skn.visibility = "visible";
  }
    else if (dom)
  {
	  document.getElementById("topdeck").innerHTML = content;
	  skn.visibility = "visible";
  }
    else if (iex)
  {
	  document.all("topdeck").innerHTML = content;
	  skn.visibility = "visible";
  }
}
function kill()
{
	skn.visibility = "hidden";
}
document.onclick = kill;
document.write('<DIV ID=topgauche><TABLE BORDER=0 CELLPADDING=0 CELLSPACING=0 BGCOLOR=#000000 WIDTH=300><TR><TD><TABLE CELLPADING=0 CELLSPACING=1 BORDER=0 WIDTH=100% HEIGHT=25><TR>')
document.write('<TD WIDTH=150 ALIGN=center BGCOLOR='+bgcolor+' onMouseOver="this.style.background=\''+bgcolor2+'\';pop(zlien[0],0)" onMouseOut="this.style.background=\''+bgcolor+'\'"><A onClick="return(false)" onMouseOver="pop(zlien[0],0)" href=# CLASS=ejsmenu><FONT SIZE=2 FACE="Arial"><b>Simples Query</b></FONT></a></TD>')
document.write('<TD WIDTH=150 ALIGN=center BGCOLOR='+bgcolor+' onMouseOver="this.style.background=\''+bgcolor2+'\';pop(zlien[1],150)" onMouseOut="this.style.background=\''+bgcolor+'\'"><A onClick="return(false)" onMouseOver="pop(zlien[1],100)" href=# CLASS=ejsmenu><FONT SIZE=2 FACE="Arial"><b>Financial</b></FONT></a></TD>')
document.write('</TR></TABLE></TD></TR></TABLE></DIV>')