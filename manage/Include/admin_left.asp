<style>
body{background-image: url(../Images/left_bg.gif);}
table  { border:0px; }
td  { font:normal 12px ����; }
img  { vertical-align:bottom; border:0px; }
a  { font:normal 12px ����; color:#000000; text-decoration:none; }
a:hover  { color:#428EFF;text-decoration:underline; }
.sec_menu  { border-left:1px solid white; border-right:1px solid white; border-bottom:1px solid white; overflow:hidden; background:#D6DFF7; }
.menu_title  { }
.menu_title span  { position:relative; top:2px; left:8px; color:#215DC6; font-weight:bold; }
.menu_title2  { }
.menu_title2 span  { position:relative; top:2px; left:8px; color:#428EFF; font-weight:bold; }
</style>
<script language=javascript>
function showsubmenu(sid)
{
whichEl = eval("submenu" + sid);
imgmenu = eval("imgmenu" + sid);
if (whichEl.style.display == "none")
{
eval("submenu" + sid + ".style.display=\"\";");
imgmenu.background="../images/menuup.gif";
}
else
{
eval("submenu" + sid + ".style.display=\"none\";");
imgmenu.background="../images/menudown.gif";
}
}

top.document.title="��̨����"; 

//Obj�����Ķ���,ButtonSymbol��ť״̬,MainUrl--MainFrame������
function SelectObjItem(Obj,ButtonSymbol,MainUrl)
{  
    if(MainUrl!='')
	{window.parent.parent.frames['main'].location.href=MainUrl;
	}

}
//����Ҽ����Խ�ֹ��
if (window.Event) 
document.captureEvents(Event.MOUSEUP); 
function nocontextmenu() 
{
event.cancelBubble = true
event.returnValue = false;
return false;
}
function norightclick(e) 
{
if (window.Event) 
{
if (e.which == 2 || e.which == 3)
return false;
}
else
if (event.button == 2 || event.button == 3)
{
event.cancelBubble = true
event.returnValue = false;
return false;
}
}
document.oncontextmenu = nocontextmenu; // for IE5+
document.onmousedown = norightclick; // for all others

</script>
<!--#include file="Check_Sql.asp"-->
<!--#include file="../../include/connect.asp "-->
<table cellspacing="0" cellpadding="0" width="158" align="center">
  <tr>
    <td valign="bottom" height="42"><img height="40" src="../images/title.gif" width="158" border="0"></td>
  </tr>
  <tr>
    <td class="menu_title" background="../images/title_bg_quit.gif" height="25"><a href="/index.asp" target=_blank><b><span>ϵͳ��ҳ </span></b></a>| <B><span style='cursor: hand' onclick="{if(confirm('ȷ��Ҫ�˳���̨��')){location.href='../quit.asp';}}">�˳�����</span></B></td>
  </tr>
  <tr>
    <td align="center"><font face="Webdings" color="#FFFFFF" style="cursor:hand">5</font></td>
  </tr>
</table>
<% if request.cookies("buyok")("admin")="manage" or request.cookies("buyok")("admin")="�����"  or request.cookies("buyok")("admin")="����" then %>
<table cellspacing="0" cellpadding="0" width="158" align="center">
  <tr>
    <td class="menu_title" id="imgmenu1" onmouseover="this.className='menu_title2';" onclick="showsubmenu(1)" onmouseout="this.className='menu_title';" style="cursor:hand" background="../images/menudown.gif" height="25">&nbsp;<span>�ۺ�����</span> </td>
  </tr>
  <tr>
    <td id="submenu1" style="DISPLAY: none" align="center" >
	<div class="sec_menu" style="WIDTH: 158px">
        <table cellpadding=0 cellspacing=0 align=center width=136>
          <tr>
            <td width="134" height=20><a href="#" onClick="SelectObjItem(this,'Disabled','../Config/manage_System1.asp')">������Ϣ����</a></td>
          </tr>
		  <tr>
            <td width="134" height=20><a href="#" onClick="SelectObjItem(this,'Disabled','../Config/SiteConfig.asp')">����������</a></td>
          </tr>
		  <tr> 
            <td height=20><a href="#" onClick="SelectObjItem(this,'Disabled','../message/manage_message.asp')">��վ������Ϣ</a></td>
          </tr>
          <tr> 
            <td height=20><a href="#" onClick="SelectObjItem(this,'Disabled','../Notice/manage_notice.asp')">֪ͨ����</a></td>
          </tr>
        </table></div><br>
    </td>
  </tr>
</table>
<% end if %>
<% if request.cookies("buyok")("admin")="manage" or request.cookies("buyok")("admin")="�����"  or request.cookies("buyok")("admin")="����"  or request.cookies("buyok")("admin")="������" then %>
<table cellspacing="0" cellpadding="0" width="158" align="center">
  <tr>
    <td class="menu_title" id="imgmenu19" onmouseover="this.className='menu_title2';" onclick="showsubmenu(19)" onmouseout="this.className='menu_title';" style="cursor:hand" background="../images/menudown.gif" height="25">&nbsp;<span>��Ʒ����</span> </td>
  </tr>
  <tr>
    <td id="submenu19" style="DISPLAY: none" align="center" >
	<div class="sec_menu" style="WIDTH: 158px">
        <table cellpadding=0 cellspacing=0 align=center width=136>
          <tr>
            <td width="134" height=20><a href="#" onClick="SelectObjItem(this,'Disabled','../product/manage_product.asp')">�������</a> 
			<a href="#" onClick="SelectObjItem(this,'Disabled','../product/manage_product.asp?action=output')">[����]</a></td>
          </tr>
		  <tr>
            <td width="134" height=20><a href="#" onClick="SelectObjItem(this,'Disabled','../product/manage_product1.asp')">�����Ʒ </a> </td>
          </tr>
		  <tr>
            <td width="134" height=20><a href="#" onClick="SelectObjItem(this,'Disabled','../product/manage_product2.asp?show=yes')">��Ʒ���� </a> </td>
          </tr>
		  <tr>
            <td width="134" height=20><a href="#" onClick="SelectObjItem(this,'Disabled','../product/manage_product4.asp')">��Ʒ����</a> </td>
          </tr>
		  <tr>
            <td width="134" height=20><a href="#" onClick="SelectObjItem(this,'Disabled','../product/manage_product5.asp')">��Ʒת��</a> </td>
          </tr>
		  <tr>
            <td width="134" height=20><a href="#" onClick="SelectObjItem(this,'Disabled','../product/manage_product6.asp')">�¼���Ʒ</a> </td>
          </tr>
        </table>
	  </div><br>
    </td>
  </tr>
</table>
<%end if%>
<% if request.cookies("buyok")("admin")="manage" or request.cookies("buyok")("admin")="�����"  or request.cookies("buyok")("admin")="����һ"  or request.cookies("buyok")("admin")="����"  or request.cookies("buyok")("admin")="������" then %>
<table cellspacing="0" cellpadding="0" width="158" align="center">
  <tr>
    <td class="menu_title" id="imgmenu18" onmouseover="this.className='menu_title2';" onclick="showsubmenu(18)" onmouseout="this.className='menu_title';" style="cursor:hand" background="../images/menudown.gif" height="25">&nbsp;<span>��������</span> </td>
  </tr>
  <tr>
    <td id="submenu18" style="DISPLAY: none" align="center" >
	<div class="sec_menu" style="WIDTH: 158px">
        <table cellpadding=0 cellspacing=0 align=center width=136>
          <tr>
            <td width="134" height=20><a href="#" onClick="SelectObjItem(this,'Disabled','../order/manage_order1.asp?show=yes')">��������</a> </td>
          </tr>
		  <tr>
            <td width="134" height=20><a href="#" onClick="SelectObjItem(this,'Disabled','../order/manage_order2.asp')">��������</a> </td>
          </tr>
		  <% if request.cookies("buyok")("admin")="manage" or request.cookies("buyok")("admin")="�����" or request.cookies("buyok")("admin")="����"then %>
		  <tr>
            <td width="134" height=20><a href="#" onClick="SelectObjItem(this,'Disabled','../order/manage_order5.asp?Status=0')">�ȴ�������</a> </td>
          </tr>
		  <tr>
            <td width="134" height=20><a href="#" onClick="SelectObjItem(this,'Disabled','../order/manage_order5.asp?Status=66')">�Ѹ��������</a> </td>
          </tr>
		  <tr>
            <td width="134" height=20><a href="#" onClick="SelectObjItem(this,'Disabled','../order/manage_order5.asp?Status=88')">����֧���ɹ�</a> </td>
          </tr>
		  <tr>
            <td width="134" height=20><a href="#" onClick="SelectObjItem(this,'Disabled','../order/manage_sto.asp')">��ͨ</a> | <a href="#" onClick="SelectObjItem(this,'Disabled','../order/manage_yto.asp')">Բͨ</a> | <a href="#" onClick="SelectObjItem(this,'Disabled','../order/manage_yd.asp')">�ϴ�</a></td>
          </tr>
		  <tr>
            <td width="134" height=20><a href="#" onClick="SelectObjItem(this,'Disabled','../order/manage_EMS.asp')">EMS</a>   &nbsp;| <a href="#" onClick="SelectObjItem(this,'Disabled','../order/manage_sf-express.asp')">˳��</a></td>
          </tr>
		  <tr>
            <td width="134" height=20><a href="#" onClick="SelectObjItem(this,'Disabled','../order/manage_order3.asp')">��������վ</a> </td>
          </tr>
		  <tr>
            <td width="134" height=20><a href="#" onClick="SelectObjItem(this,'Disabled','../order/manage_order7.asp')">����ͳ��</a> </td>
          </tr>
		  <tr>
            <td width="134" height=20><a href="#" onClick="SelectObjItem(this,'Disabled','../order/manage_order6.asp')">����ͳ��</a> </td>
          </tr>
		  <% end if%>
		  <tr>
            <td width="134" height=20><a href="#" onClick="SelectObjItem(this,'Disabled','../order/numberslist.asp')">�˻�����ݵ�</a> | <a href="#" onClick="SelectObjItem(this,'Disabled','../order/numbers.asp')">���</a> </td>
          </tr>
        </table> 
	  </div><br>
    </td>
  </tr>
</table>
<% end if %>

<% if request.cookies("buyok")("admin")="manage" or request.cookies("buyok")("admin")="�����"  or request.cookies("buyok")("admin")="����" then %>
<table cellspacing="0" cellpadding="0" width="158" align="center">
  <tr>
    <td class="menu_title" id="imgmenu20" onmouseover="this.className='menu_title2';" onclick="showsubmenu(20)" onmouseout="this.className='menu_title';" style="cursor:hand" background="../images/menudown.gif" height="25">&nbsp;<span>��Ա����</span> </td>
  </tr>
  <tr>
    <td id="submenu20" style="DISPLAY: none" align="center" >
	<div class="sec_menu" style="WIDTH: 158px">
        <table cellpadding=0 cellspacing=0 align=center width=136>
          <tr>
            <td width="134" height=20><a href="#" onClick="SelectObjItem(this,'Disabled','../user/manage_user_list.asp?show=yes')">��Ա����</a> </td>
          </tr>
		  <tr>
            <td width="134" height=20><a href="#" onClick="SelectObjItem(this,'Disabled','../user/manage_user_jibie.asp')">��Ա�������� </a> </td>
          </tr>
		  <tr>
            <td width="134" height=20><a href="#" onClick="SelectObjItem(this,'Disabled','../user/manage_user_search.asp')">��Ա����</a> </td>
          </tr>
        </table>
	  </div><br>
    </td>
  </tr>
</table>

<table cellspacing="0" cellpadding="0" width="158" align="center">
  <tr>
    <td class="menu_title" id="imgmenu21" onmouseover="this.className='menu_title2';" onclick="showsubmenu(21)" onmouseout="this.className='menu_title';" style="cursor:hand" background="../images/menudown.gif" height="25">&nbsp;<span>��������</span> </td>
  </tr>
  <tr>
    <td id="submenu21" style="DISPLAY: none" align="center" >
	<div class="sec_menu" style="WIDTH: 158px">
        <table cellpadding=0 cellspacing=0 align=center width=136>
          <tr>
            <td width="134" height=20><a href="#" onClick="SelectObjItem(this,'Disabled','../links/manage_links_list.asp')">�������ӹ���</a> | <a href="#" onClick="SelectObjItem(this,'Disabled','../links/manage_links_list.asp?action=add_links')">���</a></td>
          </tr>
        </table>
	  </div><br>
    </td>
  </tr>
</table>
<table cellspacing="0" cellpadding="0" width="158" align="center">
  <tr>
    <td class="menu_title" id="imgmenu22" onmouseover="this.className='menu_title2';" onclick="showsubmenu(22)" onmouseout="this.className='menu_title';" style="cursor:hand" background="../images/menudown.gif" height="25">&nbsp;<span>��ȫ����</span> </td>
  </tr>
  <tr>
    <td id="submenu22" style="DISPLAY: none" align="center">
	<div class="sec_menu" style="WIDTH: 158px">
        <table cellpadding=0 cellspacing=0 align=center width=136>
          <tr>
            <td width="134" height=20><a href="#" onClick="SelectObjItem(this,'Disabled','../manage/manage_user_add.asp')">��ӹ���Ա </a></td>
          </tr>
		  <tr>
            <td width="134" height=20><a href="#" onClick="SelectObjItem(this,'Disabled','../manage/edit_manage_user.asp')">����Ȩ������</a></td>
          </tr>
        </table>
	  </div><br>
    </td>
  </tr>
</table>
<% end if %>
</tr>
</table><br>


