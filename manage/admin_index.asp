<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>��ӭʹ��-����-��̨����</title>
<link href="images/admin_index.css" rel="stylesheet" type="text/css" >
</head>
<SCRIPT>
var status = 1;
//var Menus = new DvMenuCls;
//document.onclick=Menus.Clear;
function switchSysBar(){
     if (1 == window.status){
		  window.status = 0;
          switchPoint.innerHTML = '<img src="images/left.gif">';
          document.all("frmTitle").style.display="none"
     }
     else{
		  window.status = 1;
          switchPoint.innerHTML = '<img src="images/right.gif">';
          document.all("frmTitle").style.display=""
     }
}
function fullmenu(url){
	if (url==null) {url = "admin_left.asp";}
	parent.leftFrame.location = url;
}
parent.window.defaultStatus="��ǰ�û���IP��ַ��";
function SelectObjItem(Obj,ButtonSymbol,MainUrl)
{  
    if(MainUrl!='')
	{window.parent.parent.frames['main'].location.href=MainUrl;
	}

}
</SCRIPT>
<!--#include file="include/Check_Sql.asp"-->
<!--#include file="../include/connect.asp "-->
<body onselectstart="return false;" oncontextmenu="return false;">
<!--�������ֿ�ʼ-->
<div class="top_table">
<div class="top_table_leftbg">
  <div class="system_logo"><img src="images/logo_up.gif"></div>
  <div class="system_lano">
	<!--ul id="menuTabs">
	  <li><a href="#" onClick="SelectObjItem(this,'Disabled','Config/manage_System1.asp')"><span>������Ϣ����</span></a></li>
      <li onClick="return fullmenu('Include/admin_left.asp');"><a href="#" onClick="SelectObjItem(this,'Disabled','../user/manage_user_list.asp')"><span>����ϵͳ</span></a></li>
	  <li onClick="return fullmenu('index.asp?m=3');"><a href="#" onClick="SelectObjItem(this,'Disabled','admin_mb.asp')"><span>ģ�����</span></a></li>
	  <li onClick="return fullmenu('index.asp?m=4');"><a href="#" onClick="SelectObjItem(this,'Disabled','admin_admin.asp')"><span>�û�����</span></a></li>
	  <li onClick="return fullmenu('index.asp?m=6');"><a href="#" onClick="SelectObjItem(this,'Disabled','admin_severchk.asp')"><span>ϵͳ���</span></a></li>
	  <li><a href="#" onClick="SelectObjItem(this,'Disabled','Admin_UploadFile.asp')"><span>�ϴ��ļ�����</span></a></li>
      </ul-->
	</div>
</div>
</div>
<!--�������ֽ���-->
<table width="100%" height="93%" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td id=frmTitle valign=top name="fmTitle" class="main_left" bgcolor="#000000"><iframe id=frmleft name=leftFrame src="Include/admin_left.asp" class="left_iframe" scrolling="no" frameborder="0" width="180"></iframe></td>
    <td width="1%" valign="top" bgcolor="#337ABB">
    <table border="0" cellpadding="0" cellspacing="0" height="100%">
      <tbody>
        <tr>
          <td onClick="switchSysBar()" style="HEIGHT: 100%"><span class="navPoint" id="switchPoint" title="�ر�/������"><img src="images/right.gif" width="9" height="41" ></span> </td>
        </tr>
      </tbody>
    </table></td>
    <td width="100%" valign="top">
    <table width="100%" border="0" cellpadding="0" cellspacing="0">
      <tr>
        <td height="20" bgcolor="#337ABB">&nbsp;</td>
      </tr>
    </table>
        <table width="100%" border="0" cellspacing="0" cellpadding="0" bgcolor="#C4D8ED">
          <tr height="32">
            <td valign="top" width="8" background="images/bg2.gif"><img src="images/teble_top_left.gif"></td>
            <td background="images/bg2.gif"width="51" ><img src="images/arrow.gif" alt="" align="absmiddle" ></td>
            <td width="486" height="32" background="images/bg2.gif"> ��ӭ����<font color="#FF0066"> <%=request.cookies("buyok")("admin")%></font>			
		    </td>
            <td width="397" background="images/bg2.gif" style="text-align:right; color: #135294; ">
			<a href="admin_index.asp" target='_top'><font color="#135294">��̨��ҳ</font></a> | 
			<a href="../" target="_blank"><font color="#135294">��վ��ҳ</font></a> | 
			<a target="main" href="admin_admin_information.asp" onMouseOver="window.status='�û���Ϣ';return true;" onMouseOut="window.status='';return true;" title='�û���Ϣ'><font color="#135294">����Ա��Ϣ</font></a> | 
			<a href="quit.asp"><font color="#135294">�˳�</font></a> </td>
            <td align="right" valign="top" background="images/bg2.gif" width="8" ><img src="images/teble_top_right.gif" alt="" width="8" height="8"></td>
            <td align="right" width="23" bgcolor="#337ABB"></td>
          </tr>
        </table><iframe frameborder=0 id=frmright name=main scrolling=yes src="mian_index.asp" class="main_iframe"></iframe></td> 
  </tr>
 <tr><td colspan="3" height="20"   bgcolor="#337ABB"> 
 		<table width="100%" border="0" cellspacing="0" cellpadding="0" >
		  <tr height="25">
			<td style="width:38px;"></td>
			<td style="text-align:left;font-family:arial;font-size:11px; color:#dde8f2;">Copyright 2008-2010 Tencent.All Rights Reserved , ��Ȩ���� www.XiDui.net ����
</td>
			<td style="text-align:right;padding-right:20px; color:#FFFFFF"><a href="#" onClick="SelectObjItem(this,'Disabled','admin_reloadcache.asp')">������</a> | <a href="#" onClick="SelectObjItem(this,'Disabled','admin_exit.asp')" >ע���˳�</a></td>
		  </tr>
		</table>

 
 </td></tr> 
  
  
</table>

</body>
</html>