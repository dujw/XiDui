<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<!--#include file="../include/connect.asp"-->
<%
	Dim check
	Randomize timer
	check=Int((899999)*Rnd +100000)
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title><%=sitename%></title>
<meta name="keywords" content="<%=sitedescription%>">
<meta name="description" content="<%=sitekeywords%>">
<link href="../image/web_css/login.css" rel="stylesheet" type="text/css" />
</head>
<script language="JavaScript">
function JM_sh(ob){
if (ob.style.display=="none"){ob.style.display=""}else{ob.style.display="none"};
}

function fucPWDchk(str) 
{ 
var strSource ="0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ"; 
var ch; 
var i; 
var temp; 

for (i=0;i<=(str.length-1);i++) 
{ 

ch = str.charAt(i); 
temp = strSource.indexOf(ch); 
if (temp==-1) 
{ 
return 0; 
} 
} 
if (strSource.indexOf(ch)==-1) 
{ 
return 0; 
} 
else 
{ 
return 1; 
} 
} 
function jtrim(str) 
{ while (str.charAt(0)==" ") 
{str=str.substr(1);} 
while (str.charAt(str.length-1)==" ") 
{str=str.substr(0,str.length-1);} 
return(str); 
} 

//�жϱ���������
function Checkreg()
{
	if (document.ADDUser.UserId.value.length < 4 || document.ADDUser.UserId.value.length >16) {
		alert("��ʾ���������û�������Ч����Ϊ4-16λ");
		document.ADDUser.UserId.focus();
		return false;
	}
	if (document.ADDUser.pw1.value.length <6 || document.ADDUser.pw1.value.length >16) {
		alert("��ʾ�����������룬����6λ���16λ��");
		document.ADDUser.pw1.focus();
		return false;
	}
	if (document.ADDUser.pw2.value.length <6 || document.ADDUser.pw1.value.length >16) {
		alert("��ʾ�����������룬����6λ���16λ��");
		document.ADDUser.pw2.focus();
		return false;
	}
	if (document.ADDUser.pw1.value != document.ADDUser.pw2.value) {
		alert("��ʾ����������������벻һ�����������������");
		document.ADDUser.pw2.focus();
		return false;
	}
	if (!fucPWDchk(document.ADDUser.pw1.value)){
		alert("��ʾ��ֻ��ʹ�ô����ֺʹ�Сд��ĸ��Ϊ����");
		document.ADDUser.pw1.focus();
		return false;
		}
	if (!fucPWDchk(document.ADDUser.pw2.value)){
		alert("��ʾ��ֻ��ʹ�ô����ֺʹ�Сд��ĸ��Ϊ����");
		document.ADDUser.pw1.focus();
		return false;
		}
	if (document.ADDUser.Username.value.length <1 || document.ADDUser.Username.value.length >12) {
		alert("��ʾ�������������ʵ����");
		document.ADDUser.Username.focus();
		return false;
	}
	if (document.ADDUser.HomePhone.value.length <10 || document.ADDUser.HomePhone.value.length >12) {
		alert("��ʾ������������ֻ������������");
		document.ADDUser.HomePhone.focus();
		return false;
	}
	if (document.ADDUser.province.value.length < 1 ) {
		alert("��ʾ����ѡ�����ʡ");
		document.ADDUser.province.focus();
		return false;
	}
	if (document.ADDUser.city.value.length < 1 ) {
		alert("��ʾ����ѡ���������");
		document.ADDUser.city.focus();
		return false;
	}
	if (document.ADDUser.Address.value.length < 6 ) {
		alert("��ʾ����������ϸ�ĵ�ַ");
		document.ADDUser.Address.focus();
		return false;
	}
	if (document.ADDUser.taobao.value.length < 10 ) {
		alert("��ʾ�����������ĵ��̵�ַ");
		document.ADDUser.taobao.focus();
		return false;
	}
	if (document.ADDUser.taobao_qq.value.length < 5 ) {
		alert("��ʾ������������QQ�ʺ�");
		document.ADDUser.taobao_qq.focus();
		return false;
	}
	if (document.ADDUser.check1.value.length < 1 ) {
		alert("��ʾ��������������֤��");
		document.ADDUser.check1.focus();
		return false;
	}
}
</script>

<script language="javascript"> 
<!-- 
var where = new Array(35);  
function comefrom(loca,locacity) { this.loca = loca; this.locacity = locacity; }  
where[0]= new comefrom("",""); 
where[1] = new comefrom("����ʡ","|����|����|����|����|����|��̨|ʯ��ɽ|����|��ͷ��|��ɽ|ͨ��|˳��|��ƽ|����|ƽ��|����|����|����");  
where[2] = new comefrom("�Ϻ�ʡ","|����|¬��|���|����|����|����|բ��|���|����|����|��ɽ|�ζ�|�ֶ�|��ɽ|�ɽ�|����|�ϻ�|����|����");  
where[3] = new comefrom("���ʡ","|��ƽ|����|�Ӷ�|����|����|����|�Ͽ�|����|�ӱ�|����|����|����|����|���|����|����|����|����");  
where[4] = new comefrom("����ʡ","|����|����|����|��ɿ�|����|ɳƺ��|������|�ϰ�|����|��ʢ|˫��|�山|����|ǭ��|����|�뽭|����|ͭ��|����|�ٲ�|��ɽ|��ƽ|�ǿ�|�ᶼ|�潭|��¡|����|����|����|���|��ɽ|��Ϫ|ʯ��|��ɽ|����|��ˮ|����|�ϴ�|����|�ϴ�");  
where[5] = new comefrom("�ӱ�ʡ","|ʯ��ׯ|����|��̨|����|�żҿ�|�е�|�ȷ�|��ɽ|�ػʵ�|����|��ˮ");  
where[6] = new comefrom("ɽ��ʡ","|̫ԭ|��ͬ|��Ȫ|����|����|˷��|����|����|����|�ٷ�|�˳�");  
where[7] = new comefrom("���ɹ�ʡ","|���ͺ���|��ͷ|�ں�|���|���ױ�����|��������|����ľ��|�˰���|�����첼��|���ֹ�����|�����׶���|��������");  
where[8] = new comefrom("����ʡ","|����|����|��ɽ|��˳|��Ϫ|����|����|Ӫ��|����|����|�̽�|����|����|��«��");  
where[9] = new comefrom("����ʡ","|����|����|��ƽ|��Դ|ͨ��|��ɽ|��ԭ|�׳�|�ӱ�");  
where[10] = new comefrom("������ʡ","|������|�������|ĵ����|��ľ˹|����|�绯|�׸�|����|�ں�|˫Ѽɽ|����|��̨��|���˰���");  
where[11] = new comefrom("����ʡ","|�Ͼ�|��|����|��ͨ|����|�γ�|����|���Ƹ�|����|����|��Ǩ|̩��|����");  
where[12] = new comefrom("�㽭ʡ","|����|����|����|����|����|����|��|����|��ɽ|̨��|��ˮ");  
where[13] = new comefrom("����ʡ","|�Ϸ�|�ߺ�|����|��ɽ|����|ͭ��|����|��ɽ|����|����|����|����|����|����|����|����|����");  
where[14] = new comefrom("����ʡ","|����|����|����|����|Ȫ��|����|��ƽ|����|����");  
where[15] = new comefrom("����ʡ","|�ϲ���|������|�Ž�|ӥ̶|Ƽ��|����|����|����|�˴�|����|����");  
where[16] = new comefrom("ɽ��ʡ","|����|�ൺ|�Ͳ�|��ׯ|��Ӫ|��̨|Ϋ��|����|̩��|����|����|����|����|����|�ĳ�|����|����");  
where[17] = new comefrom("����ʡ","|֣��|����|����|ƽ��ɽ|����|�ױ�|����|����|���|���|���|����Ͽ|����|����|����|�ܿ�|פ���|��Դ");  
where[18] = new comefrom("����ʡ","|�人|�˲�|����|�差|��ʯ|����|�Ƹ�|ʮ��|��ʩ|Ǳ��|����|����|����|����|Т��|����"); 
where[19] = new comefrom("����ʡ","|��ɳ|����|����|��̶|����|����|����|����|¦��|����|����|����|����|�żҽ�");  
where[20] = new comefrom("�㶫ʡ","|����|����|�麣|��ͷ|��ݸ|��ɽ|��ɽ|�ع�|����|տ��|ï��|����|����|÷��|��β|��Դ|����|��Զ|����|����|�Ƹ�");  
where[21] = new comefrom("����ʡ","|����|����|����|����|����|���Ǹ�|����|���|����|��������|���ݵ���|����|��ɫ|�ӳ�");  
where[22] = new comefrom("����ʡ","|����|����");  
where[23] = new comefrom("�Ĵ�ʡ","|�ɶ�|����|����|�Թ�|��֦��|��Ԫ|�ڽ�|��ɽ|�ϳ�|�˱�|�㰲|�ﴨ|�Ű�|üɽ|����|��ɽ|����");  
where[24] = new comefrom("����ʡ","|����|����ˮ|����|��˳|ͭ��|ǭ����|�Ͻ�|ǭ����|ǭ��");  
where[25] = new comefrom("����ʡ","|����|����|����|��Ϫ|��ͨ|����|���|��ɽ|˼é|��˫����|��ɽ|�º�|����|ŭ��|����|�ٲ�"); 
where[26] = new comefrom("����ʡ","|����|�տ���|ɽ��|��֥|����|����|����");  
where[27] = new comefrom("����ʡ","|����|����|����|ͭ��|μ��|�Ӱ�|����|����|����|����");  
where[28] = new comefrom("����ʡ","|����|������|���|����|��ˮ|��Ȫ|��Ҵ|����|����|¤��|ƽ��|����|����|����");  
where[29] = new comefrom("����ʡ","|����|ʯ��ɽ|����|��ԭ");  
where[30] = new comefrom("�ຣʡ","|����|����|����|����|����|����|����|����");  
where[31] = new comefrom("�½�ʡ","|��³ľ��|ʯ����|��������|����|��������|����|�������տ¶�����|��������|��³��|����|��ʲ|����|������");  
where[32] = new comefrom("���������","|���");  
where[33] = new comefrom("����������","|����");  
where[34] = new comefrom("̨��������","|̨��|����|̨��|̨��|����|��Ͷ|����|����|�û�|����|����|����|��԰|����|��¡|̨��|����|����|���");  
where[35] = new comefrom("����ʡ","|������|������|����|����|ŷ��|������");  
function select() { 
with(document.ADDUser.province) { var loca2 = options[selectedIndex].value; } 
for(i = 0;i < where.length;i ++) { 
if (where[i].loca == loca2) { 
loca3 = (where[i].locacity).split("|"); 
for(j = 0;j < loca3.length;j++) { with(document.ADDUser.city) { length = loca3.length; options[j].text = loca3[j]; options[j].value = loca3[j]; var loca4=options[selectedIndex].value;}} 
break; 
}} 
document.ADDUser.Address.value=loca2+loca4; 
} 
function init() { 
with(document.ADDUser.province) { 
length = where.length; 
for(k=0;k<where.length;k++) { options[k].text = where[k].loca; options[k].value = where[k].loca; } 
options[selectedIndex].text = where[0].loca; options[selectedIndex].value = where[0].loca; 
} 
with(document.ADDUser.city) { 
loca3 = (where[0].locacity).split("|"); 
length = loca3.length; 
for(l=0;l<length;l++) { options[l].text = loca3[l]; options[l].value = loca3[l]; } 
options[selectedIndex].text = loca3[0]; options[selectedIndex].value = loca3[0]; 
}} 
--> 
</script> 
<body onLoad="init();init2();"> 
<!--#include file="../include/login/top.asp"-->
<div id='div_Public'>
  <div id='mid_login'>
    <table width="700" border="0" align="center" cellpadding="0" cellspacing="0">
	<form name="ADDUser" method="POST" action="reg_save.asp" onSubmit="return Checkreg();">
      <tr>
        <td height="50" colspan="2" valign="bottom" style="border-bottom-width: 2px;border-bottom-style: solid;border-bottom-color: #010101;"><img src="../image/login/mid/zhuce.jpg" width="102" height="25" />
		</td>
      </tr>
      <tr>
        <td width="188" height="30" align="right">�û�����</td>
        <td width="512" height="30"><input name="UserId" type="text" id="UserId" size="30" /><input name="button" type="button"  onClick="javascript:window.open('check1.asp?'+document.ADDUser.UserId.value,'newwindow','scrollbars=no,resizable=no,left=400,top=300,width=282,height=220;')"  value="����û���">����д��Ч���´ε�¼���û���</td>
      </tr>
      <tr>
        <td height="30" align="right">���趨���룺</td>
        <td height="30"><input name="pw1" type="password" id="pw1" size="33" />
          ��������Ϊ6-16λ��ĸ������</td>
      </tr>
      <tr>
        <td height="30" align="right">���ٴ������趨���룺</td>
        <td height="30"><input name="pw2" type="password" id="pw2" size="33" /></td>
      </tr>
      <tr>
        <td height="30" align="right" bgcolor="#FFFFFF" class="size_16">��ʵ������</td>
        <td height="30" bgcolor="#FFFFFF"><input name="Username" type="text" class=uset_input size="30">
          <span class=red> *</span>��ʵ����</td>
      </tr>
      <tr>
        <td height="30" align="right" bgcolor="#FFFFFF" class="size_16">�ƶ��绰��</td>
        <td height="30" bgcolor="#FFFFFF"><input name="HomePhone"  class=uset_input id="HomePhone" size="30">
          <span class=red> *</span>��ʽ��:028-86956007����13901234567 </td>
      </tr>
      <tr>
        <td height="30" align="right" bgcolor="#FFFFFF" class="size_16">���ڳ��У�</td>
        <td height="30" bgcolor="#FFFFFF">ʡ��<select name="province" class="uset_input" onChange = "select()"></select>����<select name="city" class="uset_input" onChange = "select()"></select>���ܳ���12���ַ���6�����֣� </td>
      </tr>
      <tr>
        <td height="30" align="right" bgcolor="#FFFFFF" class="size_16">��ϸ��ַ��</td>
        <td height="30" bgcolor="#FFFFFF"><input  name=Address class="uset_input" id=Address size=60>
          <span class="size_16"><span class=red>*</span></span>��ϸ��ַ</td>
      </tr>
      <TR>
        <TD height=30 align=right bgcolor="#FAFCF3" ><span class="red">�����ַ��</span></TD>
        <TD height="30" colspan="2" align=left bgcolor="#FAFCF3" ><INPUT name=taobao class=uset_input id="taobao" size="30">
          <span class="red">* ��ʽ��:http://shop*******.taobao.com </span></TD>
      </TR>
      <TR>
        <TD height=30 align=right bgcolor="#FAFCF3" ><span class="red">�Ա��û�����</span></TD>
        <TD height="30" colspan="2" align=left bgcolor="#FAFCF3" ><INPUT name=taobao_name class=uset_input id="taobao_name" size="30">
          <span class="red"> �����Ա��˺�</span></TD>
      </TR>
      <TR>
        <TD height=30 align=right bgcolor="#FAFCF3"><span class="red">QQ���룺</span></TD>
        <TD height="30" colspan="2" align=left bgcolor="#FAFCF3" ><INPUT name=taobao_qq class=uset_input id="taobao_qq" size="30">
          <span class="red">* ����õ�QQ��</span></TD>
      </TR>
      
	  <tr>
        <td height="30" align="right">��������֤�룺</td>
        <td height="30"><input name="check1" type="text" size="6" maxlength="6" /> <font color="#C04411"><%=check%></font> <input type="hidden" name="check2" value="<%=check%>" />��������ߵ���֤��!<input type="hidden" name="check" value="reg" /></td>
      </tr>
	  <tr>
        <td height="30" align="right">&nbsp;</td>
        <td height="30"><input src="../image/login/mid/wangc.jpg" type="image" /></td>
      </tr>
	  </form>
    </table>
  </div>
</div>
<!--#include file="../include/login/bootom.asp"-->
</body>
</html>
<% call CloseConn() %>