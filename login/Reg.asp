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

//判断表单输入正误
function Checkreg()
{
	if (document.ADDUser.UserId.value.length < 4 || document.ADDUser.UserId.value.length >16) {
		alert("提示：请输入用户名，有效长度为4-16位");
		document.ADDUser.UserId.focus();
		return false;
	}
	if (document.ADDUser.pw1.value.length <6 || document.ADDUser.pw1.value.length >16) {
		alert("提示：请输入密码，最少6位，最长16位。");
		document.ADDUser.pw1.focus();
		return false;
	}
	if (document.ADDUser.pw2.value.length <6 || document.ADDUser.pw1.value.length >16) {
		alert("提示：请输入密码，最少6位，最长16位。");
		document.ADDUser.pw2.focus();
		return false;
	}
	if (document.ADDUser.pw1.value != document.ADDUser.pw2.value) {
		alert("提示：您两次输入的密码不一样，请检查后重新输入");
		document.ADDUser.pw2.focus();
		return false;
	}
	if (!fucPWDchk(document.ADDUser.pw1.value)){
		alert("提示：只能使用纯数字和大、小写字母作为密码");
		document.ADDUser.pw1.focus();
		return false;
		}
	if (!fucPWDchk(document.ADDUser.pw2.value)){
		alert("提示：只能使用纯数字和大、小写字母作为密码");
		document.ADDUser.pw1.focus();
		return false;
		}
	if (document.ADDUser.Username.value.length <1 || document.ADDUser.Username.value.length >12) {
		alert("提示：请输入你的真实姓名");
		document.ADDUser.Username.focus();
		return false;
	}
	if (document.ADDUser.HomePhone.value.length <10 || document.ADDUser.HomePhone.value.length >12) {
		alert("提示：请输入你的手机或则你的座机");
		document.ADDUser.HomePhone.focus();
		return false;
	}
	if (document.ADDUser.province.value.length < 1 ) {
		alert("提示：请选择你的省");
		document.ADDUser.province.focus();
		return false;
	}
	if (document.ADDUser.city.value.length < 1 ) {
		alert("提示：请选择你的市区");
		document.ADDUser.city.focus();
		return false;
	}
	if (document.ADDUser.Address.value.length < 6 ) {
		alert("提示：请输入详细的地址");
		document.ADDUser.Address.focus();
		return false;
	}
	if (document.ADDUser.taobao.value.length < 10 ) {
		alert("提示：请输入您的店铺地址");
		document.ADDUser.taobao.focus();
		return false;
	}
	if (document.ADDUser.taobao_qq.value.length < 5 ) {
		alert("提示：请输入您的QQ帐号");
		document.ADDUser.taobao_qq.focus();
		return false;
	}
	if (document.ADDUser.check1.value.length < 1 ) {
		alert("提示：请输入您的验证码");
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
where[1] = new comefrom("北京省","|东城|西城|崇文|宣武|朝阳|丰台|石景山|海淀|门头沟|房山|通州|顺义|昌平|大兴|平谷|怀柔|密云|延庆");  
where[2] = new comefrom("上海省","|黄浦|卢湾|徐汇|长宁|静安|普陀|闸北|虹口|杨浦|闵行|宝山|嘉定|浦东|金山|松江|青浦|南汇|奉贤|崇明");  
where[3] = new comefrom("天津省","|和平|东丽|河东|西青|河西|津南|南开|北辰|河北|武清|红挢|塘沽|汉沽|大港|宁河|静海|宝坻|蓟县");  
where[4] = new comefrom("重庆省","|万州|涪陵|渝中|大渡口|江北|沙坪坝|九龙坡|南岸|北碚|万盛|双挢|渝北|巴南|黔江|长寿|綦江|潼南|铜梁|大足|荣昌|壁山|梁平|城口|丰都|垫江|武隆|忠县|开县|云阳|奉节|巫山|巫溪|石柱|秀山|酉阳|彭水|江津|合川|永川|南川");  
where[5] = new comefrom("河北省","|石家庄|邯郸|邢台|保定|张家口|承德|廊坊|唐山|秦皇岛|沧州|衡水");  
where[6] = new comefrom("山西省","|太原|大同|阳泉|长治|晋城|朔州|吕梁|忻州|晋中|临汾|运城");  
where[7] = new comefrom("内蒙古省","|呼和浩特|包头|乌海|赤峰|呼伦贝尔盟|阿拉善盟|哲里木盟|兴安盟|乌兰察布盟|锡林郭勒盟|巴彦淖尔盟|伊克昭盟");  
where[8] = new comefrom("辽宁省","|沈阳|大连|鞍山|抚顺|本溪|丹东|锦州|营口|阜新|辽阳|盘锦|铁岭|朝阳|葫芦岛");  
where[9] = new comefrom("吉林省","|长春|吉林|四平|辽源|通化|白山|松原|白城|延边");  
where[10] = new comefrom("黑龙江省","|哈尔滨|齐齐哈尔|牡丹江|佳木斯|大庆|绥化|鹤岗|鸡西|黑河|双鸭山|伊春|七台河|大兴安岭");  
where[11] = new comefrom("江苏省","|南京|镇江|苏州|南通|扬州|盐城|徐州|连云港|常州|无锡|宿迁|泰州|淮安");  
where[12] = new comefrom("浙江省","|杭州|宁波|温州|嘉兴|湖州|绍兴|金华|衢州|舟山|台州|丽水");  
where[13] = new comefrom("安徽省","|合肥|芜湖|蚌埠|马鞍山|淮北|铜陵|安庆|黄山|滁州|宿州|池州|淮南|巢湖|阜阳|六安|宣城|亳州");  
where[14] = new comefrom("福建省","|福州|厦门|莆田|三明|泉州|漳州|南平|龙岩|宁德");  
where[15] = new comefrom("江西省","|南昌市|景德镇|九江|鹰潭|萍乡|新馀|赣州|吉安|宜春|抚州|上饶");  
where[16] = new comefrom("山东省","|济南|青岛|淄博|枣庄|东营|烟台|潍坊|济宁|泰安|威海|日照|莱芜|临沂|德州|聊城|滨州|菏泽");  
where[17] = new comefrom("河南省","|郑州|开封|洛阳|平顶山|安阳|鹤壁|新乡|焦作|濮阳|许昌|漯河|三门峡|南阳|商丘|信阳|周口|驻马店|济源");  
where[18] = new comefrom("湖北省","|武汉|宜昌|荆州|襄樊|黄石|荆门|黄冈|十堰|恩施|潜江|天门|仙桃|随州|咸宁|孝感|鄂州"); 
where[19] = new comefrom("湖南省","|长沙|常德|株洲|湘潭|衡阳|岳阳|邵阳|益阳|娄底|怀化|郴州|永州|湘西|张家界");  
where[20] = new comefrom("广东省","|广州|深圳|珠海|汕头|东莞|中山|佛山|韶关|江门|湛江|茂名|肇庆|惠州|梅州|汕尾|河源|阳江|清远|潮州|揭阳|云浮");  
where[21] = new comefrom("广西省","|南宁|柳州|桂林|梧州|北海|防城港|钦州|贵港|玉林|南宁地区|柳州地区|贺州|百色|河池");  
where[22] = new comefrom("海南省","|海口|三亚");  
where[23] = new comefrom("四川省","|成都|绵阳|德阳|自贡|攀枝花|广元|内江|乐山|南充|宜宾|广安|达川|雅安|眉山|甘孜|凉山|泸州");  
where[24] = new comefrom("贵州省","|贵阳|六盘水|遵义|安顺|铜仁|黔西南|毕节|黔东南|黔南");  
where[25] = new comefrom("云南省","|昆明|大理|曲靖|玉溪|昭通|楚雄|红河|文山|思茅|西双版纳|保山|德宏|丽江|怒江|迪庆|临沧"); 
where[26] = new comefrom("西藏省","|拉萨|日喀则|山南|林芝|昌都|阿里|那曲");  
where[27] = new comefrom("陕西省","|西安|宝鸡|咸阳|铜川|渭南|延安|榆林|汉中|安康|商洛");  
where[28] = new comefrom("甘肃省","|兰州|嘉峪关|金昌|白银|天水|酒泉|张掖|武威|定西|陇南|平凉|庆阳|临夏|甘南");  
where[29] = new comefrom("宁夏省","|银川|石嘴山|吴忠|固原");  
where[30] = new comefrom("青海省","|西宁|海东|海南|海北|黄南|玉树|果洛|海西");  
where[31] = new comefrom("新疆省","|乌鲁木齐|石河子|克拉玛依|伊犁|巴音郭勒|昌吉|克孜勒苏柯尔克孜|博尔塔拉|吐鲁番|哈密|喀什|和田|阿克苏");  
where[32] = new comefrom("香港行政区","|香港");  
where[33] = new comefrom("澳门行政区","|澳门");  
where[34] = new comefrom("台湾行政区","|台北|高雄|台中|台南|屏东|南投|云林|新竹|彰化|苗栗|嘉义|花莲|桃园|宜兰|基隆|台东|金门|马祖|澎湖");  
where[35] = new comefrom("其它省","|北美洲|南美洲|亚洲|非洲|欧洲|大洋洲");  
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
        <td width="188" height="30" align="right">用户名：</td>
        <td width="512" height="30"><input name="UserId" type="text" id="UserId" size="30" /><input name="button" type="button"  onClick="javascript:window.open('check1.asp?'+document.ADDUser.UserId.value,'newwindow','scrollbars=no,resizable=no,left=400,top=300,width=282,height=220;')"  value="检测用户名">请填写有效的下次登录的用户名</td>
      </tr>
      <tr>
        <td height="30" align="right">请设定密码：</td>
        <td height="30"><input name="pw1" type="password" id="pw1" size="33" />
          密码请设为6-16位字母或数字</td>
      </tr>
      <tr>
        <td height="30" align="right">请再次输入设定密码：</td>
        <td height="30"><input name="pw2" type="password" id="pw2" size="33" /></td>
      </tr>
      <tr>
        <td height="30" align="right" bgcolor="#FFFFFF" class="size_16">真实姓名：</td>
        <td height="30" bgcolor="#FFFFFF"><input name="Username" type="text" class=uset_input size="30">
          <span class=red> *</span>真实姓名</td>
      </tr>
      <tr>
        <td height="30" align="right" bgcolor="#FFFFFF" class="size_16">移动电话：</td>
        <td height="30" bgcolor="#FFFFFF"><input name="HomePhone"  class=uset_input id="HomePhone" size="30">
          <span class=red> *</span>格式如:028-86956007或者13901234567 </td>
      </tr>
      <tr>
        <td height="30" align="right" bgcolor="#FFFFFF" class="size_16">所在城市：</td>
        <td height="30" bgcolor="#FFFFFF">省份<select name="province" class="uset_input" onChange = "select()"></select>城市<select name="city" class="uset_input" onChange = "select()"></select>不能超过12个字符（6个汉字） </td>
      </tr>
      <tr>
        <td height="30" align="right" bgcolor="#FFFFFF" class="size_16">详细地址：</td>
        <td height="30" bgcolor="#FFFFFF"><input  name=Address class="uset_input" id=Address size=60>
          <span class="size_16"><span class=red>*</span></span>详细地址</td>
      </tr>
      <TR>
        <TD height=30 align=right bgcolor="#FAFCF3" ><span class="red">网店地址：</span></TD>
        <TD height="30" colspan="2" align=left bgcolor="#FAFCF3" ><INPUT name=taobao class=uset_input id="taobao" size="30">
          <span class="red">* 格式如:http://shop*******.taobao.com </span></TD>
      </TR>
      <TR>
        <TD height=30 align=right bgcolor="#FAFCF3" ><span class="red">淘宝用户名：</span></TD>
        <TD height="30" colspan="2" align=left bgcolor="#FAFCF3" ><INPUT name=taobao_name class=uset_input id="taobao_name" size="30">
          <span class="red"> 您的淘宝账号</span></TD>
      </TR>
      <TR>
        <TD height=30 align=right bgcolor="#FAFCF3"><span class="red">QQ号码：</span></TD>
        <TD height="30" colspan="2" align=left bgcolor="#FAFCF3" ><INPUT name=taobao_qq class=uset_input id="taobao_qq" size="30">
          <span class="red">* 你最常用的QQ号</span></TD>
      </TR>
      
	  <tr>
        <td height="30" align="right">请输入验证码：</td>
        <td height="30"><input name="check1" type="text" size="6" maxlength="6" /> <font color="#C04411"><%=check%></font> <input type="hidden" name="check2" value="<%=check%>" />请输入左边的验证码!<input type="hidden" name="check" value="reg" /></td>
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