<%
	if kaiguan = "0" then 
		response.Write "<br><center>"&guanbi&"</font>"
		response.End()
	end if 
%>
<%
		SSQL="select top 1 id,Ntitle,NEditTime from DU_Notice where NOverTime>getdate() order by id desc"
'response.write SSQL
set djw_rs=Server.Createobject("ADODB.RecordSet")
djw_rs.open SSQL,conn,1,3
if djw_rs.eof and djw_rs.bof  then
	response.write "<!--by djw 2012.2.25-->"
else
%>
<script type="text/javascript" src="js_lib/djw/lhgcore.lhgdialog.min.js?self=true&skin=mac"></script>
<script type="text/javascript">
function dk_notice(){
	$.dialog({width:700,height:280,opacity:0.5,background:'#ccc',lock:true,title: '���� <%=djw_rs(2)%> ����',content: 'url:Notice.asp?id=<%=djw_rs(0)%>',button: [{ name: '�ر�'}],padding:5});
}
</script>
<table width="950" border="0" align="center" cellpadding="0" cellspacing="0" style=" border:#CCCCCC solid 1px;">
  <tr >
    <td height="30" align="left"><strong><center><a href="#" onclick="dk_notice()">��������֪ͨ:<font color="#FF0000"><%=djw_rs(1)%></font>&nbsp;&nbsp;[<%=djw_rs(2)%>]</a></center></strong>
</td>
  </tr>
</table>
<!--<table width="950" border="0" align="center" cellpadding="0" cellspacing="0" style=" border:#CCCCCC solid 0px;">
  <tr >
    <td height="30" align="left"><img src="" border="0" />
	</td>
  </tr>
</table>-->
<%
djw_rs.close()
end if%>
<div class="logo">
<div class="logo_a">
<div class="logoleft"><a href="index.php"><img src="../image/logo/logo.png" width="342" height="86" border="0" /></a></div>
<div class="logolefta" id="welcome"><div id="append_parent"> <% if request.cookies("buyok")("userid")="" then%>
<a href="../login.asp" target="_blank"  class="huang">��½</a> | <a href="../reg.asp" target="_blank"  class="huang">ע��</a>
<%else
	response.Write "��ã�" & request.cookies("buyok")("userid") & "��  <a href='../login/logout.asp' class='huang'>[�˳�]<a> <a href='../user_center.asp' class='huang'>�ҵĶ���</a>  <a href='../check.asp' class='huang'>���ﳵ</a>"
	end if%></div>
<a href="../user_sjb.asp"><strong>���ݰ�����</strong></a>
    <a target="_blank" href="http://wpa.qq.com/msgrd?v=3&uin=1504490396&site=����С��&menu=yes"><img border="0" src="http://wpa.qq.com/pa?p=2:1504490396:45" alt="���������ҷ���Ϣ" title="���������ҷ���Ϣ">����С��</a>
	<a target="_blank" href="http://wpa.qq.com/msgrd?v=3&uin=2262462307&site=����С��&menu=yes"><img border="0" src="http://wpa.qq.com/pa?p=2:2262462307:45" alt="���������ҷ���Ϣ" title="���������ҷ���Ϣ">����С��</a>
	<a target="_blank" href="http://wpa.qq.com/msgrd?v=3&uin=2404470991&site=����С��&menu=yes"><img border="0" src="http://wpa.qq.com/pa?p=2:2404470991:45" alt="���������ҷ���Ϣ" title="���������ҷ���Ϣ">����С��</a>
    <a target="_blank" href="http://wpa.qq.com/msgrd?v=3&uin=2486488687&site=����С��&menu=yes"><img border="0" src="http://wpa.qq.com/pa?p=2:2486488687:45" alt="���������ҷ���Ϣ" title="���������ҷ���Ϣ">����С��</a>
</div>
<div class="logoleftb">
 <a href="Contact.asp?id=12" target="_blank">��������</a> | <a href="#" onclick="window.external.addFavorite('http://www.xidui.net/','������װ ��ó���� �����Դ�ṩ һ������')">�����ղ�</a> </div>
</div>
<div class="dhbj">
    <div class="dh">
        <div class="dhb">
            <div class="menu_box">
                <div class="menu" id="chromemenu">
                    <ul>
                              <%
                LarCode=trim(request("LarCode"))'��Ʒ�����
                MidCode=trim(request("MidCode"))'��ƷС����
                if LarCode = "" then
                    response.Write "<li class='m_2' onMouseOver=this.className='m_2_h' onMouseOut=this.className='m_2'><a href='../index.asp' class='cur'>��վ��ҳ</a></li>"
                    'response.Write "<td  width=85 height=30 align=center background=images/index/top1.jpg><a href=../index.asp class=titile_1>��վ��ҳ</a></td><td width=2></td>"
                else
                    'response.Write "<td  width=85 height=30 align=center background=images/index/top1.jpg><a class=text_1 href=../index.asp>��վ��ҳ</a></td><td width=2></td>"
                    response.Write "<li class='m_2' onMouseOver=this.className='m_2_h' onMouseOut=this.className='m_2'><a href='../index.asp' class='cur'>��վ��ҳ</a></li>"
                end if
                set rs_product=server.createobject("adodb.recordset") 
                sql_product="select distinct LarCode,LarSeq from dk501_productclass order by LarSeq" 
                rs_product.Open sql_product,conn,1,1 		
                if rs_product.bof and rs_product.eof then
                    response.write "�������ڸ����У������ĵȺ�"
                else
                    Do While Not rs_product.eof	
                    if LarCode =  Trim(rs_product("LarCode")) then 
                        'response.write "<td  width=85 height=30 align=center background=images/index/top1.jpg><a class=titile_1 href=../product.asp?LarCode="&Server.URLEncode(Trim(rs_product("LarCode")))&">"&Trim(rs_product("LarCode"))&"</a></td><td width=2></td>"
                        response.Write "<li class='m_2' onMouseOver=this.className='m_2_h' onMouseOut=this.className='m_2'><a href=../product.asp?LarCode="&Server.URLEncode(Trim(rs_product("LarCode")))&" class='cur'>"&Trim(rs_product("LarCode"))&"</a></li>"
            
                    else
                        'response.write "<td  width=85 height=30 align=center background=images/index/top1.jpg><a class=text_1 href=../product.asp?LarCode="&Server.URLEncode(Trim(rs_product("LarCode")))&">"&Trim(rs_product("LarCode"))&"</a></td><td width=2></td>"
                        response.Write "<li class='m_2' onMouseOver=this.className='m_2_h' onMouseOut=this.className='m_2'><a href=../product.asp?LarCode="&Server.URLEncode(Trim(rs_product("LarCode")))&"  class='cur'>"&Trim(rs_product("LarCode"))&"</a></li>"
                    end if 
                    rs_product.movenext
                      if rs_product.eof then exit do
                    loop		
                end if
                rs_product.close
                set rs_product=nothing
                'response.Write "<td  width=85 height=30 align=center background=images/index/top1.jpg><a class=text_1 href='xuanchuan.asp'><span style='color:#F00; font-weight:bold;'>PSDԴ�ļ�</span></a></td><td width=2></td>"
                response.Write "<li class='m_2' onMouseOver=this.className='m_2_h' onMouseOut=this.className='m_2'><a href=xuanchuan.asp  class='cur'><span style='color:#F00; font-weight:bold;'>PSDԴ�ļ�</span></a></li>"
                %>
                                </ul>
                </div>
            </div>
        </div>
    <!--<div class="dhgouwu"><div class="dgwz">&nbsp;<a href="flow.php" title="�鿴���ﳵ"><a href=��flow.php��>���ﳵ��<span id=prdCount class=ba10000> 0 </span>����Ʒ</a></a><a href="flow.php" class="huangc" target="_blank">&nbsp;<img src="themes/ok-moban_langrand/images/jsbtn.gif"></a></div>
    </div>-->
    </div>
</div>
	<div class="so">
	<DIV class=sogjc>&nbsp;&nbsp;&nbsp;&nbsp;<B>���������ʣ�</B>
<a href="product.asp?name_key=2012%D0%C2%BF%EE">2012�¿�</a> &nbsp;&nbsp;<a href="product.asp?name_key=T%D0%F4">T��</a> &nbsp;&nbsp;<a href="product.asp?name_key=%B3%C4%D2%C2">����</a> &nbsp;&nbsp;<a href="product.asp?name_key=%C5%A3%D7%D0%BF%E3">ţ�п�</a> &nbsp;&nbsp;<a href="product.asp?name_key=%D0%DD%CF%D0%D0%AC">����Ь</a> &nbsp;&nbsp;</DIV>
	<form id="search" name="search" method="get" action="product.asp" onSubmit="return checkSearchForm()">
        <DIV class=sov>
        <INPUT id="name_key" value="<%if trim(request("name_key"))="" then %>������Ʒ�ؼ���<%else response.Write(trim(request("name_key"))) end if%>" onclick="javascript:this.value='';" name="name_key" onfocus="if(this.value=='������Ʒ�ؼ���') this.value='';" onBlur="if(this.value=='') this.value='������Ʒ�ؼ���';" class="searchtxt" type="text"> <input class="searchbtn" onclick="javascript:search(this);" type="submit" value="">
        </DIV>
	</form>
	</DIV>
</div>
	<div style="clear:both;"></div>
</div>
   <script type="text/javascript">
    
    <!--
    function checkSearchForm()
    {
		//alert('okkk��');
        if(document.getElementById('name_key').value.length < 1||document.getElementById('name_key').value=='������Ʒ�ؼ���')
        {
           alert("�����������ؼ��ʣ�");
            return false;
        }
        else
        {
			document.searchForm.submit();
            return true;
        }
    }
    -->
    
    </script>