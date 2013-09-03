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
	$.dialog({width:700,height:280,opacity:0.5,background:'#ccc',lock:true,title: '裁难 <%=djw_rs(2)%> 发布',content: 'url:Notice.asp?id=<%=djw_rs(0)%>',button: [{ name: '关闭'}],padding:5});
}
</script>
<table width="950" border="0" align="center" cellpadding="0" cellspacing="0" style=" border:#CCCCCC solid 1px;">
  <tr >
    <td height="30" align="left"><strong><center><a href="#" onclick="dk_notice()">裁男最新通知:<font color="#FF0000"><%=djw_rs(1)%></font>&nbsp;&nbsp;[<%=djw_rs(2)%>]</a></center></strong>
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
<a href="../login.asp" target="_blank"  class="huang">登陆</a> | <a href="../reg.asp" target="_blank"  class="huang">注册</a>
<%else
	response.Write "你好，" & request.cookies("buyok")("userid") & "！  <a href='../login/logout.asp' class='huang'>[退出]<a> <a href='../user_center.asp' class='huang'>我的订单</a>  <a href='../check.asp' class='huang'>购物车</a>"
	end if%></div>
<a href="../user_sjb.asp"><strong>数据包下载</strong></a>
    <a target="_blank" href="http://wpa.qq.com/msgrd?v=3&uin=1504490396&site=裁难小刘&menu=yes"><img border="0" src="http://wpa.qq.com/pa?p=2:1504490396:45" alt="点击这里给我发消息" title="点击这里给我发消息">裁男小刘</a>
	<a target="_blank" href="http://wpa.qq.com/msgrd?v=3&uin=2262462307&site=裁难小周&menu=yes"><img border="0" src="http://wpa.qq.com/pa?p=2:2262462307:45" alt="点击这里给我发消息" title="点击这里给我发消息">裁男小周</a>
	<a target="_blank" href="http://wpa.qq.com/msgrd?v=3&uin=2404470991&site=裁难小杨&menu=yes"><img border="0" src="http://wpa.qq.com/pa?p=2:2404470991:45" alt="点击这里给我发消息" title="点击这里给我发消息">裁男小杨</a>
    <a target="_blank" href="http://wpa.qq.com/msgrd?v=3&uin=2486488687&site=裁难小许&menu=yes"><img border="0" src="http://wpa.qq.com/pa?p=2:2486488687:45" alt="点击这里给我发消息" title="点击这里给我发消息">裁男小许</a>
</div>
<div class="logoleftb">
 <a href="Contact.asp?id=12" target="_blank">帮助中心</a> | <a href="#" onclick="window.external.addFavorite('http://www.xidui.net/','西堆男装 外贸批发 网店货源提供 一件代发')">加入收藏</a> </div>
</div>
<div class="dhbj">
    <div class="dh">
        <div class="dhb">
            <div class="menu_box">
                <div class="menu" id="chromemenu">
                    <ul>
                              <%
                LarCode=trim(request("LarCode"))'产品大分类
                MidCode=trim(request("MidCode"))'产品小分类
                if LarCode = "" then
                    response.Write "<li class='m_2' onMouseOver=this.className='m_2_h' onMouseOut=this.className='m_2'><a href='../index.asp' class='cur'>网站首页</a></li>"
                    'response.Write "<td  width=85 height=30 align=center background=images/index/top1.jpg><a href=../index.asp class=titile_1>网站首页</a></td><td width=2></td>"
                else
                    'response.Write "<td  width=85 height=30 align=center background=images/index/top1.jpg><a class=text_1 href=../index.asp>网站首页</a></td><td width=2></td>"
                    response.Write "<li class='m_2' onMouseOver=this.className='m_2_h' onMouseOut=this.className='m_2'><a href='../index.asp' class='cur'>网站首页</a></li>"
                end if
                set rs_product=server.createobject("adodb.recordset") 
                sql_product="select distinct LarCode,LarSeq from dk501_productclass order by LarSeq" 
                rs_product.Open sql_product,conn,1,1 		
                if rs_product.bof and rs_product.eof then
                    response.write "数据正在更新中，请耐心等候！"
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
                'response.Write "<td  width=85 height=30 align=center background=images/index/top1.jpg><a class=text_1 href='xuanchuan.asp'><span style='color:#F00; font-weight:bold;'>PSD源文件</span></a></td><td width=2></td>"
                response.Write "<li class='m_2' onMouseOver=this.className='m_2_h' onMouseOut=this.className='m_2'><a href=xuanchuan.asp  class='cur'><span style='color:#F00; font-weight:bold;'>PSD源文件</span></a></li>"
                %>
                                </ul>
                </div>
            </div>
        </div>
    <!--<div class="dhgouwu"><div class="dgwz">&nbsp;<a href="flow.php" title="查看购物车"><a href=’flow.php‘>购物车有<span id=prdCount class=ba10000> 0 </span>件商品</a></a><a href="flow.php" class="huangc" target="_blank">&nbsp;<img src="themes/ok-moban_langrand/images/jsbtn.gif"></a></div>
    </div>-->
    </div>
</div>
	<div class="so">
	<DIV class=sogjc>&nbsp;&nbsp;&nbsp;&nbsp;<B>热门搜索词：</B>
<a href="product.asp?name_key=2012%D0%C2%BF%EE">2012新款</a> &nbsp;&nbsp;<a href="product.asp?name_key=T%D0%F4">T恤</a> &nbsp;&nbsp;<a href="product.asp?name_key=%B3%C4%D2%C2">衬衣</a> &nbsp;&nbsp;<a href="product.asp?name_key=%C5%A3%D7%D0%BF%E3">牛仔裤</a> &nbsp;&nbsp;<a href="product.asp?name_key=%D0%DD%CF%D0%D0%AC">休闲鞋</a> &nbsp;&nbsp;</DIV>
	<form id="search" name="search" method="get" action="product.asp" onSubmit="return checkSearchForm()">
        <DIV class=sov>
        <INPUT id="name_key" value="<%if trim(request("name_key"))="" then %>输入商品关键字<%else response.Write(trim(request("name_key"))) end if%>" onclick="javascript:this.value='';" name="name_key" onfocus="if(this.value=='输入商品关键字') this.value='';" onBlur="if(this.value=='') this.value='输入商品关键字';" class="searchtxt" type="text"> <input class="searchbtn" onclick="javascript:search(this);" type="submit" value="">
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
		//alert('okkk！');
        if(document.getElementById('name_key').value.length < 1||document.getElementById('name_key').value=='输入商品关键字')
        {
           alert("请输入搜索关键词！");
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