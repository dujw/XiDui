<!--#include file="include/connect.asp" -->
<%
call aspsql()
'�ж��û��Ƿ��½
call checklogin()
Name   = request.cookies("buyok")("userid")
		set rsadd=server.createobject("adodb.recordset")
		sql="select * from dk501_border where UserId='"& Name &"' and OrderNum='"& Name &"'"
		rsadd.open sql,conn,1,1
		if (rsadd.bof and rsadd.eof) then
			response.redirect "../error/error.asp?error=007"
			response.end
		end if

%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title><%=sitename%></title>
<meta name="keywords" content="<%=sitedescription%>">
<meta name="description" content="<%=sitekeywords%>">
<link href="image/web_css/index.css" rel="stylesheet" type="text/css" />
<script type="text/javascript">
function $(){var elements=new Array();for(var i=0;i<arguments.length;i++) {var element=arguments[i];if(typeof(arguments[i])=='string'){element=document.getElementById(arguments[i]);}if(arguments.length==1){return element;}elements.Push(element);}return elements;} 
//Array.prototype.Push=function(){var startLength=this.length;for(var i=0;i<arguments.length;i++){this[startLength+i]=arguments[i];}return this.length;} 
//String.prototype.HexToDec=function(){return parseInt(this,16);} 
//String.prototype.cleanBlank=function(){return this.isEmpty()?"":this.replace(/\s/g,"");} 
//function checkColor(){var color_tmp=(arguments[0]+"").replace(/\s/g,"").toUpperCase();var model_tmp1=arguments[1].toUpperCase();var model_tmp2="rgb("+arguments[1].substring(1,3).HexToDec()+","+arguments[1].substring(1,3).HexToDec()+","+arguments[1].substring(5).HexToDec()+")";model_tmp2=model_tmp2.toUpperCase();if(color_tmp==model_tmp1 ||color_tmp==model_tmp2){return true;}return false;} 
function $V(){return $(arguments[0]).value;} 
function do_add(n){
	$('sumkiss').value=eval($('youfei').value)+n;
	$('sumkiss_text').innerText=eval($('youfei').value)+n;
	}
</script>
</head>
<body>
<!--#include file="include/web_top.asp" -->
<div id="div_Public">
  <table width="950" border="0" align="center" cellpadding="0" cellspacing="0">
    <tr>
      <td height="40" background="image/product/list.jpg" style="padding-left:200px;">
	  <a href="/index.asp" class="text">��ҳ</a>&nbsp;|&nbsp;
	  <%
	  SubLar=trim(request("LarCode"))
	  set rs_da=server.createobject("adodb.recordset") 
	  sql_da="select LarCode from dk501_productclass where MidSeq = 1 order by LarSeq asc" 
	  rs_da.Open sql_da,conn,1,1
	  if rs_da.bof and rs_da.eof then
		  response.write "��ʱû�з���"
	  else
	  	  Do While Not rs_da.eof	
	  %>
	  <a href="/product/product.asp?LarCode=<%=Server.URLEncode(Trim(rs_da("LarCode")))%>" class=text><%=rs_da("LarCode")%></a>&nbsp;|&nbsp;<%
	  rs_da.movenext
		  if rs_da.eof then exit do
		loop		
	end if
	rs_da.close
	set rs_da=nothing
	  %>	  </td>
    </tr>
    <tr>
      <td><img src="../adv/list.jpg" width="950" height="90" /></td>
    </tr>
    <tr>
      <td height="5"></td>
    </tr>
    <tr>
      <td align="center" class="b-top b-left b-right b-bottom" style="padding:6px 6px 6px 6px;">
	  <table width="900" border="0" align="center" cellpadding="0" cellspacing="0">
<tr>
        <td height="65" align="left" background="image/order/news2.jpg" class="b-left b-right">&nbsp;&nbsp;<img src="image/order/user1.gif" width="11" height="11" /> ��ǰλ�ã�<a href="index.asp">��ҳ</a> ��<a href="shop.asp">��������</a> ���ҵĶ���</td>
      </tr>
      <tr>
        <td height="28" align="left" valign="middle" background="image/order/bucou1.jpg" class="b-left b-right">&nbsp;&nbsp;&nbsp;<font color="#FFFFFF">�����ύ���1.ѡ���Ʒ������Ʒ 2.��дҪ�ĳ��루��������������Ʒ��3.��д��ϵ��ʽ�Լ����� 4.�ɹ��ύ����</font></td>
  </tr>
      <tr>
        <td height="89" align="left" valign="top" class="b-left b-right b-bottom" style="padding-top:10px; padding-right:10px; padding-bottom:10px; padding-left:10px;">
          <table border="0" align="center" width="830" cellpadding="0" cellspacing="0">
              <tr align="center" bgcolor="#EEEEEE">
                <td width="837"><table width="100%"  border="0" cellpadding="0" cellspacing="1" bgcolor="#999999">
                  <tr bgcolor="#f0f0f0">
                    <td height="23" colspan="2" align="center">�� Ʒ �� ��</td>
                    <td width="65" height="23" align="center">�г���</td>
                    <td width="65" height="23" align="center">����۸�</td>
                    <td width="65" align="center">��Ʒ����</td>
                    <td width="55" align="center">�� ��</td>
                    <td width="59" height="23" align="center" nowrap="nowrap">����</td>
                    <td width="60" align="center" nowrap="nowrap">����С��</td>
                    <td width="94" height="23" align="center">С��</td>
                  </tr>
                  <%
					Set rs=conn.execute("select A.id,A.ProdId,A.ProdUnit,A.BuyPrice,A.xinhao,B.ProdName,B.ImgPrev,B.PriceOrigin,B.zhongliang,B.jifenzhi from dk501_border A,dk501_product B where A.ProdId = B.ProdId and  A.UserId='"& Name &"' and A.OrderNum='"& Name &"' order by A.ProdId")
					if rs.eof and rs.bof then
						response.redirect "../error/error.asp?error=007"
						response.end
					else
						Total = 0
						TotalA= 0
						TotalB= 0
						do while not rs.eof
						Sum = formatnum(csng(rs("BuyPrice"))*rs("ProdUnit"),2)
						Sum = FormatNumber(Sum,2)
						SumA = formatnum(csng(rs("zhongliang"))*rs("ProdUnit"),2)
						SumB = formatnum(csng(rs("jifenzhi"))*rs("ProdUnit"),2)
						Total = Sum + Total
						TotalA = SumA + TotalA
						TotalB = SumB + TotalB
						
						
					%>
                  <tr bgcolor="#FFFFFF">
                    <td width="53" height="52" align="center" valign="middle"><img src="/<%=rs("ImgPrev")%>" width="50" height="50" /></td>
                    <td width="355" align="left" valign="middle"><a href='list.asp?ProdId=<%=rs("ProdId")%>'><%=rs("ProdName")%></a></td>
                    <td align="center" class="shich_cont"><%=FormatNum(Trim(rs("PriceOrigin")),2)%></td>
                    <td align="center" nowrap="nowrap"><%=FormatNum(Trim(rs("BuyPrice")),2)%>Ԫ</td>
                    <td align="center" nowrap="nowrap"><%=rs("zhongliang")%>��</td>
                    <td align="center"><input type="text" name="spec_<%=rs("ProdId")%>_<%=rs("xinhao")%>" class="uset_input" size="4" maxlength="4"  value="<%=rs("ProdUnit")%>" /></td>
                    <td align="center"><%=rs("xinhao")%></td>
                    <td align="center"><%=rs("zhongliang")%>��</td>
                    <td align="center" nowrap="nowrap"><font color="#FF0000"><b><%=FormatNum(Trim(rs("BuyPrice"))*rs("ProdUnit"),2)%></b></font></td>
                  </tr>
                  <%
				  rs.movenext
				  loop
				  rs.close
				  set rs=nothing
				  end if 
				  Response.Write TotalA&"@"&TotalB
				  %>
                  
                </table></td>
              </tr>
              <tr align="left">
                <td height="25" bgcolor="#ffffff">������ܼƣ�<font color="#FF0000">��</font><font color="#FF0000"><%=formatnumber(total,2)%></font> ���أ�<font color="#FF0000"><%=formatnumber(totalA,2)%></font><font color="#FF0000">��</font> ����û��֣�<%=totalB%><a href="check.asp"></a></td>
                </tr>
              <tr align="left">
                <td height="18" bgcolor="#ffffff">&nbsp;</td>
              </tr>
        </table>
          <table id="bb"  cellspacing="0" cellpadding="0" width="830" align="center" border="0">
            <form action="preview.asp" method="post" name="order" id="order"  onsubmit="return CheckForm();">
              <tr>
                <td height="2"></td>
              </tr>
              <tr bgcolor="#d1d2d3">
                <td colspan="3" height="4"></td>
              </tr>
              <tr>
                <td width="4" bgcolor="#d1d2d3"></td>
                <td width="768"><table cellspacing="0" cellpadding="0" width="100%" border="0">
                    <tr>
                      <td valign="top" width="20"><img height="48" src="image/order/800buy_car_arrow.gif" width="20" /></td>
                      <td width="361" height="25"><img height="31" src="image/order/800buy_car_songhuo.gif" width="242" /></td>
                      <td width="379">&nbsp;</td>
                    </tr>
                  </table>
                    <table cellspacing="0" cellpadding="0" width="109%" border="0">
                      <tr>
                        <td width="10">&nbsp;</td>
                        <td width="750"><table cellspacing="0" cellpadding="0" width="100%" border="0">
                            <tr>
                              <td><table height="20" cellspacing="0" cellpadding="0" width="780" border="0">
                                  <tr>
                                    <td width="103" height="25" align="right"><font color="#9d282f">*</font>�ջ��˵�������</td>
                                    <td width="241" height="25"><input name="Recname2"  type="text" id="Recname2" size="28" maxlength="30" />
                                        <font color="#9d282f">&nbsp;</font></td>
                                    <td width="70" height="25" nowrap="nowrap">���ʽ��</td>
                                    <td width="366" height="25" class="800_title">����վ�Ƽ��òƸ�ͨ����
                                      <input name="t1" type="hidden" onBlur="javascript:a();" value="<%=FormatNum(total,2)%>" size="2" maxlength="2"/></td>
                                  </tr>
                                  <tr>
                                    <td height="25" align="right" nowrap="nowrap"><font color="#9d282f">&nbsp;</font><font color="#9d282f">*</font>�ջ��˵ĵ绰��</td>
                                    <td height="25"><input name="RecPhone2" type="text" id="RecPhone2" size="28" maxlength="30" /></td>
                                    <td height="25" nowrap="nowrap">�ͻ���ʽ��</td>
                                    <td height="25"><%if totalA >1001 then%>
                          <input onClick="youfei.value='<%=Cint(left(totalA,len(totalA)-3))*6+10%>';do_add(<%=FormatNum(total,2)%>);" type="radio" name="kuaidi" value="��ͨ" checked="checked"/>
                          ��ͨ
                          <input onClick="youfei.value='<%=Cint(left(totalA,len(totalA)-3))*6+13%>';do_add(<%=FormatNum(total,2)%>);" type="radio" name="kuaidi" value="Բͨ" />
                          Բͨ
                          <input onClick="youfei.value='<%=Cint(left(totalA,len(totalA)-3))*12+20%>';do_add(<%=FormatNum(total,2)%>);" type="radio" name="kuaidi" value="˳��" />
                          ˳��
                          <input onClick="youfei.value='<%=Cint(left(totalA,len(totalA)-3))*6+15%>';do_add(<%=FormatNum(total,2)%>);" type="radio" name="kuaidi" value="EMS"  />
                          EMS <span class="800_title"> �ʷѣ� </span>
                          <input name="youfei" type="text" class="buttonface" id="youfei" onBlur="javascript:a();" onKeyUp="value=value.replace(/[^\d]/g,'') " size="6" maxlength="16onbeforepaste=&quot;clipboardData.setData('text',clipboardData.getData('text').replace(/[^\d]/g,''))&quot;" value="<%=Cint(left(totalA,len(totalA)-3))*6+10%>"/>
                          <% else %>
						  <input onClick="youfei.value='10';" type="radio" name="kuaidi" value="��ͨ" checked="checked"/>
                          ��ͨ
						  <input onClick="youfei.value='13';" type="radio" name="kuaidi" value="Բͨ" />
                          Բͨ
                          <input onClick="youfei.value='20';" type="radio" name="kuaidi" value="˳��" />
                          ˳��
                          <input onClick="youfei.value='15';"type="radio" name="kuaidi" value="EMS" />
                          EMS <span class="800_title"> �ʷѣ� </span>
                          <input name="youfei" type="text" class="buttonface" id="youfei" onBlur="javascript:a();" onKeyUp="value=value.replace(/[^\d]/g,'') " size="6" maxlength="16onbeforepaste=&quot;clipboardData.setData('text',clipboardData.getData('text').replace(/[^\d]/g,''))&quot;" value=10>
                          <% end if %>
                                                         </td>
                                  </tr>
                                  <tr>
                                    <td height="25" align="right" nowrap="nowrap"><font color="#9d282f">*</font>�ջ��˵ĵ�ַ��</td>
                                    <td height="25" colspan="2"><input name="Recaddress2" type="text" id="Recaddress2" value="" size="38" /></td>
                                    <td rowspan="2" align="left" bgcolor="#D9ECC6" class="b-top b-left b-right b-bottom"><input name="alipay" type="radio" value="�Ƹ�ͨ" checked="checked" />
                                      �Ƹ�ͨ<span class="800_hzp_red">(86851270 ŷ��)</span> <a href="http://www.tenpay.com" target="_blank" class="red">�������Ƹ�ͨ</a><br />
                                      <input name="alipay" type="radio" value="֧����" />
                                      ֧����<span class="800_left_red">��oujing915@163.com ŷ����</span><a href="http://www.alipay.com" target="_blank" class="red">�������֧����</a><br />
                                      <input type="radio" name="alipay" value="����" />
                                      ���У�����ϵ������Ա��</td>
                                  </tr>
                                  <tr>
                                    <td align="right" nowrap="nowrap">�ջ��˵��ʱࣺ</td>
                                    <td><input  type="text" name="ZipCode2" size="6" maxlength="6" /></td>
                                    <td nowrap="nowrap">���ʽ��</td>
                                  </tr>
                                  <tr>
                                    <td height="25" align="right" nowrap="nowrap">&nbsp;�ջ��˵��Ա�</td>
                                    <td height="25"><select class="buttonface" id="select19" size="1" name="t_sex">
                                        <option value="M" selected="selected">����</option>
                                        <option value="F">Ůʿ</option>
                                    </select></td>
                                    <td height="25" nowrap="nowrap" class="800_hzp_red">���׺��룺</td>
                                    <td height="25" class="800_hzp_red"><input name="Transaction" type="text" id="Transaction" size="30" maxlength="30" />
                                        </label>
                                      ��Ӧ���
                                      <span name="sumkiss_text" id="sumkiss_text" style=" color:#F00; font-weight:bold;"><%'=Cint(left(totalA,len(totalA)-3))*6+10+FormatNum(total,2)%></span> Ԫ</td>
                                  </tr>
                                  <tr>
                                    <td align="right" nowrap="nowrap">&nbsp;������ϸ˵����<br /></td>
                                    <td colspan="3"><span class="titile">����ֻ�ṩ���ƺ�ճ���ã����ԶԴ˶��������κ����ã�������ע��.-���ս���Ȩ�鱾վ����</span><br />
                                        <textarea  name="Notes" cols="80" rows="3"></textarea></td>
                                  </tr>
                                  <input type="hidden" name="Notes" />
                                  <tr>
                                    <td colspan="4" align="left" nowrap="nowrap"><input name="Submit3" type="image" height="42" width="105" src="image/order/800buy_car_botton2.gif" />
                                        <input type="hidden"  name="gettime" size="30"  maxlength="30"  value='��' />
                                        <input type="hidden" name="pei" value="<%=pei%>" />
                                        <input type="hidden" name="fei" value="<%=fei%>" />
                                       <input type="hidden" name="sumkiss" id="sumkiss" value="<%=FormatNum(total,2)%>"  style="width:45px;" /> 
                                        <%'=total%>
                                        <input type="hidden" name="zongzhong" value="<%=TotalA%>" />
                                        <%'=TotalA%>
                                        <input type="hidden" name="TotalB" value="<%=TotalB%>" />
                                      <%'=TotalB%> </td>
                                  </tr>
                              </table></td>
                            </tr>
                        </table></td>
                      </tr>
                  </table></td>
                <td width="4" bgcolor="#d1d2d3"></td>
              </tr>
              <tr bgcolor="#d1d2d3">
                <td colspan="3" height="4"></td>
              </tr>
            </form>
          </table></td>
      </tr>
</table>
	  </td>
    </tr>
  </table>
</div>
<!--#include file="include/web_boot.asp" -->
</body>
</html>
<% call CloseConn() %>
