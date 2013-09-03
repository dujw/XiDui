<!--#include file="include/connect.asp" -->
<%
call aspsql()
'判断用户是否登陆
call checklogin()
Name   = request.cookies("buyok")("userid")
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title><%=sitename%></title>
<meta name="keywords" content="<%=sitedescription%>">
<meta name="description" content="<%=sitekeywords%>">
<link href="image/web_css/index.css" rel="stylesheet" type="text/css" />
</head>
<body>
<!--#include file="include/web_top.asp" -->
<div id="div_Public">
  <table width="950" border="0" align="center" cellpadding="0" cellspacing="0">
    <tr>
      <td width="260" align="left" valign="top"><!--#include file="left.asp"--></td>
      <td valign="top"><TABLE cellSpacing=1 cellPadding=0 width="690" align=center bgColor=#c9c9c9 border=0>
          <TBODY>
            <TR>
              <TD style="PADDING-RIGHT: 4px; PADDING-LEFT: 4px; PADDING-BOTTOM: 4px; PADDING-TOP: 4px" align=left bgColor=#ffffff height=25><strong>数据包下载：</strong></TD>
            </TR>
            <TR>
              <TD style="PADDING-RIGHT: 4px; PADDING-LEFT: 4px; PADDING-BOTTOM: 4px; PADDING-TOP: 4px" bgColor=#eeeeee height=25><font color="#FF0000">【淘宝/天猫版】</font></TD>
            </TR>
            <TR>
              <TD style="PADDING-RIGHT: 4px; PADDING-LEFT: 4px; PADDING-BOTTOM: 4px; PADDING-TOP: 4px;color:#ff0000;" bgColor=#ffffff height=25><A href="http://img.xidui.net/shujubao/taobao_2013-08-18.rar" target="_blank">点击下载淘宝用完整数据包</A> ---最近更新日期 2013-08-18 22:00:00 <IMG src="images/new.gif"></TD>
            </TR>
            <TR>
              <TD style="PADDING-RIGHT: 4px; PADDING-LEFT: 4px; PADDING-BOTTOM: 4px; PADDING-TOP: 4px;color:#ff0000;" bgColor=#ffffff height=25><A href="http://img.xidui.net/shujubao/taobao_2013-07-10.rar" target="_blank">点击下载淘宝用完整数据包</A> ---最近更新日期 2013-07-10 09:00:00 </TD>
            </TR>
            <TR>
              <TD style="PADDING-RIGHT: 4px; PADDING-LEFT: 4px; PADDING-BOTTOM: 4px; PADDING-TOP: 4px;color:#ff0000;" bgColor=#ffffff height=25><A href="http://img.xidui.net/shujubao/taobao_2013-05-31.rar" target="_blank">点击下载淘宝用完整数据包</A> ---最近更新日期 2013-05-31 14:59:00 </TD>
            </TR>
           <TR>
              <TD style="PADDING-RIGHT: 4px; PADDING-LEFT: 4px; PADDING-BOTTOM: 4px; PADDING-TOP: 4px;color:#ff0000;" bgColor=#ffffff height=25><A href="http://img.xidui.net/shujubao/taobao_2013-05-22.rar" target="_blank">点击下载淘宝用完整数据包</A> ---最近更新日期 2013-05-22 23:00:00 </TD>
            </TR>
            <TR>
              <TD style="PADDING-RIGHT: 4px; PADDING-LEFT: 4px; PADDING-BOTTOM: 4px; PADDING-TOP: 4px;color:#ff0000;" bgColor=#ffffff height=25><A href="http://img.xidui.net/shujubao/taobao_2013-05-11.rar" target="_blank">点击下载淘宝用完整数据包</A> ---最近更新日期 2013-05-12 13:00:00</TD>
            </TR>
            <TR>
              <TD style="PADDING-RIGHT: 4px; PADDING-LEFT: 4px; PADDING-BOTTOM: 4px; PADDING-TOP: 4px;color:#ff0000;" bgColor=#ffffff height=25><A href="http://img.xidui.net/shujubao/taobao_2013-05-03.rar" target="_blank">点击下载淘宝用完整数据包</A> ---最近更新日期 2013-05-06 14:00:00 </TD>
            </TR>
            <TR>
              <TD style="PADDING-RIGHT: 4px; PADDING-LEFT: 4px; PADDING-BOTTOM: 4px; PADDING-TOP: 4px;color:#ff0000;" bgColor=#ffffff height=25><A href="http://img.xidui.net/shujubao/taobao_2013-04-09.rar" target="_blank">点击下载淘宝用完整数据包</A> ---最近更新日期 2013-04-17 09:50:00 </TD>
            </TR>
            <TR>
              <TD style="PADDING-RIGHT: 4px; PADDING-LEFT: 4px; PADDING-BOTTOM: 4px; PADDING-TOP: 4px;color:#ff0000;" bgColor=#ffffff height=25><A href="http://img.xidui.net/shujubao/taobao_dongji.rar" target="_blank">点击下载淘宝用完整数据包（冬装清仓）</A> ---最近更新日期 2013-04-17 09:50:00 </TD>
            </TR>
                        <TR>
              <TD style="PADDING-RIGHT: 4px; PADDING-LEFT: 4px; PADDING-BOTTOM: 4px; PADDING-TOP: 4px" bgColor=#eeeeee height=25><font color="#FF0000">【腾讯拍拍版】</font></TD>
            </TR>
           <TR>
              <TD style="PADDING-RIGHT: 4px; PADDING-LEFT: 4px; PADDING-BOTTOM: 4px; PADDING-TOP: 4px;color:#ff0000;" bgColor=#ffffff height=25><A href="http://img.xidui.net/shujubao/paipai_2013-08-18.rar" target="_blank">点击下载拍拍用完整数据包</A> ---最近更新日期 2013-08-18 23:00:30 <IMG src="images/new.gif"></TD>
            </TR>
            <TR>
              <TD style="PADDING-RIGHT: 4px; PADDING-LEFT: 4px; PADDING-BOTTOM: 4px; PADDING-TOP: 4px;color:#ff0000;" bgColor=#ffffff height=25><A href="http://img.xidui.net/shujubao/paipai_2013-05-31.rar" target="_blank">点击下载拍拍用完整数据包</A> ---最近更新日期 2013-05-31 15:00:30 </TD>
            </TR>
            <TR>
              <TD style="PADDING-RIGHT: 4px; PADDING-LEFT: 4px; PADDING-BOTTOM: 4px; PADDING-TOP: 4px;color:#ff0000;" bgColor=#ffffff height=25><A href="http://img.xidui.net/shujubao/paipai_2013-05-22.rar" target="_blank">点击下载拍拍用完整数据包</A> ---最近更新日期 2013-05-22 23:00:30 </TD>
            </TR>
             <TR>
              <TD style="PADDING-RIGHT: 4px; PADDING-LEFT: 4px; PADDING-BOTTOM: 4px; PADDING-TOP: 4px;color:#ff0000;" bgColor=#ffffff height=25><A href="http://img.xidui.net/shujubao/paipai_2013-05-11.rar" target="_blank">点击下载拍拍用完整数据包</A> ---最近更新日期 2013-05-12 13:00:00</TD>
            </TR>
            <TR>
              <TD style="PADDING-RIGHT: 4px; PADDING-LEFT: 4px; PADDING-BOTTOM: 4px; PADDING-TOP: 4px;color:#ff0000;" bgColor=#ffffff height=25><A href="http://img.xidui.net/shujubao/paipai_2013-05-03.rar" target="_blank">点击下载拍拍用完整数据包</A> ---最近更新日期 2013-05-06 14:00:00 </TD>
            </TR>
            <TR>
              <TD style="PADDING-RIGHT: 4px; PADDING-LEFT: 4px; PADDING-BOTTOM: 4px; PADDING-TOP: 4px;color:#ff0000;" bgColor=#ffffff height=25><A href="http://img.xidui.net/shujubao/paipai_2013-04-09.rar" target="_blank">点击下载拍拍用完整数据包</A> ---最近更新日期 2013-04-17 09:50:00 </TD>
            </TR>
            <TR>
              <TD style="PADDING-RIGHT: 4px; PADDING-LEFT: 4px; PADDING-BOTTOM: 4px; PADDING-TOP: 4px;color:#ff0000;" bgColor=#ffffff height=25><A href="http://img.xidui.net/shujubao/paipai_dongji.rar" target="_blank">点击下载拍拍用完整数据包（冬装清仓）</A> ---最近更新日期 2013-04-17 09:50:00 </TD>
            </TR>
          </TBODY>
        </TABLE></td>
    </tr>
  </table>
</div>
<!--#include file="include/web_boot.asp" -->
</body>
</html>
<% call CloseConn() %>
