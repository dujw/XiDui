<a target=_self href="javascript:goUrl()">
<script type="text/javascript">imgUrl1="<%=index_left_1%>";
imgtext1="宣传"
imgLink1="<%=index_left_url_1%>"
imgUrl2="<%=index_left_2%>";
imgtext2="联系我们"
imgLink2="<%=index_left_url_2%>"
imgUrl3="<%=index_left_3%>";
imgtext3="代理需知"
imgLink3="<%=index_left_url_3%>"
imgUrl4="<%=index_left_4%>";
imgtext4="退换说明"
imgLink4="<%=index_left_url_4%>"
imgUrl5="<%=index_left_5%>";
imgtext5="常见问题"
imgLink5="<%=index_left_url_5%>"

 var focus_width=710
 var focus_height=400
 var text_height=0
 var swf_height = focus_height+text_height
 
 var pics=imgUrl1+"|"+imgUrl2+"|"+imgUrl3+"|"+imgUrl4+"|"+imgUrl5
 var links=imgLink1+"|"+imgLink2+"|"+imgLink3+"|"+imgLink4+"|"+imgLink5
 var texts=imgtext1+"|"+imgtext2+"|"+imgtext3+"|"+imgtext4+"|"+imgtext5
 
 document.write('<object classid="clsid:d27cdb6e-ae6d-11cf-96b8-444553540000" codebase="http://fpdownload.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,0,0" width="'+ focus_width +'" height="'+ swf_height +'">');
 document.write('<param name="allowScriptAccess" value="sameDomain"><param name="movie" value="image/focus.swf"><param name="quality" value="high"><param name="bgcolor" value="#F0F0F0">');
 document.write('<param name="menu" value="false"><param name=wmode value="opaque">');
 document.write('<param name="FlashVars" value="pics='+pics+'&links='+links+'&texts='+texts+'&borderwidth='+focus_width+'&borderheight='+focus_height+'&textheight='+text_height+'">');
 document.write('<embed src="image/flash/focus.swf" wmode="opaque" FlashVars="pics='+pics+'&links='+links+'&texts='+texts+'&borderwidth='+focus_width+'&borderheight='+focus_height+'&textheight='+text_height+'" menu="false" bgcolor="#F0F0F0" quality="high" width="'+ focus_width +'" height="'+ focus_height +'" allowScriptAccess="sameDomain" type="application/x-shockwave-flash" pluginspage="http://www.macromedia.com/go/getflashplayer" />');  document.write('</object>');
 </script>
</a>