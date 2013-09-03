var zoomobj = Array();
var zoomadjust;var zoomstatus = 1;

function zoom(obj, zimg) {
    if(!zoomstatus) {
        window.open(zimg, '', '');
        return;
    }
    if(!zimg) {
        zimg = obj.src;
    }
    if(!$('zoomimglayer_bg')) {
        div = document.createElement('div');
        div.id = 'zoomimglayer_bg';
        div.style.position = 'absolute';
        div.style.left = div.style.top = '0px';
        div.style.zIndex = '998';
        div.style.width = '100%';
        div.style.height = document.body.scrollHeight + 'px';
        div.style.backgroundColor = '#000';
        div.style.display = 'none';
        div.style.filter = 'progid:DXImageTransform.Microsoft.Alpha(opacity=50,finishOpacity=100,style=0)';
        div.style.opacity = 0.8;
        $('append_parent').appendChild(div);
        div = document.createElement('div');
        div.id = 'zoomimglayer';
        div.style.position = 'absolute';
        div.style.padding = 0;
        $('append_parent').appendChild(div);
    }
    zoomobj['srcinfo'] = fetchOffset(obj);
    zoomobj['srcobj'] = obj;
    zoomobj['zimg'] = zimg;
    zoomobj['id'] = 'zoom_' + Math.random();
    $('zoomimglayer').setAttribute('pid', obj.getAttribute('pid'));
    $('zoomimglayer').style.display = '';
    $('zoomimglayer').style.left = zoomobj['srcinfo']['left'] + 'px';
    $('zoomimglayer').style.top = zoomobj['srcinfo']['top'] + 'px';
    $('zoomimglayer').style.width = zoomobj['srcobj'].width + 'px';
    $('zoomimglayer').style.height = zoomobj['srcobj'].height + 'px';
    $('zoomimglayer').style.filter = 'progid:DXImageTransform.Microsoft.Alpha(opacity=40)';
    $('zoomimglayer').style.opacity = 0.4;
    $('zoomimglayer').style.zIndex = 999;
    $('zoomimglayer').innerHTML = '<table width="100%" height="100%" cellspacing="0" cellpadding="0"><tr><td align="center" valign="middle"><img src="plus/discuz/pic/loading.gif"></td></tr></table><div style="position:absolute;top:-100000px;display:none"><img id="' + zoomobj['id'] + '" src="' + zoomobj['zimg'] + '"></div>';
    setTimeout('zoomimgresize($(\'' + zoomobj['id'] + '\'))', 100);
    if(is_ie) {
        doane(event);
    }
}

var zoomdragstart = new Array();
var zoomclick = 0;
function zoomdrag(e, op) {
    if(op == 1) {
        zoomclick = 1;
        zoomdragstart = is_ie ? [event.clientX, event.clientY] : [e.clientX, e.clientY];
        zoomdragstart[2] = parseInt($('zoomimglayer').style.left);
        zoomdragstart[3] = parseInt($('zoomimglayer').style.top);
        doane(e);
    } else if(op == 2 && zoomdragstart[0]) {
        zoomclick = 0;
        var zoomdragnow = is_ie ? [event.clientX, event.clientY] : [e.clientX, e.clientY];
        $('zoomimglayer').style.left = (zoomdragstart[2] + zoomdragnow[0] - zoomdragstart[0]) + 'px';
        $('zoomimglayer').style.top = (zoomdragstart[3] + zoomdragnow[1] - zoomdragstart[1]) + 'px';
        doane(e);
    } else if(op == 3) {
        if(zoomclick) zoomclose();
        zoomdragstart = [];
        doane(e);
    }
}

function zoomST(c) {
    if($('zoomimglayer').style.display == '') {
        $('zoomimglayer').style.left = (parseInt($('zoomimglayer').style.left) + zoomobj['x']) + 'px';
        $('zoomimglayer').style.top = (parseInt($('zoomimglayer').style.top) + zoomobj['y']) + 'px';
        $('zoomimglayer').style.width = (parseInt($('zoomimglayer').style.width) + zoomobj['w']) + 'px';
        $('zoomimglayer').style.height = (parseInt($('zoomimglayer').style.height) + zoomobj['h']) + 'px';
        c++;
        if(c <= 5) {
            setTimeout('zoomST(' + c + ')', 1);
        } else {
            $('zoomimglayer').style.filter = 'progid:DXImageTransform.Microsoft.Alpha(opacity=100,style=0)';
            $('zoomimglayer').style.opacity = 1;
            zoomadjust = 1;
            $('zoomimglayer').style.filter = '';
            $('zoomimglayer_bg').style.display = '';
            $('zoomimglayer').innerHTML = '<div class="zoominner"><p><span class="right"><a href="' + zoomobj['zimg'] + '" class="imglink" target="_blank" title="���´��ڴ�">���´��ڴ�</a><a href="javascipt:;" onclick="zoomimgadjust(event, 1)" class="imgadjust" title="ʵ�ʴ�С">ʵ�ʴ�С</a><a href="javascript:;" onclick="zoomclose()" class="imgclose" title="�ر�">�ر�</a></span>����������ͼƬ</p><div id="zoomimgbox"><img id="zoomimg" style="cursor: move;" src="' + zoomobj['zimg'] + '" width="' + parseInt($('zoomimglayer').style.width) + '" height="' + parseInt($('zoomimglayer').style.height) + '"></div></div>';
            $('zoomimglayer').style.overflow = 'visible';
            $('zoomimglayer').style.width = (parseInt($('zoomimg').width < 300 ? 300 : parseInt($('zoomimg').width)) + 20) + 'px';
            $('zoomimglayer').style.height = (parseInt($('zoomimg').height) + 20) + 'px';
            if(is_ie){
                $('zoomimglayer').onmousewheel = zoomimgadjust;
            } else {
                $('zoomimglayer').addEventListener("DOMMouseScroll", zoomimgadjust, false);
            }
            $('zoomimgbox').onmousedown = function(event) {try{zoomdrag(event, 1);}catch(e){}};
            $('zoomimgbox').onmousemove = function(event) {try{zoomdrag(event, 2);}catch(e){}};
            $('zoomimgbox').onmouseup = function(event) {try{zoomdrag(event, 3);}catch(e){}};
        }
    }
}

function zoomimgresize(obj) {
    if(!obj.complete) {
        setTimeout('zoomimgresize($(\'' + obj.id + '\'))', 100);
        return;
    }
    obj.parentNode.style.display = '';
    zoomobj['zimginfo'] = [obj.width, obj.height];
    var r = obj.width / obj.height;
    var w = document.body.clientWidth * 0.95;
    w = obj.width > w ? w : obj.width;
    var h = w / r;
    var clientHeight = document.documentElement.clientHeight ? document.documentElement.clientHeight : document.body.clientHeight;
    var scrollTop = document.body.scrollTop ? document.body.scrollTop : document.documentElement.scrollTop;
    if(h > clientHeight) {
        h = clientHeight;
        w = h * r;
    }
    var l = (document.body.clientWidth - w) / 2;
    var t = h < clientHeight ? (clientHeight - h) / 2 : 0;
    t += + scrollTop;
    zoomobj['x'] = (l - zoomobj['srcinfo']['left']) / 5;
    zoomobj['y'] = (t - zoomobj['srcinfo']['top']) / 5;
    zoomobj['w'] = (w - zoomobj['srcobj'].width) / 5;
    zoomobj['h'] = (h - zoomobj['srcobj'].height) / 5;
    $('zoomimglayer').style.filter = '';
    $('zoomimglayer').innerHTML = '';
    setTimeout('zoomST(1)', 5);
}

function zoomimgadjust(e, a) {
    if(!a) {
        if(!e) e = window.event;
        if(e.altKey || e.shiftKey || e.ctrlKey) return;
        var l = parseInt($('zoomimglayer').style.left);
        var t = parseInt($('zoomimglayer').style.top);
        if(e.wheelDelta <= 0 || e.detail > 0) {
            if($('zoomimg').width <= 200 || $('zoomimg').height <= 200) {
                doane(e);return;
            }
            $('zoomimg').width -= zoomobj['zimginfo'][0] / 10;
            $('zoomimg').height -= zoomobj['zimginfo'][1] / 10;
            l += zoomobj['zimginfo'][0] / 20;
            t += zoomobj['zimginfo'][1] / 20;
        } else {
            if($('zoomimg').width >= zoomobj['zimginfo'][0]) {
                zoomimgadjust(e, 1);return;
            }
            $('zoomimg').width += zoomobj['zimginfo'][0] / 10;
            $('zoomimg').height += zoomobj['zimginfo'][1] / 10;
            l -= zoomobj['zimginfo'][0] / 20;
            t -= zoomobj['zimginfo'][1] / 20;
        }
    } else {
        var clientHeight = document.documentElement.clientHeight ? document.documentElement.clientHeight : document.body.clientHeight;
        var scrollTop = document.body.scrollTop ? document.body.scrollTop : document.documentElement.scrollTop;
        $('zoomimg').width = zoomobj['zimginfo'][0];$('zoomimg').height = zoomobj['zimginfo'][1];
        var l = (document.body.clientWidth - $('zoomimg').clientWidth) / 2;l = l > 0 ? l : 0;
        var t = (clientHeight - $('zoomimg').clientHeight) / 2 + scrollTop;t = t > 0 ? t : 0;
    }
    $('zoomimglayer').style.left = l + 'px';
    $('zoomimglayer').style.top = t + 'px';
    $('zoomimglayer_bg').style.height = t + $('zoomimglayer').clientHeight > $('zoomimglayer_bg').clientHeight ? (t + $('zoomimglayer').clientHeight) + 'px' : $('zoomimglayer_bg').style.height;
    $('zoomimglayer').style.width = (parseInt($('zoomimg').width < 300 ? 300 : parseInt($('zoomimg').width)) + 20) + 'px';
    $('zoomimglayer').style.height = (parseInt($('zoomimg').height) + 20) + 'px';
    doane(e);
}

function zoomclose() {
    $('zoomimglayer').innerHTML = '';
    $('zoomimglayer').style.display = 'none';
    $('zoomimglayer_bg').style.display = 'none';
}