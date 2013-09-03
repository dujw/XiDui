// JavaScript Document
var userAgent = navigator.userAgent.toLowerCase();
var is_opera = userAgent.indexOf('opera') != -1 && opera.version();
var is_moz = (navigator.product == 'Gecko') && userAgent.substr(userAgent.indexOf('firefox') + 8, 3);
var is_ie = (userAgent.indexOf('msie') != -1 && !is_opera) && userAgent.substr(userAgent.indexOf('msie') + 5, 3);
var is_mac = userAgent.indexOf('mac') != -1;

if(is_moz && window.HTMLElement) {
    HTMLElement.prototype.__defineSetter__('outerHTML', function(sHTML) {
            var r = this.ownerDocument.createRange();
        r.setStartBefore(this);
        var df = r.createContextualFragment(sHTML);
        this.parentNode.replaceChild(df,this);
        return sHTML;
    });

    HTMLElement.prototype.__defineGetter__('outerHTML', function() {
        var attr;
        var attrs = this.attributes;
        var str = '<' + this.tagName.toLowerCase();
        for(var i = 0;i < attrs.length;i++){
            attr = attrs[i];
            if(attr.specified)
            str += ' ' + attr.name + '="' + attr.value + '"';
        }
        if(!this.canHaveChildren) {
            return str + '>';
        }
        return str + '>' + this.innerHTML + '</' + this.tagName.toLowerCase() + '>';
        });

    HTMLElement.prototype.__defineGetter__('canHaveChildren', function() {
        switch(this.tagName.toLowerCase()) {
            case 'area':case 'base':case 'basefont':case 'col':case 'frame':case 'hr':case 'img':case 'br':case 'input':case 'isindex':case 'link':case 'meta':case 'param':
            return false;
            }
        return true;
    });
    HTMLElement.prototype.click = function(){
        var evt = this.ownerDocument.createEvent('MouseEvents');
        evt.initMouseEvent('click', true, true, this.ownerDocument.defaultView, 1, 0, 0, 0, 0, false, false, false, false, 0, null);
        this.dispatchEvent(evt);
    }
}

function $(id) {
    return document.getElementById(id);
}
function doane(event) {
    e = event ? event : window.event;
    if(is_ie) {
        e.returnValue = false;
        e.cancelBubble = true;
    } else if(e) {
        e.stopPropagation();
        e.preventDefault();
    }
}

function fetchOffset(obj) {
    var left_offset = obj.offsetLeft;
    var top_offset = obj.offsetTop;
    while((obj = obj.offsetParent) != null) {
        left_offset += obj.offsetLeft;
        top_offset += obj.offsetTop;
    }
    return { 'left' : left_offset, 'top' : top_offset };
}

//得到一个定长的hash值， 依赖于 stringxor()
function hash(string, length) {
    var length = length ? length : 32;
    var start = 0;
    var i = 0;
    var result = '';
    filllen = length - string.length % length;
    for(i = 0; i < filllen; i++){
        string += "0";
    }
    while(start < string.length) {
        result = stringxor(result, string.substr(start, length));
        start += length;
    }
    return result;
}