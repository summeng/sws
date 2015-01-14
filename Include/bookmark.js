//调用方式：<script src="/Include/bookmark.js"></script>

function bookmark(){

        if(getCookie("bookmark")!="yes"){

                setCookie("bookmark","yes",1);window.external.AddFavorite('http://www.ip126.com/','天天网络科技');
//("bookmark","yes",1)这里数字代表多少天弹出一次，('http://www.ip126.com/','天天网络科技')改为你的网站域名和站名
        }
}
function setCookie(name,value,days){

        var exp=new Date();

        exp.setTime(exp.getTime() + days*24*60*60*1000);

        var arr=document.cookie.match(new RegExp("(^| )"+name+"=([^;]*)(;|$)"));

        document.cookie=name+"="+escape(value)+";expires="+exp.toGMTString();

}

function getCookie(name){

        var arr=document.cookie.match(new RegExp("(^| )"+name+"=([^;]*)(;|$)"));

        if(arr!=null){

                return unescape(arr[2]);

                return null;

        }
}  

function delCookie(name){

        var exp=new Date();

        exp.setTime(exp.getTime()-1);

        var cval=getCookie(name);

        if(cval!=null){

                document.cookie=name+"="+cval+";expires="+exp.toGMTString();

        }
}
document.write('<body onUnload="bookmark()">')