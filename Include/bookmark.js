//���÷�ʽ��<script src="/Include/bookmark.js"></script>

function bookmark(){

        if(getCookie("bookmark")!="yes"){

                setCookie("bookmark","yes",1);window.external.AddFavorite('http://www.ip126.com/','��������Ƽ�');
//("bookmark","yes",1)�������ִ�������쵯��һ�Σ�('http://www.ip126.com/','��������Ƽ�')��Ϊ�����վ������վ��
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