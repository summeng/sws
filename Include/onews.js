function ContentSize(size)
{
	var obj=document.all.BodyLabel;
	obj.style.fontSize=size+"px";
}

function checkspace(checkstr) {
  var str = '';
  for(i = 0; i < checkstr.length; i++) {
    str = str + ' ';
  }
  return (str == checkstr);
}
function check()
{
  if(checkspace(document.pinglunform.pinglunname.value)) {
	document.pinglunform.pinglunname.focus();
    alert("请填写您的姓名！");
	return false;
  }
  if(checkspace(document.pinglunform.pingluncontent.value)) {
	document.pinglunform.pingluncontent.focus();
    alert("请填写评论正文！");
	return false;
  }
   if(checkspace(document.pinglunform.yanzheng.value)) {
	document.pinglunform.yanzheng.focus();
    alert("请输入验证码！");
	return false;
  }
  
	  }

var currentpos,timer; 
function initializeScroll() { timer=setInterval("scrollwindow()",80);} 
function scrollclear(){clearInterval(timer);}
function scrollwindow() {currentpos=document.body.scrollTop;window.scroll(0,++currentpos);if (currentpos != document.body.scrollTop) sc();} 
document.onmousedown=scrollclear
document.ondblclick=initializeScroll