
function copyToClipBoard(){
var clipBoardContent="";
clipBoardContent+=document.title;
clipBoardContent+="\n";
clipBoardContent+=this.location.href;
window.clipboardData.setData("Text",clipBoardContent);
alert("复制成功，请粘贴到你的QQ/MSN/邮件等处推荐给你的好友");
}
