
function copyToClipBoard(){
var clipBoardContent="";
clipBoardContent+=document.title;
clipBoardContent+="\n";
clipBoardContent+=this.location.href;
window.clipboardData.setData("Text",clipBoardContent);
alert("���Ƴɹ�����ճ�������QQ/MSN/�ʼ��ȴ��Ƽ�����ĺ���");
}
