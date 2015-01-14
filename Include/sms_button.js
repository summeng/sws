   //按钮时间
/** * 倒计时函数
	* 需要在按钮上绑定单击事件
	* 如: <INPUT contentEditable=false value=发送短信 type=button data-cke-pa-onclick="setInterval('countDown(this,30)',1000);" data-cke-editable="1">  * 30代表秒数，需要倒计时多少秒可以自行更改  */
	function countDown(obj,second){
	// 如果秒数还是大于0，则表示倒计时还没结束
	if(second>=0){
	// 获取默认按钮上的文字
	if(typeof buttonDefaultValue === 'undefined' ){
	buttonDefaultValue =  obj.defaultValue;
	} 
	// 按钮置为不可点击状态
	obj.disabled = true;
   // 按钮里的内容呈现倒计时状态
   //这行是原来的，获取表单上默认按钮文字obj.value = buttonDefaultValue+'('+second+')';
   obj.value = '等待'+'('+second+')'+'后重新获取';
   // 时间减一
   second--; 
   // 一秒后重复执行
   setTimeout(function(){countDown(obj,second);},1000); 
   // 否则，按钮重置为初始状态
   }else{ 
   // 按钮置未可点击状态
   obj.disabled = false; 
   // 按钮里的内容恢复初始状态
   obj.value = buttonDefaultValue;
   }
   }