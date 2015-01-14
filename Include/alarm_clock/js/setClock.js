//电脑家园－－www.ip126.com
//预置用户设定参数
var oldTabId = "1";
var preMusic = { '1':'梁山伯与祝英台' , '2':'风中有朵雨做的云' , '3':'四季之春天'  , '4':'江南春早','5':'秋天不回来' ,'6':'该死的温柔' ,'7':'菊花台' ,'8':'童话' ,'9':'夜曲' ,'10':'土耳其进行曲','11':'致爱丽丝' ,'12':'爱的罗曼史' ,'13':'莫斯科郊外的晚上' ,'14':'泰坦尼克号' ,'15':'雨中旋律' }
var preRepeat = {'week':'周' , 'month':'月' , 'quarter':'季度'  , 'halfYear':'半年'  , 'year':'一年' }
var vdomain = "";
var setting_clock = false;
var newD = new Date();

var server_hour = newD.getHours();
var server_m = newD.getMinutes();


d_preHourValue = { '1':server_hour , '2':server_hour , '3':server_hour } ;
d_preMinutesValue = { '1':server_m , '2':server_m , '3':server_m } ;
d_preIsRepeatValue = {'1':'0' , '2':'0' , '3':'0' };
d_preRepeatValue = { '1':'week' , '2':'week' , '3':'week' } ;
d_preKgValue = { '1':'0' , '2':'0' , '3':'0' } ;
d_preTitleValue = { '1':'不早了，该休息了！' , '2':'得陪女友逛街了！' , '3':'球赛直播开始了！' } ;
d_preMidValue = { '1':'1' , '2':'2' , '3':'4' } ;
var preHourValue = [] , preMinutesValue = [] , preIsRepeatValue = [] , preRepeatValue = [] , preKgValue = [] , preTitleValue = [] , preMidValue = [];

function tabChange(tabId)
{
	if(oldTabId==tabId)return;
	SetCookie( 'clock_hotTab' , tabId );
	var tb = document.getElementById('tab'+tabId);
	var b = document.getElementById('tabBody'+tabId);
	if( oldTabId !='' )
	{
		var otb = document.getElementById('tab'+oldTabId);
		var ob = document.getElementById('tabBody'+oldTabId);
		otb.innerHTML = "<a href='###'><img onclick=\"tabChange('"+oldTabId+"');\" src='i/clock_"+oldTabId+".gif'></a>";
		ob.style.display = "none";
		oldTabId = tabId;
	}
	tb.innerHTML = "<a href='###'><img onclick=\"tabChange('"+tabId+"');\" src='i/clock_cur_"+tabId+".gif'></a>";
	b.style.display="block";
}
 
function musicSelectList(formId)
{
	var str = '<select name="music" onchange="loadMusic(this.value,'+formId+')">';
	for( var v in preMusic )
	{
		str += '<option value="'+v+'"';
		if( preMidValue[formId]==v)
			str +=' selected';
		str +='>'+preMusic[v]+'</option>';
	}
	str +='</select>';
	//alert(str);
	document.write( str );
}

function numSelectList( minN , maxN , selectName , preValue )
{
	var str = '<select name="'+selectName+'">';
	for( var i=minN;i<=maxN;i++ )
	{
		str += '<option value="'+i+'"';
		if( preValue==i)
			str +=' selected';
		str +='>'+i+'</option>';
	}
	str +='</select>';
	document.write( str );
}


function hourSelectList(formId)
{
    numSelectList(0,23,'hour',preHourValue[formId]);
}

function minutesSelectList(formId)
{
    numSelectList(0,59,'minutes',preMinutesValue[formId]);
}

function repeatSelectList(formId)
{
	if( preRepeatValue[formId]!='week' )
	{
		var fm = eval ("document.clockForm"+formId);
		fm.isRepeat[1].checked = true;
	}
	var str = '<select name="repeat">';
	for( var v in preRepeat )
	{
		str += '<option value="'+v+'"';
		if( preRepeatValue[formId]==v)
			str +=' selected';
		str +='>'+preRepeat[v]+'</option>';
	}
	str +='</select>';
	document.write( str );
}


function SetCookie( name , v ) 
{ 
	var argv = SetCookie.arguments;
	var argc = SetCookie.arguments.length; 
	var expires = (argc > 2) ? argv[2] : null; 
	var path = (argc > 3) ? argv[3] : null; 
	var domain = (argc > 4) ? argv[4] : null; 
	var secure = (argc > 5) ? argv[5] : false; 

	var v = escape( v );
	var vcook =name+"="+v;
	if(typeof(expires)=='object')vcook +=  (expires == null) ? "" : "; expires=" + expires.toGMTString();
	else if(expires>0)
	{
		var tmpD = new Date(expires*1000);
		vcook += "; expires=" + tmpD.toGMTString();
		//alert( tmpD.toGMTString() );
	}
	if(path!='')vcook +=  (path == null) ? "" : ("; path=" + path) ;
	if(domain!='')vcook +=  (domain == null) ? "" : ("; domain=" + domain) ; 
	vcook +=  (secure == true) ? "; secure" : "";
	document.cookie = vcook;
}
 

function GetCookie(sName)
{
	var aCookie = document.cookie.split("; ");
	for (var i=0; i < aCookie.length; i++)
	{
		var aCrumb = aCookie[i].split("=");
		if (sName == aCrumb[0])
		return unescape(aCrumb[1]);
	}
	return null;
}

function GetCookieObj()
{
	var o = [];
	var aCookie = document.cookie.split("; ");
	for (var i=0; i < aCookie.length; i++)
	{
		var aCrumb = aCookie[i].split("=");
		o[aCrumb[0]] = aCrumb[1];
	}
	return o;
}

function DeleteCookie (name) 
{ 
	var exp = new Date(); 
	exp.setTime (exp.getTime() - 100000 ); 
	SetCookie( name , 0 , exp , "/" , vdomain );
}

//获取用户设定参数
function setPreValue()
{
	for( var i=1;i<=3;i++ )
	{
		var fm = eval ("document.clockForm"+i);
		fm.title.value = preTitleValue[i];
		fm.kg.value = preKgValue[i];

		if(preRepeatValue[i]=='0' || preIsRepeatValue[i]=='0')
			fm.isRepeat[0].checked = true;
		
		if(fm.kg.value=='1')
		{
			lockForm( i , true , true );
		}
		else
		{
			lockForm( i , false , true );			
		}
	}
}

function lockForm( formId , lock , isInit )
{
	var fm = eval ("document.clockForm"+formId);
	fm.hour.disabled = lock;
	fm.minutes.disabled = lock;
	fm.music.disabled = lock;
	fm.repeat.disabled = lock;
	fm.title.disabled = lock;
	fm.isRepeat[0].disabled = lock;
	fm.isRepeat[1].disabled = lock;

	var vexp = new Date();
	ot = vexp.getTime()+365*86400000; //毫秒
	vexp.setTime( ot );
	if(lock==true)
	{
		SetCookie( 'clock_kg_'+formId , 1 , vexp , "/" , vdomain );
		SetCookie( 'clock_run_'+formId , '-1' , vexp , "/" , vdomain );
		if(!isInit)
		{
			fm.kg.value = '1';
		}
		fm.kgTitle.value = '取消闹钟';
	}
	else 
	{
		fm.kgTitle.value = '开启闹钟';
		if( !isInit )
		{
			fm.kg.value = '0';
		}
		SetCookie( 'clock_kg_'+formId , 0 , vexp , "/" , vdomain );
	}
    //document.getElementById('clockParams'+formId).disabled = true;
}

function getUserSetting()
{
	var co = GetCookieObj();
 	for( var i=1;i<=3;i++ )
	{
		//alert((co['clock_kg_'+i]));
		var clock_title=co['clock_t_'+i];
		if(clock_title!='undefined')
		{
			clock_title =  decodeURIComponent( unescape(clock_title) );
		}
		preHourValue[i] = ( typeof(co['clock_h_'+i])=="undefined" )?d_preHourValue[i]:co['clock_h_'+i];
		preMinutesValue[i] = ( typeof(co['clock_m_'+i])=="undefined" )?d_preMinutesValue[i]:co['clock_m_'+i];
		preIsRepeatValue[i] = ( typeof(co['clock_isRepeat_'+i])=="undefined" )?d_preIsRepeatValue[i]:co['clock_isRepeat_'+i];
		preRepeatValue[i] = ( typeof(co['clock_r_'+i])=="undefined" )?d_preRepeatValue[i]:co['clock_r_'+i];
		preKgValue[i] = ( typeof(co['clock_kg_'+i])=="undefined" )?d_preKgValue[i]:co['clock_kg_'+i];
		preTitleValue[i] = ( typeof( co['clock_t_'+i] )=="undefined" )?d_preTitleValue[i]:clock_title;
		preMidValue[i] = ( typeof( co['clock_mid_'+i] )=="undefined" )?d_preMidValue[i]:co['clock_mid_'+i];
		//alert( preKgValue[i] );
		//break;
	}
}
//设定闹钟
function setClock( formId )
{
	stopMusic();
	var fm = eval ( "document.clockForm"+formId );
	if(fm.kg.value=="1")
	{
		lockForm(formId,false,false);
		if( iTimerId[formId]!=null )clearTimeout( iTimerId[formId] );
		return;
	}

	fm.kgTitle.disabled = true;
	var kg_value = 1;
	var isRepeat = fm.isRepeat[0].checked?0:1;

	//获取cookie过期时间
	var vexp = new Date();
	var dexp = new Date();
	vdate = vexp.getFullYear() + "-"+(vexp.getMonth()+1)+"-"+vexp.getDate() ;
	vtime = vexp.getHours() + ":"+vexp.getMinutes()+":"+vexp.getSeconds() ;
	vTimeStamp = vexp.getTime()/1000;
	ot = vexp.getTime()+365*86400000;
	vexp.setTime( ot );
	dexp.setTime( dexp.getTime()-10000 );
	//系统默认值
	SetCookie( 'clock_userTimeStamp_'+formId , vTimeStamp , vexp , "/" , vdomain );
	SetCookie( 'clock_usertime_'+formId , vtime , vexp , "/" , vdomain );
	SetCookie( 'clock_userdate_'+formId , vdate , vexp , "/" , vdomain );
	SetCookie( 'clock_done_'+formId , '' , vexp , "/" , vdomain );
	SetCookie( 'clock_running' , '0' , dexp , "/" , vdomain );

	//用户值
	SetCookie( 'clock_kg_'+formId , kg_value , vexp , "/" , vdomain );
	SetCookie( 'clock_h_'+formId , fm.hour.value , vexp , "/" , vdomain );
	SetCookie( 'clock_m_'+formId , fm.minutes.value , vexp , "/" , vdomain );
	SetCookie( 'clock_mid_'+formId , fm.music.value , vexp , "/" , vdomain );
	SetCookie( 'clock_t_'+formId , encodeURIComponent( fm.title.value ) , vexp , "/" , vdomain );
	SetCookie( 'clock_isRepeat_'+formId , isRepeat , vexp , "/" , vdomain );
	SetCookie( 'clock_r_'+formId , fm.repeat.value , vexp , "/" , vdomain );
	SetCookie( 'clock_run_'+formId , '0' , vexp , "/" , vdomain );
	
	clockStart();
	//alert(document.cookie);
	window.status ='设置闹钟'+formId+"成功";
	lockForm(formId,true,false);
	fm.kgTitle.disabled = false;
}

function stopMusic()
{
	document.getElementById('clock_mid').src = "";
}

function playMusic( formId )
{
	var fm = eval ( "document.clockForm"+formId );
	var v = fm.testMusicButton.value;
	var mid = ((v=="试听")?fm.music.value:"unknown");

	document.getElementById('clock_mid').src = "i/mid/"+ mid +".mid";
	fm.testMusicButton.value=((v=="试听")?"停止":"试听");
}


//预载音乐
function loadMusic(mid,formId)
{
	//alert(formId);
	var fm = eval ( "document.clockForm"+formId );
	fm.testMusicButton.value = '试听';
	//alert( fm.testMusicButton.value );
	document.getElementById('clock_mid').src = "i/mid/"+ mid+".mid";
	document.getElementById('clock_mid').src = "";
}

getUserSetting();