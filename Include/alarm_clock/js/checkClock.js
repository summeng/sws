//电脑家园－－www.ip126.com
//本文件来执行闹钟
var vdomain = "";
var co = GetCookieObj();
var clock_running = false;
var CLOCK_FILE_ID = 'checkClock';
var clock_setting = false;
var clock_info = "";	//闹钟设好后，鼠标移动上去的提示文字
var clock_num = 0; //闹钟个数
var clock_originalInfo = "";
var time , server_year , server_month , server_day , server_hour , server_m , server_s , server_w , server_now_seconds , server_date;
var iTimerId = { '1':null,'2':null,'3':null };
var CLOCK_VERSION = "0.1Beta 2007-10-25";

function warnUser()
{
	co = GetCookieObj();
	for( var i=1;i<=3;i++ )
	{
		var cKey = "clock_kg_"+i;
		if( co[cKey]=="1" )
		{
			event.returnValue = "关闭该页会使闹钟失效，是否继续？";
		}
	}
}

//设置闹的时间
function afterDoClock(resp)
{
	//alert(resp);
	clock_running = false;
	co = GetCookieObj();
	for( var i=1;i<=3;i++ )
	{
		var cKey = "clock_run_"+i;
		var runTime = co[cKey];
		var isRepeat =  co["clock_isRepeat_"+i];
		//alert(runTime); 
		//window.status += runTime+" = ";	//test
		if( runTime>0 )
		{
			//break;
			//runTime = 1;
			runTime = runTime * 1000; //转化成毫秒
			var t = co["clock_t_"+i];
			t = decodeURIComponent( unescape(t) );
			var mid = co['clock_mid_'+i];
			if( iTimerId[i]!=null )clearTimeout( iTimerId[i] );
			iTimerId[i] = setTimeout( "alertClock( "+i+" , '"+t+"' , '"+mid+"' , '"+isRepeat+"' )" , runTime );
			//alert( i +"="+ iTimerId[i] );
			SetCookie( "clock_running" , "1" , '0' , "/" , vdomain );
			clock_running = true;
		}
	}
}

//setTimeout( "alertClock( 1 , 'xi ban' , 3 )" , 2000 );
//闹
function alertClock( formId , title , mid , isRepeat )
{
	document.getElementById('clock_mid').src = "i/mid/"+ mid +".mid";
	DeleteCookie( "clock_run_"+formId );
	if( co['clock_isRepeat_'+formId]=='0' )
	{
		//lockForm( formId , false , false );
		SetCookie( 'clock_kg_'+formId , 0 , 0 , "/" , vdomain );
	}
	window.focus();
	alert(title);
	stopMusic();
	if(isRepeat=='0')lockForm( formId , false , false );
}



//执行闹钟
function runClock( formId , nextrun , time , domain )
{
	 //alert( formId+"\n"+nextrun+"\n"+typeof(nextrun)+"\n"+time);
    SetCookie( "clock_run_"+formId , nextrun , time+86400*365 , "/" , domain );
	 if(nextrun>0)clock_running = true;
}

//将闹钟开关设为关闭
function stopClock( formId , time , domain )
{
   SetCookie( "clock_kg_"+formId , 0 , time+86400*365 , "/" , domain );
}

//将闹钟暂停
function pauseClock( formId , time , server_date , domain )
{
	runClock( formId , -1 , time , domain )
}

var repeatSeconds = { 'week':86400*7 , 'month':86400*30 , 'quarter':86400*120 , 'halfYear':86400*180 , 'year':86400*365 };

//开始执行闹钟：对比本地记录，取出时间
function doLocalClock()
{
	setAlarmTime();
	afterDoClock('');
}

//本地设置距离闹的时间
function setAlarmTime()
{
	//alert(document.cookie);
	clock_running = false;
	var co = GetCookieObj();
	var newD = new Date();
	var time = parseInt(newD.getTime()/1000);
	var userTimeStamp	//用户设定闹钟时的时间戳
	server_year = newD.getYear();
	server_month = newD.getMonth()+1;
	server_day = newD.getDate();
	server_hour = newD.getHours();
	server_m = newD.getMinutes();
	server_s = newD.getSeconds();
	server_w = newD.getDay();
	server_now_seconds = ( ( server_hour*60+server_m ) * 60 )+server_s;
	server_date = server_year +"-"+ server_month +"-"+server_day;

	for (var i=1;i<=3 ;i++ ) 
	{
		var formId = i;
		if( co["clock_kg_"+i]=='0' )
		{
			//alert('kg'+i);
			continue;
		}
		//alert(i);

		var h = parseInt(co['clock_h_'+i])*60;
		var m = parseInt( co['clock_m_'+i] );
		run_seconds = ( h+m )*60; 
		nextrun = run_seconds - server_now_seconds;
		//alert( server_hour +":"+server_m +":"+server_s +"\n"+ server_now_seconds +"\n\n"+co['clock_h_'+i]+"\n"+co['clock_m_'+i]+"\n"+ run_seconds +"\n"+nextrun);
		//alert(nextrun); 
		//break;
		//如果不重复
		if( co['clock_isRepeat_'+i]=='0' )
		{
			//时间已经过去，停掉闹钟
			if(nextrun<=0){ stopClock( i , time , vdomain ); continue; }
		}
		//如果重复
		else 
		{
			var repeat = co['clock_r_'+i];
			userTimeStamp = parseInt( co['clock_userTimeStamp_'+i] );
			//执行结束时间
			endTimeStamp =  userTimeStamp + repeatSeconds[repeat];
			//如果重复周期已经过去
			if( endTimeStamp<time )
			{
				 stopClock( i, time , vdomain );
				 break;
			}
			else 
			{
				//时间已经过去，暂停闹钟
				//if(nextrun<=0){ pauseClock( formId , time , server_date , vdomain ); continue; }
				var endDiff = endTimeStamp-time;
				//alert( endTimeStamp+"\n"+time );
				if(nextrun<=0 && endDiff>86400)
				{ 
					nextrun +=86400;
				}
			}
		}
		//alert(nextrun+"\n"+typeof(nextrun));
		runClock( formId , nextrun , time , vdomain  );
	}//end for
}//end function

//检查用户是否打开了闹钟，如果打开了闹钟则执行
function clockStart()
{
	//alert('闹钟开始执行');
	//alert('checkStart');
	var co = GetCookieObj();
	for( var i=1;i<=3;i++ )
	{
		var cKey = "clock_kg_"+i;
		if( co[cKey]=='1' )
		{
			doLocalClock();
			break;
		}
	}
	//alert('checkEnd');
}

CLOCK_FILE_END = true;