				var week; 
				if(new Date().getDay()==0)          week="������"
				if(new Date().getDay()==1)          week="����һ"
				if(new Date().getDay()==2)          week="���ڶ�" 
				if(new Date().getDay()==3)          week="������"
				if(new Date().getDay()==4)          week="������"
				if(new Date().getDay()==5)          week="������"
				if(new Date().getDay()==6)          week="������"
				document.write(new Date().getFullYear()+"��"+(new Date().getMonth()+1)+"��"+new Date().getDate()+"�� "+week);