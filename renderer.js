// This file is required by the index.html file and will
// be executed in the renderer process for that window.
// All of the Node.js APIs are available in this process.
//根据点表结构改了，L列是点号，B列是值
//下一步要添加的功能：
// 1.如果两个表相同，那就输出“无不同”--实现了
// 2.导出excel
// 3.界面再搞好看一点
// 4.点几次就只显示一次表
// 5.选择要比较哪两列  --实现了



window.onload=function() {

	
	var dianji=document.getElementById('dianji');
	var xianshi=document.getElementById('xianshi');
	var xianshi0=document.getElementById('xianshi0');
	var xlsx=require('node-xlsx');
	//获取路径
	var filePath1=null;
	var filePath2=null;
	var holder1 = document.getElementById('holder1');

	holder1.ondragover = function () {
		return false;
	};
	holder1.ondragleave = holder1.ondragend = function () {
		return false;
	};
	holder1.ondrop = function (e) {
		e.preventDefault();
		var file = e.dataTransfer.files[0];
		filePath1=file.path;
		holder1.innerHTML=filePath1;
		// alert(filePath1);
		return false;
	};
	
	var holder2 = document.getElementById('holder2'); 

	holder2.ondragover = function () {
		return false;
	};
	holder2.ondragleave = holder2.ondragend = function () {
		return false;
	};
	holder2.ondrop = function (e) {
		e.preventDefault();
		var file = e.dataTransfer.files[0];
		filePath2=file.path;
		holder2.innerHTML=filePath2;
		return false;
	};

	//此处为调试增加两个path，不调了删了就行
	filePath1='F:\\功能说明.xlsx';holder1.innerHTML=filePath1;
	filePath2='F:\\功能说明 - 副本.xlsx';holder2.innerHTML=filePath2;
	// alert('aaa')
	
	// alert(filePath1);
	// alert(filePath2);
	var obj1=xlsx.parse(filePath1);
	var obj2=xlsx.parse(filePath2);
	
	//第一个工作表的数据，二维数组
	var data1=obj1[0].data;
	
	//alert(data1[0][1]);//表内0.1的内容
	//创建存放信息的数组
	var data2=obj2[0].data;
		
	var myArray1=[];
	var myArray2=[];

	var bijiao =[];//存放比较结果
			
	var myBiaotou=biaotou(data1);
	
	var select1=document.getElementById('select1');
	var select2=document.getElementById('select2');

	//将表头写入select列表框
	for(var i=0;i<myBiaotou.length;i++){
		var op1=document.createElement('option');
		var op2=document.createElement('option');
		select1.appendChild(op1);
		select2.appendChild(op2);
		op1.innerHTML=op2.innerHTML=myBiaotou[i];
	}
	//获取需要比较的是第几列
	var index11=null;
	var index22=null;
	// function b(index0){index11=index0;}
	var index1=index2=null;
	select1.onchange=function(){
		index1=select1.selectedIndex;
		for(var i=1;i<data1[0].length;i++){
			if(select1.options[index1].innerHTML==data1[0][i]){
				index11=i;
				break;
			}

		}
	
		
	}

	select2.onchange=function(){
		index2=select2.selectedIndex;
		for(var i=1;i<data1[0].length;i++){
			if(select2.options[index2].innerHTML==data1[0][i]){
				index22=i;
				break;
			}
		}
		
	}
	
	dianji.onclick=function(){
		bijiao =[];
		if(index11==null){alert('请选择比较内容');}
		else{
			myArray1=shuzu(data1,index11,index22);
			myArray2=shuzu(data2,index11,index22);
			
			for(var i =0;i<myArray1.length;){
				// alert('循环执行:i='+i)
				var myData1=myArray1[i];
				var flag=0;//找到点号=1，没找着=0
				// alert(myData1.zhi);
				//遍历myArray2,找点号，比较值是否相同
				for(var j=0;j<myArray2.length;j++){
					
					//点号相同
					if(myData1.dianhao==myArray2[j].dianhao){
						//值也相同
						if(myData1.zhi==myArray2[j].zhi){
							myArray2.splice(j,1);
							myArray1.splice(i,1);
							flag=1;
							// alert('值也相同')
							break;
						}
						//值不相同
						else{
							var jieguo=new Jieguo();
							jieguo.dianhao=myData1.dianhao;
							jieguo.zhi1=myData1.zhi;
							jieguo.zhi2=myArray2[j].zhi;
							bijiao.push(jieguo);
							myArray2.splice(j,1);
							myArray1.splice(i,1);
							flag=1;
							// alert('值不相同，比较长度'+bijiao.length)
							break;
						}
					}
				// alert(j);	
				}
				//未找到点号
				if(flag==0){
					var jieguo=new Jieguo();
					jieguo.dianhao=myData1.dianhao;
					jieguo.zhi1=myData1.zhi;
					jieguo.zhi2='无此点号';
					bijiao.push(jieguo);
					
					// alert('未找到点号,bijiao长度'+bijiao.length);
					myArray1.splice(i,1);
					continue;
				}
				
			}
			
			for(var j in myArray2){
				var jieguo=new Jieguo();
				jieguo.dianhao=myArray2[j].dianhao;
				jieguo.zhi1='无此点号';
				jieguo.zhi2=myArray2[j].zhi;
				bijiao.push(jieguo);
			}
			
			// 实现了下一步要添加的功能： 1.如果两个表相同，那就输出“无不同”
			var tbody=document.getElementsByTagName('tbody')[0];
			// 实现了下一步要添加的功能： 4.点几次就只显示一次表
			xianshi.removeChild(tbody);
			tbody=null;
			var newTbody=document.createElement('tbody');
			xianshi.appendChild(newTbody);
			
			// 写进html表格
			if(bijiao.length==0){
				
				xianshi0.innerHTML='比较结果:共有0条记录';
			
			}else{
				
				for(var i=0;i<bijiao.length;i++){
				var tr=xianshi.tBodies[0].insertRow(i);
				tr.insertCell(0).innerHTML=bijiao[i].dianhao;
				tr.insertCell(1).innerHTML=bijiao[i].zhi1;
				tr.insertCell(2).innerHTML=bijiao[i].zhi2;
				}
				
				//writeXls(bijiao);
				xianshi.caption.innerHTML='比较结果:共有'+i+'条记录'+'<br><br>结果列表';
				xianshi.style.display='inline-block';//表格显示居中
			}
			
		}
			
	};

}

//构造函数创建对象：点号，值
function Box(dianhao,zhi){
	this.dianhao=dianhao;
	this.zhi=zhi;
}
//构造函数创建结果对象：点号，值1，值2
function Jieguo(dianhao,zhi1,zhi2){
	this.dianhao=dianhao;
	this.zhi1=zhi1;
	this.zhi2=zhi2;
}

//读取表的内容，将其放在数组里
//v1.0.2改动：根据点表结构，excel中L列是点号，B列是值
function shuzu(data2,index11,index22){
	var myArray2=[];
	for(var i in data2){
		var box=new Box();//一条信息的结构
		box.dianhao=data2[i][index11];
		box.zhi=data2[i][index22];
		//alert(data1[i]);//表中一行信息的内容
		myArray2[i]=box;
		box=null;
	}
	return myArray2;
}
//读取表头
function biaotou(data2){
	var myBiaotou=[];
	for(var i in data2){
		if(data2[0][i]!=null){
			myBiaotou.push(data2[0][i]);
		}
	}
	
	return myBiaotou;
}

//导出bijiao中的内容，形成excel.!!这个功能不好使
function writeXls(datas){
	var buffer=xlsx.build([
		{
			name:'sheet1',
			data:datas
		}
	]);
	fs.writeFileSync('test1.xlsx',buffer,{'flag':'w'});
}