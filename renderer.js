// This file is required by the index.html file and will
// be executed in the renderer process for that window.
// All of the Node.js APIs are available in this process.
//根据点表结构改了，L列是点号，B列是值
//下一步要添加的功能：
// 1.如果两个表相同，那就输出“无不同”
// 2.导出excel
// 3.界面再搞好看一点



window.onload=function() {

	
	var dianji=document.getElementById('dianji');
	var xianshi=document.getElementById('xianshi');
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
  
  
	dianji.onclick=function(){
		// alert('aaa')
		
		// alert(filePath1);
		// alert(filePath2);
		var obj1=xlsx.parse(filePath1);
		var obj2=xlsx.parse(filePath2);
		
		//第一个工作表的数据，二维数组
		var data1=obj1[0].data;
		//  alert(data1[2].length);
		//alert(data1[0][1]);//表内0.1的内容
		//创建存放信息的数组
		var data2=obj2[0].data;
			
		var myArray1=[];
		var myArray2=[];

		var bijiao =[];//存放比较结果
		
		// //将表1内容读到数组中
		// for(var i in data1){
		// 	var box=new Box();//一条信息的结构
		// 	box.dianhao=data1[i][0];
		// 	box.zhi=data1[i][1];
		// 	//alert(data1[i]);//表中一行信息的内容
		// 	myArray1[i]=box;
		// 	box=null;
		// }

		// alert(myArray1[3].zhi)
		
		// //将表2内容读到数组中
		// for(var i in data2){
		// 	var box=new Box();//一条信息的结构
		// 	box.dianhao=data2[i][0];
		// 	box.zhi=data2[i][1];
		// 	//alert(data1[i]);//表中一行信息的内容
		// 	myArray2[i]=box;
		// 	box=null;
		// }
		// alert(myArray2[3].zhi)
		myArray1=shuzu(data1);
		myArray2=shuzu(data2);
		//alert(myArray1[1].zhi)
		
		for(var i =0;i<myArray1.length;){
			// alert('循环执行:i='+i)
			var myData1=myArray1[i];
			var flag=0;//
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
		
		

		//将比较结果比较写入textarea,点击按钮可显示出来
		//进一步添加的功能1.   不好使！！
		// if(bijiao.length!=0){
		// 	var thead=bijiao.createTHead();
		// 	var tr=thead.insertRow(0);
		// 	tr.insertCell(0).innerHTML='&nbsp&nbsp点号&nbsp&nbsp';
		// 	tr.insertCell(0).innerHTML='&nbsp&nbsp表1值&nbsp&nbsp';
		// 	tr.insertCell(0).innerHTML='&nbsp&nbsp表2值&nbsp&nbsp';			
		
		// 	for(var i;i<bijiao.length;i++){
		// 		var tr=xianshi.tBodies[0].insertRow(i);
		// 		tr.insertCell(0).innerHTML=bijiao[i].dianhao;
		// 		tr.insertCell(1).innerHTML=bijiao[i].zhi1;
		// 		tr.insertCell(2).innerHTML=bijiao[i].zhi2;
				
		// 	}
		// }
		for(var i;i<bijiao.length;i++){
			var tr=xianshi.tBodies[0].insertRow(i);
			tr.insertCell(0).innerHTML=bijiao[i].dianhao;
			tr.insertCell(1).innerHTML=bijiao[i].zhi1;
			tr.insertCell(2).innerHTML=bijiao[i].zhi2;
			
		}
		//writeXls(bijiao);
		xianshi.caption.innerHTML='比较结果:共有'+i+'条记录'+'<br><br>结果列表';
		xianshi.style.display='inline-block';//表格显示居中
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
function shuzu(data2){
	var myArray2=[];
	for(var i in data2){
		var box=new Box();//一条信息的结构
		box.dianhao=data2[i][11];
		box.zhi=data2[i][1];
		//alert(data1[i]);//表中一行信息的内容
		myArray2[i]=box;
		box=null;
	}
	return myArray2;
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