/* layerlib.js: Simple Layer library with basic compatibility checking */
/* 检测对象 */
var layerobject = ((document.layers) ? (true) : (false));
var dom = ((document.getElementById) ? (true) : (false));
var allobject = ((document.all) ? (true) : (false));
/* 检测浏览器 */
opera=navigator.userAgent.toLowerCase().indexOf('opera')!=-1;
/* 为传递来的layerName值传递对象 */
function getElement(layerName,parentLayer)   //函数：获取layerName的相关属性
{
 if(layerobject)
  {
    parentLayer = (parentLayer)? parentLayer : self;
    layerCollection = parentLayer.document.layers;
    if (layerCollection[layerName])
      return layerCollection[layerName];
    /*在嵌套层次中搜索 */
    for(i=0; i << layerCollection.length;)
      return(getElement(layerName, layerCollection[i++]));
  }
  if (allobject)
    return document.all[layerName];                  //设置all[]环境
  if (dom)
    return document.getElementById(layerName);    //设置DOM环境
}
/* 隐藏 id = layerName 的层*/
function hide(layerName)
{
   var theLayer = getElement(layerName);  //获取layerName的相关属性
   if (layerobject)
     theLayer.visibility = 'hide';
   else
     theLayer.style.visibility = 'hidden';       //设置其他环境
}
/* 显示 id = layerName的层 */
function show(layerName)
{
   var theLayer = getElement(layerName);  //获取layerName的相关属性
   if (layerobject)
     theLayer.visibility = 'show';
   else
     theLayer.style.visibility = 'visible';          //设置其他环境
}
/* 设置名为layerName 的层的x坐标 */
function setX(layerName, x) 
{
   var theLayer = getElement(layerName);  //获取layerName的相关属性
   if (layerobject)
     theLayer.left=x;
   else if (opera)                        //设置在opera中的情况
     theLayer.style.pixelLeft=x;
   else  
     theLayer.style.left=x+"px";              //设置其他环境
} 
/*设置名为layerName 的层的y坐标*/
function setY(layerName, y) 
{
   var theLayer = getElement(layerName);   //获取layerName的相关属性
   if (layerobject)
     theLayer.top=y;
   else if (opera)                     //设置在opera中的情况
     theLayer.style.pixelTop=y;
   else  
     theLayer.style.top=y+"px";        //设置其他环境
}
/*设置名为layerName 的层的z-Index */
function setZ(layerName, zIndex) 
{
   var theLayer = getElement(layerName);  //获取layerName的相关属性
   if (layerobject)
     theLayer.zIndex = zIndex;
   else
     theLayer.style.zIndex = zIndex;    //设置其他环境
} 
/*设置名为layerName 的层的高度*/
function setHeight(layerName, height) 
{
   var theLayer = getElement(layerName);  //获取layerName的相关属性
   if (layerobject)
     theLayer.clip.height = height;
   else if (opera)                       //设置在opera中的情况
     theLayer.style.pixelHeight = height;
   else
     theLayer.style.height = height+"px";  //设置其他环境
}
/*设置名为layerName 的层的宽度*/
function setWidth(layerName, width) 
{
  var theLayer = getElement(layerName);
  if (layerobject)
     theLayer.clip.width = width;
  else if (opera)                            //设置在opera中的情况
     theLayer.style.pixelWidth = width;
  else
     theLayer.style.width = width+"px";        //设置其他环境
}
/* 设置由 top、right、bottom和left 定义的名为layerName 的层的矩形裁切*/
function setClip(layerName, top, right, bottom, left) 
{
   var theLayer = getElement(layerName);       //获取layerName的相关属性
   if (layerobject)
     {                              //分别设置top、right、bottom和left的值
        theLayer.clip.top = top;
        theLayer.clip.right = right;
        theLayer.clip.bottom = bottom;
        theLayer.clip.left = left;
     }
   else
     //设置其他环境
     theLayer.style.clip = "rect("+top+"px "+right+"px "+" "+bottom+"px "+left+"px )";
}
/* 通过content参数设置layerName 的内容*/
function setContents(layerName, content)
{
   var theLayer = getElement(layerName);       //获取layerName的相关属性
   if (layerobject)
     {
       theLayer.document.write(content);
       theLayer.document.close();
       return;
     }
   if (theLayer.innerHTML)                  //设置其他环境
      theLayer.innerHTML = content;
}
