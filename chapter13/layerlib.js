/* layerlib.js: Simple Layer library with basic compatibility checking */
/* ������ */
var layerobject = ((document.layers) ? (true) : (false));
var dom = ((document.getElementById) ? (true) : (false));
var allobject = ((document.all) ? (true) : (false));
/* �������� */
opera=navigator.userAgent.toLowerCase().indexOf('opera')!=-1;
/* Ϊ��������layerNameֵ���ݶ��� */
function getElement(layerName,parentLayer)   //��������ȡlayerName���������
{
 if(layerobject)
  {
    parentLayer = (parentLayer)? parentLayer : self;
    layerCollection = parentLayer.document.layers;
    if (layerCollection[layerName])
      return layerCollection[layerName];
    /*��Ƕ�ײ�������� */
    for(i=0; i << layerCollection.length;)
      return(getElement(layerName, layerCollection[i++]));
  }
  if (allobject)
    return document.all[layerName];                  //����all[]����
  if (dom)
    return document.getElementById(layerName);    //����DOM����
}
/* ���� id = layerName �Ĳ�*/
function hide(layerName)
{
   var theLayer = getElement(layerName);  //��ȡlayerName���������
   if (layerobject)
     theLayer.visibility = 'hide';
   else
     theLayer.style.visibility = 'hidden';       //������������
}
/* ��ʾ id = layerName�Ĳ� */
function show(layerName)
{
   var theLayer = getElement(layerName);  //��ȡlayerName���������
   if (layerobject)
     theLayer.visibility = 'show';
   else
     theLayer.style.visibility = 'visible';          //������������
}
/* ������ΪlayerName �Ĳ��x���� */
function setX(layerName, x) 
{
   var theLayer = getElement(layerName);  //��ȡlayerName���������
   if (layerobject)
     theLayer.left=x;
   else if (opera)                        //������opera�е����
     theLayer.style.pixelLeft=x;
   else  
     theLayer.style.left=x+"px";              //������������
} 
/*������ΪlayerName �Ĳ��y����*/
function setY(layerName, y) 
{
   var theLayer = getElement(layerName);   //��ȡlayerName���������
   if (layerobject)
     theLayer.top=y;
   else if (opera)                     //������opera�е����
     theLayer.style.pixelTop=y;
   else  
     theLayer.style.top=y+"px";        //������������
}
/*������ΪlayerName �Ĳ��z-Index */
function setZ(layerName, zIndex) 
{
   var theLayer = getElement(layerName);  //��ȡlayerName���������
   if (layerobject)
     theLayer.zIndex = zIndex;
   else
     theLayer.style.zIndex = zIndex;    //������������
} 
/*������ΪlayerName �Ĳ�ĸ߶�*/
function setHeight(layerName, height) 
{
   var theLayer = getElement(layerName);  //��ȡlayerName���������
   if (layerobject)
     theLayer.clip.height = height;
   else if (opera)                       //������opera�е����
     theLayer.style.pixelHeight = height;
   else
     theLayer.style.height = height+"px";  //������������
}
/*������ΪlayerName �Ĳ�Ŀ��*/
function setWidth(layerName, width) 
{
  var theLayer = getElement(layerName);
  if (layerobject)
     theLayer.clip.width = width;
  else if (opera)                            //������opera�е����
     theLayer.style.pixelWidth = width;
  else
     theLayer.style.width = width+"px";        //������������
}
/* ������ top��right��bottom��left �������ΪlayerName �Ĳ�ľ��β���*/
function setClip(layerName, top, right, bottom, left) 
{
   var theLayer = getElement(layerName);       //��ȡlayerName���������
   if (layerobject)
     {                              //�ֱ�����top��right��bottom��left��ֵ
        theLayer.clip.top = top;
        theLayer.clip.right = right;
        theLayer.clip.bottom = bottom;
        theLayer.clip.left = left;
     }
   else
     //������������
     theLayer.style.clip = "rect("+top+"px "+right+"px "+" "+bottom+"px "+left+"px )";
}
/* ͨ��content��������layerName ������*/
function setContents(layerName, content)
{
   var theLayer = getElement(layerName);       //��ȡlayerName���������
   if (layerobject)
     {
       theLayer.document.write(content);
       theLayer.document.close();
       return;
     }
   if (theLayer.innerHTML)                  //������������
      theLayer.innerHTML = content;
}
