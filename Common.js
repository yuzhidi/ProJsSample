/*

       ���֣�Common.js

       ���ܣ�ͨ��JavaScript�ű�������

       ������

                     1.Trim(str)����ȥ���ַ������ߵĿո�

                     2.XMLEncode(str)�������ַ�������XML����

            3.ShowLabel(str,str)���������ʾ���ܣ���ʾ�ַ�����ʾ�ַ���

                     4.IsEmpty(obj)������֤������Ƿ�Ϊ��

                     5.IsInt(objStr,sign,zero)������֤�Ƿ�Ϊ����

                     6.IsFloat(objStr,sign,zero)������֤�Ƿ�Ϊ������

                     7.IsEnLetter(objStr,size)������֤�Ƿ�Ϊ26����ĸ

 

    */

 

/*

==================================================================

�ַ�������

Trim(string):ȥ���ַ������ߵĿո�

==================================================================

*/

 

/*

==================================================================

LTrim(string):ȥ����ߵĿո�

==================================================================

*/

function LTrim(str)

{

    var whitespace = new String(" \t\n\r");

    var s = new String(str);

    

    if (whitespace.indexOf(s.charAt(0)) != -1)

    {

        var j=0, i = s.length;

        while (j < i && whitespace.indexOf(s.charAt(j)) != -1)

        {

            j++;

        }

        s = s.substring(j, i);

    }

    return s;

}

 

/*

==================================================================

RTrim(string):ȥ���ұߵĿո�

==================================================================

*/

function RTrim(str)

{

    var whitespace = new String(" \t\n\r");

    var s = new String(str);

 

    if (whitespace.indexOf(s.charAt(s.length-1)) != -1)

    {

        var i = s.length - 1;

        while (i >= 0 && whitespace.indexOf(s.charAt(i)) != -1)

        {

            i--;

        }

        s = s.substring(0, i+1);

    }

    return s;

}

 

/*

==================================================================

Trim(string):ȥ��ǰ��ո�

==================================================================

*/

function Trim(str)

{

    return RTrim(LTrim(str));

}

 

 

 

/*

================================================================================

XMLEncode(string):���ַ�������XML����

================================================================================

*/

function XMLEncode(str)

{

       str=Trim(str);

       str=str.replace("&","&amp;");

       str=str.replace("<","&lt;");

       str=str.replace(">","&gt;");

       str=str.replace("'","&apos;");

       str=str.replace("\"","&quot;");

       return str;

}

 

/*

================================================================================

��֤�ຯ��

================================================================================

*/

 

function IsEmpty(obj)

{

    obj=document.getElementsByName(obj).item(0);

    if(Trim(obj.value)=="")

    {

        alert("�ֶβ���Ϊ�ա�");        

        if(obj.disabled==false && obj.readOnly==false)

        {

            obj.focus();

        }

    }

}

 

/*

IsInt(string,string,int or string):(�����ַ���,+ or - or empty,empty or 0)

���ܣ��ж��Ƿ�Ϊ����������������������������+0��������+0

*/

function IsInt(objStr,sign,zero)

{

    var reg;    

    var bolzero;    

    

    if(Trim(objStr)=="")

    {

        return false;

    }

    else

    {

        objStr=objStr.toString();

    }    

    

    if((sign==null)||(Trim(sign)==""))

    {

        sign="+-";

    }

    

    if((zero==null)||(Trim(zero)==""))

    {

        bolzero=false;

    }

    else

    {

        zero=zero.toString();

        if(zero=="0")

        {

            bolzero=true;

        }

        else

        {

            alert("����Ƿ����0������ֻ��Ϊ(�ա�0)");

        }

    }

    

    switch(sign)

    {

        case "+-":

            //����

            reg=/(^-?|^\+?)\d+$/;            

            break;

        case "+": 

            if(!bolzero)           

            {

                //������

                reg=/^\+?[0-9]*[1-9][0-9]*$/;

            }

            else

            {

                //������+0

                //reg=/^\+?\d+$/;

                reg=/^\+?[0-9]*[0-9][0-9]*$/;

            }

            break;

        case "-":

            if(!bolzero)

            {

                //������

                reg=/^-[0-9]*[1-9][0-9]*$/;

            }

            else

            {

                //������+0

                //reg=/^-\d+$/;

                reg=/^-[0-9]*[0-9][0-9]*$/;

            }            

            break;

        default:

            alert("�����Ų�����ֻ��Ϊ(�ա�+��-)");

            return false;

            break;

    }

    

    var r=objStr.match(reg);

    if(r==null)

    {

        return false;

    }

    else

    {        

        return true;     

    }

}

 

/*

IsFloat(string,string,int or string):(�����ַ���,+ or - or empty,empty or 0)

���ܣ��ж��Ƿ�Ϊ������������������������������������+0����������+0

*/

function IsFloat(objStr,sign,zero)

{

    var reg;    

    var bolzero;    

    

    if(Trim(objStr)=="")

    {

        return false;

    }

    else

    {

        objStr=objStr.toString();

    }    

    

    if((sign==null)||(Trim(sign)==""))

    {

        sign="+-";

    }

    

    if((zero==null)||(Trim(zero)==""))

    {

        bolzero=false;

    }

    else

    {

        zero=zero.toString();

        if(zero=="0")

        {

            bolzero=true;

        }

        else

        {

            alert("����Ƿ����0������ֻ��Ϊ(�ա�0)");

        }

    }

    

    switch(sign)

    {

        case "+-":

            //������

            reg=/^((-?|\+?)\d+)(\.\d+)?$/;

            break;

        case "+": 

            if(!bolzero)           

            {

                //��������

                reg=/^\+?(([0-9]+\.[0-9]*[1-9][0-9]*)|([0-9]*[1-9][0-9]*\.[0-9]+)|([0-9]*[1-9][0-9]*))$/;

            }

            else

            {

                //��������+0

                reg=/^\+?\d+(\.\d+)?$/;

            }

            break;

        case "-":

            if(!bolzero)

            {

                //��������

                reg=/^-(([0-9]+\.[0-9]*[1-9][0-9]*)|([0-9]*[1-9][0-9]*\.[0-9]+)|([0-9]*[1-9][0-9]*))$/;

            }

            else

            {

                //��������+0

                reg=/^((-\d+(\.\d+)?)|(0+(\.0+)?))$/;

            }            

            break;

        default:

            alert("�����Ų�����ֻ��Ϊ(�ա�+��-)");

            return false;

            break;

    }

    

    var r=objStr.match(reg);

    if(r==null)

    {

        return false;

    }

    else

    {        

        return true;     

    }

}

