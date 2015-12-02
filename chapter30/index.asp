
<style type="text/css">
.spanstyle {
COLOR: #0000ff; FONT-FAMILY: 华文新魏; FONT-SIZE: 12pt; POSITION: absolute; TOP: -50px; VISIBILITY: visible
}
</style>

<script language="javaScript">
<!--
var x,y
var step=18
var flag=0
var message=""

message=message.split("")
var xpos=new Array()
for (i=0;i<=message.length-1;i++) {
xpos[i]=-50
}

var ypos=new Array()
for (i=0;i<=message.length-1;i++) {
ypos[i]=-200
}

function handlerMM(e){
x = (document.layers) ? e.pageX : document.body.scrollLeft+event.clientX
y = (document.layers) ? e.pageY : document.body.scrollTop+event.clientY
flag=1
}

function www_helpor_net() {
if (flag==1 && document.all) {
for (i=message.length-1; i>=1; i--) {
xpos[i]=xpos[i-1]+step
ypos[i]=ypos[i-1]
}
xpos[0]=x+step
ypos[0]=y

for (i=0; i<message.length-1; i++) {
var thisspan = eval("span"+(i)+".style")
thisspan.posLeft=xpos[i]
thisspan.posTop=ypos[i]
}
}

else if (flag==1 && document.layers) {
for (i=message.length-1; i>=1; i--) {
xpos[i]=xpos[i-1]+step
ypos[i]=ypos[i-1]
}
xpos[0]=x+step
ypos[0]=y

for (i=0; i<message.length-1; i++) {
var thisspan = eval("document.span"+i)
thisspan.left=xpos[i]
thisspan.top=ypos[i]
}
}
var timer=setTimeout("www_helpor_net()",30)
}

for (i=0;i<=message.length-1;i++) {
document.write("<span id='span"+i+"' class='spanstyle'>")
document.write(message[i])
document.write("</span>")
}

if (document.layers){
document.captureEvents(Event.MOUSEMOVE);
}
document.onmousemove = handlerMM;
www_helpor_net();
//-->
</script>

<%dim time1,time2
	  time1=timer
dim page,indexfilename,indeximg,db,n,x,bookbg,txt,jd100_top,jd100_foot,m,jd100_fla
indexfilename=right(Request.ServerVariables("PATH_TRANSLATED"),(len(Request.ServerVariables("PATH_TRANSLATED"))-instrRev(Request.ServerVariables("PATH_TRANSLATED"),"\"))) 
imdeximg="images/" '图片文件夹
db="guestbook.mdb" '数据库
           set Conn=Server.CreateObject("ADODB.Connection")
           Conn.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" & Server.MapPath(db)

n=10 '每页显示留言数
x=5  '每页显示的页数
m=10 '留言头像可选个数,男101-199.gif 女001-099.gif,各可增加到99个
bookbg="bookbg.gif"  '背景图片,当不使用背景图时,保持为空 ""
txt=1000  '留言的最大字数
jd100_top="<IMG src="&imdeximg&"welcome.gif>"   '设置页头信息welcome.gif可换成你的logo放在图片目录下
           dim webtitle,webname,webyn,webgl,webyn2,view2
           set rs1 = conn.execute("select * from admin")
           webtitle=rs1("title")
           if rs1("webname")<>"" then webname=rs1("webname")
           if rs1("gbyn")<>"" then webyn=rs1("gbyn")
           webgl=rs1("gl")
           rs1.close
           set rs1=nothing

'设置页脚信息,这里可以加入你的地址

page =Request.QueryString("page")
if page="" or page=0 then page=1

action = Request.QueryString("action")
action_e = Request.Form("action_e")
if action_e <>"" then
  server_v1=Cstr(Request.ServerVariables("HTTP_REFERER")) 
  server_v2=Cstr(Request.ServerVariables("SERVER_NAME")) 
   if mid(server_v1,8,len(server_v2))<>server_v2 then 
    response.write "<br><br><center><table border=1 cellpadding=20 bordercolor=black bgcolor=#EEEEEE width=450>" 
    response.write "<tr><td style='font:9pt Verdana'>" 
    response.write "你提交的路径有误，禁止从站点外部提交数据请不要乱该参数！" 
    response.write "</td></tr></table></center>" 
    response.end 
   end if 
end if

%>
<html>
	
<head>
<title>留言板</title>
<meta name="keywords" content="留言">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="gbstyle.css" type="text/css">
<style type="text/css">
<!--
.unnamed1 {
	font-size: 12px;
	line-height: 18px;
}
form {margin-bottom:0;margin-top:0}
.style1 {
	color: #990000;
	font-weight: bold;
}
.unnamed2 {
	font-size: 14px;
	line-height: 24px;
}
-->
</style>
</head>

<script language="JavaScript">
<!--
function gbcount(message,total,used,remain)
{
	var max;
	max = total.value;
	if (message.value.length > max) {
	message.value = message.value.substring(0,max);
	used.value = max;
	remain.value = 0;
	alert("留言不能超过 <%=txt%> 个字!");
	}
	else {
	used.value = message.value.length;
	remain.value = max - used.value;
	}
}
//-->
</script>

  <% if len(bookbg)<3 then
  bookbg=""
  else
  bookbg="background="& imdeximg & bookbg
  end if %>

<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="100%" height="51" border="0" align="center" cellpadding="0" cellspacing="1">
  <tr>
    <td height="49" valign="top" <%=bookbg%>>
	<div align="center">
      <CENTER>
        <%=jd100_top%>      </CENTER> 
    </div></td>
  </tr>
</table>
<table width="100%" height="18" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF" >
  <tr>
    <td height="18" align="center" valign="top" <%=bookbg%>> 
      <%
'主程序 
Select Case action_e
	Case ""

	Case "Add_New"
		Call Add_New_Execute()
	Case "reply"
		Call Reply_Execute()
	Case "admin"
		Call Admin_Login_Execute()
	Case "EditPWD"
		Call EditPWD_Execute()
	Case "Edit"
		Call Edit_Execute()
		
    Case "Edit_web"
		Call Edit_web()
		
End Select
Call Main_Menu()
Select Case action
    Case "UbbHelp"
        Call UbbHelp()
	Case "Admin_Login"
		Call Admin_Login()
	Case "Exit"
		Call Exit_Admin()
		
		Call View_Words()
		
	Case ""
		
		Call View_Words()
		
	Case "Add_New"
		Call Add_New()
	Case "reply"
		Call Reply()
	Case "View_Words"
		
		Call View_Words()
		
	Case "Delete"
		Call Delete()
		Call View_Words()
	Case "EditPWD"
		Call EditPWD()
	Case "Edit"
		Call Edit()
		
	Case "Edit_web"
		Call Edit_web()
				
End Select
%>
    </td>
  </tr>
</table>
<table width="100%" height="77" border="0" align="center" cellpadding="0" cellspacing="1">
  <tr> 
    <td height="75" valign="top" bgcolor="#FFFFFF"><table align=center cellpadding=0 cellspacing=0>
      <tr>
        <td width="760" ></td>
      </tr>
      <tr>
        <td class="footline"></td>
      </tr>
      <tr align=center height=60 style="line-height:130%">
          <td> <span class="unnamed1">
            </span><br>
            <table width="594" border=0 align=center cellPadding=0 cellSpacing=0>
              <tr>
                <td width="571" height="47" align="center"><span class="unnamed1"> Copyright 
                  (c) 2006-2007  
				  编制：<a href=mailto:mxqqn@163.com>Mengite</a>
				  <% time2=timer
                  Response.Write "执行时间"&(time2-time1)*1000&"毫秒"
                  %>
                </span></td>
              </tr>
          </table></td>
      </tr>
    </table>	</td>
  </tr>
</table>
<%
'添加一条新留言
%>
<% Sub Add_New() %>
<table width="598" border="0" align="center" cellpadding="4" cellspacing="1" class="unnamed1">
  <form name="form" method="post" action="<%=indexfilename%>">
    <tr> 
      <td height="25" colspan="3" align="center"> <div align="center"><font size="3"><strong>留　言</strong></font><font color="#000000"> 
          </font></div>
	  <img src="<%=imdeximg%>line.gif" width="500" height="1">	  </td>
    </tr>
    <tr> 
      <td width="80"  > <div align="right">姓名：</div></td>
      <td width="249"> <input type="text" name="name" class="input1" size="20" maxLength=10>
        *10个字内</td>
            <script>
			function showimage(){document.images.showimages.src="<%=imdeximg%>"+document.form.sex.options[document.form.sex.selectedIndex].value+""+document.form.img.options[document.form.img.selectedIndex].value+".gif";}
			</script>
      <td width="241">选择头像:
        <select name="img" size="1" onChange="showimage()">
		<% if m>99 then m=99
		for i=1 to m 
		g=""
		g=i
		if len(i)<2 then g="0"&i
		%>
		
                <option value='<%=g%>'><%=g%></option>
        <% next %>
	    </select>
	  </td>
    </tr>
    <tr> 
      <td align="right"> 性别： </td>
      <td> 
	  <select name="sex" size="1" onChange="showimage()">
              <option value="1">男</option>
              <option value="0">女</option>
      </select>
	  </td>
      <td rowspan="5">
	  <img src="<%=imdeximg%>101.gif" name=showimages id="showimages">
	  </td>
    </tr>
    <tr>
      <td align="right">QQ：</td>
      <td><input name="qq" type="text" class="input1" id="qq" size="35" maxLength=25></td>
    </tr>
    <tr> 
      <td align="right">主页： </td>
      <td> <input name="web" type="text" class="input1" value="http://" size="35" maxLength=50> </td>
    </tr>
    <tr> 
      <td align="right">来自：</td>
      <td><input name="come" type="text" class="input1" id="come" size="35"></td>
    </tr>
    <tr> 
      <td align="right"> 电子邮箱： </td>
      <td> <input name="email" type="text" class="input1" value="@" size="35" maxLength=50>
      * </td>
    </tr>
    <tr>
      <td align="right" valign="top">
     </td>
      <td colspan="2">
	  <% call ubb_jd100() %>
	  </td>
    </tr>
    <tr> 
      <td align="right" valign="top"> 留言内容： </td>
      <td colspan="2"> <textarea name="words" cols="50" rows="8" class="input1" 
	  onkeydown=gbcount(this.form.words,this.form.total,this.form.used,this.form.remain); 
	  onkeyup=gbcount(this.form.words,this.form.total,this.form.used,this.form.remain);></textarea>
      *
      </td>
    </tr>
    <tr>
      <td align="right" valign="top">&nbsp;</td>
      <td colspan="2">最多字数：<INPUT disabled maxLength=4 name=total size=3 value=<%=txt%>>
				已用字数：<INPUT disabled maxLength=4 name=used size=3 value=0>
				剩余字数：<INPUT disabled maxLength=4 name=remain size=3 value=<%=txt%>></td>
    </tr>
	
    <tr align="center"> 
      <td colspan="3"> <input type="hidden" name="action_e" value="Add_New"> <input type="submit" name="Submit" value="提交" class="input1"> 
        <input type="reset" name="Submit2" value="重写" class="input1"> <br>
		<img src="<%=imdeximg%>line.gif" width="500" height="1">
	  </td>
    </tr>
  </form>
</table>
		<br>
<% End Sub %>
		
<% Sub Main_Menu() %>
<table width="700" border="0" align="center" class="unnamed1">
  <tr>
    <td width="287"> <a href="../guestbook/index.asp"><img src="<%=imdeximg%>back.gif" width="99" height="25" border="0"></a><a href="<%=indexfilename%>?action=Add_New"><img src="<%=imdeximg%>post.gif" width="99" height="25" border="0"></a></td>
    <td width="353"> <div align="right">
        <% If Session("Admin")="Login" Then %>
        <a href="<%=indexfilename%>?action=Exit">退出管理</a> 
        <% Else %>
        <a href="<%=indexfilename%>?action=Admin_Login">管理留言</a> 
        <% End If %>
        <% If Session("Admin")="Login" Then %>
		<a href="<%=indexfilename%>?action=Edit_web">基本设置</a> 
        <a href="<%=indexfilename%>?action=EditPWD">修改密码</a> 
        <% End If %>
      </div></td>
  </tr>
</table>
<% End Sub 
'''''''''''''''''''''''
'查看留言
Sub View_Words() 
         dim gbcount,y,j,k
         set rs = conn.execute("select COUNT(*) as gbcount From words")
		 gbcount=rs("gbcount")
		 rs.close
		 
		 if gbcount/n = int(gbcount/n) then '计算出分页数
		 y=int(gbcount/n)
		 else
		 y=int(gbcount/n)+1
		 end if

         page2= int(page/x)
		 if page/x>page2 then page2=page2+1
		 k=page2*x
		 if k>y then k=y

		 '打开留言字段'
		 if page=1 then
		 sql="select top "&n&" id,name,sex,head,web,email,title,words,date,reply,ip,come,view,qq From words Order By id Desc"
		 else
		 sql="select id,name,sex,head,web,email,title,words,date,reply,ip,come,view,qq From words Order By id Desc"
		 end if
		 if Page >100 then
            rs.Open sql,Conn,1
         else
            Set Rs=Conn.Execute(sql)
         end if
         if Page>1 then RS.Move n*page-n
        %>
<table width="700" border="0" cellspacing="1" cellpadding="4" align="center">
          <tr>
            <td width="667" height="20" align="right" class="unnamed1">有<%=gbcount %>条留言  <%=page %>/<%=y %>页 分页
                <a href="?page=1"><<</a>
                <% if page2>1 then %>
                <a href="<%=indexfilename%>?page=<%=page2*x-x%>"><</a>
                <% end if %>
                <% For m =page2*x-(x-1) To k %>
      [<a href="<%=indexfilename%>?page=<%=m%>"><%=m%></a>]
      <%
    Next
    %>
      <% if page2*x < y then %>
      <a href="<%=indexfilename%>?page=<%=m%>">></a>
      <% end if %>
	  <a href="?page=<%=y %>">>></a>
            </td>
          </tr>
     <% if len(webtitle)>2 then %>
          <tr>
            <td height="20" align="right" class="unnamed1">  
	        <marquee onMouseOut=start(); onMouseOver=stop(); scrollamount=3>
            </marquee></td>
          </tr>
		  <%  end if %> 
</table>
<% if rs.bof  and rs.eof then Response.Write "当前没有留言记录" %>
<%
dim lou,words,reply,email,qq,web,come
if Request.QueryString("page")<2 then
lou=gbcount
else
lou=gbcount-((Request.QueryString("page")-1)*n)
end if 
i=0
do while not rs.eof and i<n
i=i+1
reply=""
words=""  
  words=rs("words")
  reply=rs("reply")
  %>

 
    <TABLE width=700 border=0 align="center" 
cellPadding=0 cellSpacing=0 borderColor=#111111 style="BORDER-COLLAPSE: collapse">
      <TBODY>
        <TR>
          <TD width="2%"><IMG src="<%=imdeximg%>T_left.gif" border=0></TD>
          <TD width="96%" background=<%=imdeximg%>Tt_bg.gif></TD>
          <TD width="2%"><IMG src="<%=imdeximg%>T_right.gif" 
  border=0></TD>
        </TR>
      </TBODY>
</TABLE>



  <TABLE width=700 height=51 border=1 align=center cellPadding=3 cellSpacing=0 bordercolor="#85ACE0" style="border-collapse:collapse" >
    <TBODY>
      <TR >
        <TD colSpan=2 height=25><table width="686"  border="0" class="unnamed1">
          <tr>
            <td width="28%" height="21">留言者:<b><%=rs("name")%></b></td>
            <td width="60%"> <div align="right">发表于:<%=year(Rs("date"))%>年<%=month(Rs("date"))%>月<%=day(Rs("date"))%>日
			    <% if len(trim(rs("web")))>8  then %>
                <a href="<%=rs("web")%>" target="_blank"><img src="<%=imdeximg%>homepage.gif" alt="我的主页是:<%=rs("web")%>" width=16 height=16 border="0"></a> 
                <% end if %>
				<% if len(trim(rs("email")))>6  then %>
             	<a href="mailto:<%=rs("email")%>"><img src="<%=imdeximg%>email.gif" alt="我的EMAIL是:<%=rs("email")%>" width="16" height="16" border="0"></a> 
				<% end if %>
                <% if len(trim(rs("qq")))>3  then %>
				<img src="<%=imdeximg%>oicq.gif" alt="我的QQ号:<%=Rs("qq")%>" width="16" height="16" border="0"> 
                <% end if %>
				<% if len(trim(Rs("come")))>1  then %>
				<img src="<%=imdeximg%>come.gif" alt="我来自:<%=Rs("come")%>" width="16" height="16"> 
                <% end if %>
				
				<% If Session("Admin") = "Login" Then %>
                <img src="<%=imdeximg%>ip.gif" align=absMiddle> <%=Rs("ip")%> 
                <font color="#666666"><a href="<%=indexfilename%>?action=Edit&id=<%=Rs("id")%>"> 
                <img src="<%=imdeximg%>reply.gif" alt="编辑回复" width="16" height="16" border="0"></a> 
                <a href="<%=indexfilename%>?action=Delete&id=<%=Rs("id")%>" onClick="return confirm('确定要删除吗？\n\n该操作不可恢复！')"> 
                <img src="<%=imdeximg%>del.gif" alt="删除留言" width="15" height="15" border="0"></a></font>
                <% end if %>
              </div></td>
            <td width="12%"><div align="right">第 <font color="#ff0000"><%=lou %></font> 条留言</div></td>
          </tr>
        </table>          </TD>
      </TR>
      <TR>
        <TD width="100" height="21" align=middle valign="top"><table width="100" border="0" align="center" >
          <tr>
            <td width="94">
			    <%if rs("head")=""  then %>

			                  <%if rs("sex")=1 then %>
                              <img src="<%=imdeximg%>101.gif">
                              <% else %>
                              <img src="<%=imdeximg%>001.gif">
                              <% end if %>

                <% else %>
                             <img src="<%=imdeximg & rs("sex") & rs("head") %>.gif">
                <% end if %></td>
          </tr>
        </table></TD>
        <TD width="582" height="21" valign="top" class=unnamed2  style="word-break:break-all">
		<table width="582" border="0" style="TABLE-LAYOUT: fixed" class=unnamed2>
          <tr>
            <td width="576" style="word-break:break-all">

              <% if webyn=1 and rs("view")=1 then %>
              <%=Ubb(unHtml(words))%> 
                  <% if len(trim(reply))>1 then%>
                      <hr size="1"> <span class="style1">斑竹回复:</span><br><font color="#990000"> <%=Ubb(unHtml(reply))%> </font>
                  <%end if %>
              <%end if %>

              <% if webyn<>1 then %>
              <%=Ubb(unHtml(words))%> 
                  <% if len(trim(reply))>1 then%>
                     <hr size="1"> <span class="style1">斑竹回复:</span><br> <font color="#990000"><%=Ubb(unHtml(reply))%> </font>
                  <%end if %>
              <%end if %>

              <% if webyn=1 and rs("view")=0 then%>
              留言需要经过审批才能查看 
              <%end if %>
            </td>
          </tr>
        </table> 
        
      </TD>
      </TR>
    </TBODY>
</TABLE>
 
      <TABLE width=700 border=0 align="center" 
cellPadding=0 cellSpacing=0 borderColor=#111111 style="BORDER-COLLAPSE:collapse">
        <TBODY>
          <TR>
            <TD width="1%"><IMG src="<%=imdeximg%>T_bottomleft.gif" border=0></TD>
            <TD width="97%" background=<%=imdeximg%>T_bottombg.gif></TD>
            <TD width="2%"><IMG src="<%=imdeximg%>T_bottomright.gif" 
  border=0></TD>
          </TR>
        </TBODY>
</TABLE>

  
  <br>

<%
	    lou=lou-1	 
		rs.movenext
    	loop
		Rs.Close
		Set Rs = Nothing
		%>
<table width="700" border="0" cellspacing="1" cellpadding="4" align="center">
  <tr>
    <td height="20" align="right" class="unnamed1"> 
    
    有<%=gbcount %>条留言  <%=page %>/<%=y %>页 分页
                <a href="?page=1"><<</a>
                <% if page2>1 then %>
                <a href="<%=indexfilename%>?page=<%=page2*x-x%>"><</a>
                <% end if %>
                <% For m =page2*x-(x-1) To k %>
      [<a href="<%=indexfilename%>?page=<%=m%>"><%=m%></a>]
      <%
    Next
    %>
      <% if page2*x < y then %>
      <a href="<%=indexfilename%>?page=<%=m%>">></a>
      <% end if %>
	  <a href="?page=<%=y %>">>></a>

	</td>
  </tr>
</table>
		<% End Sub %>
		<%
'''''''''管理员登陆接口
		%>
		<% Sub Admin_Login() 
			dim num1
			dim rndnum
			Randomize
			Do While Len(rndnum)<4
			num1=CStr(Chr((57-48)*rnd+48))
			rndnum=rndnum&num1
			loop
			session("jd100_rn")=rndnum
			
			%>
		<br>
		
<table width="499" border="0" align="center" cellpadding="4" cellspacing="1" class="unnamed1">
  <form name="reply" method="post" action="<%=indexfilename%>">
    <tr> 
      <td colspan="2" align="center"> 管理登陆 </td>
    </tr>
    <tr> 
      <td width="122"> <div align="right">用户名： </div></td>
      <td width="358"> <input name="username" type="text" class="input1" maxlength="25"> </td>
    </tr>
    <tr> 
      <td width="122"> <div align="right">密 码： </div></td>
      <td width="358"> <input name="password" type="password" class="input1" maxlength="25"> 
        <input type="hidden" name="action_e" value="admin"> </td>
    </tr>
    <tr>
      <td><div align="right">输入验证码：</div></td>
      <td> 
        <table width="258" border="0">
          <tr>
            <td width="175"><input name="jd100rz" type="text" class="input1" maxlength="4"></td>
            <td width="60" bgcolor="#cccccc"><div align="center"><%=session("jd100_rn")%></div></td>
          </tr>
        </table> </td>
    </tr>
    <tr> 
      <td colspan="2" align="center"> 
        <input type="submit" name="Submit32" value="登陆" class="input1"> 
      </td>
    </tr>
    <tr> 
      <td height="49" colspan="2" align="center">&nbsp;</td>
    </tr>
  </form>
</table>
		<br>
<% End Sub%>
        <%
            '''''''''''
		%>
<%Sub UbbHelp()%>
<div align="left">
  <%End Sub%>
</div>
<%Sub EditPWD()%>
<table width="499" border="0" align="center" cellpadding="4" cellspacing="1" class="unnamed1">
  <form name="editpwd" method="post" action="<%=indexfilename%>">
    <tr> 
      <td colspan="2" align="center"> <b>修改密码</b></td>
    </tr>
    <tr> 
      <td align="right" valign="middle" width="105" height="38">旧用户名：</td>
      <td width="226" height="38" align="left"> 
        <input type="text" name="oldusername" class="input1">
      </td>
    </tr>
    <tr> 
      <td align="right" valign="middle" width="105" height="38">新用户名：</td>
      <td width="226" height="38" align="left"> 
        <input type="text" name="username" class="input1">
      </td>
    </tr>
    <tr> 
      <td align="right" valign="middle" width="105" height="38">确认新用户名：</td>
      <td width="226" height="38" align="left"> 
        <input type="text" name="username_c" class="input1">
      </td>
    </tr>
    <tr> 
      <td align="right" valign="middle" width="105" height="38"> 旧 密 码： </td>
      <td width="226" height="38" align="left"> 
        <input type="password" name="oldpwd" class="input1">
      </td>
    </tr>
    <tr>
      <td align="right" valign="middle" width="105">新 密 码： </td>
      <td width="226" align="left">
        <input type="password" name="newpwd" class="input1">
      </td>
    </tr>
    <tr> 
      <td align="right" valign="middle" width="105" height="38">确认新密码：</td>
      <td width="226" align="left" height="38"> 
        <input type="password" name="newpwd_c" class="input1">
        <input type="hidden" name="action_e" value="EditPWD">
      </td>
    </tr>
    <tr align="center"> 
      <td colspan="2"> 
        <input type="submit" name="EditPWD" value="修改密码" class="input1">
      </td>
    </tr>
  </form>
</table>
<%End Sub%>

<% Sub Edit() %>
<%

Sql="Select * From words Where id="&Request.QueryString("id")
set rs=conn.execute(sql)
view2=""
if rs("view")=1 then
view2="checked"
end if
%>
<table width="690" border="0" align="center" cellpadding="4" cellspacing="1" class="unnamed1">
  <form name="reply" method="post" action="<%=indexfilename%>">
    <tr> 
      <td colspan="2" align="center"> 编辑留言内容及回复</td>
    </tr>
    <tr> 
      <td align="right" valign="top">留言者：</td>
      <td>网名:<font color="#ff0000"><%=Rs("name")%></font>&nbsp;性别:
        <font color="#ff0000"><%if Rs("sex")=1 then Response.Write "男" else Response.Write "女" end if%></font> &nbsp;来自:<font color="#ff0000"><%=Rs("come")%></font>
		 <br> 时间:<font color="#ff0000"><%=Rs("date")%></font>&nbsp;ip:<font color="#ff0000"><%=Rs("ip")%></font></td>
    </tr>
    <tr> 
      <td align="right" valign="top">邮箱：</td>
      <td>&nbsp;<b><%=Rs("email")%></b></td>
    </tr>
    <tr> 
      <td width="139" align="right" valign="top"> 留言内容：<br> 
	  <br> <input name="replyedit" id="replyedit" type="checkbox" value="1">
        <font color="#FF0000">修改原文</font> </td>
      <td width="532"> <textarea name="reply" cols="60" rows="8" class="input1" id="reply"><%=Rs("words")%></textarea> 
      </td>
    </tr>
    <tr align="center">
      <td align="right">&nbsp;</td>
      <td align="left"><% call ubb_jd100() %></td>
    </tr>
    <tr align="center"> 
      <td align="right">回复：</td>
      <td align="left"> <textarea name="words" cols="60" rows="8" class="input1" id="words"><%=Rs("reply")%></textarea> 
        <br> 
		<% if webyn=1 then%>
		<br> 
		<input name="view" type="checkbox" id="view" value="1" <%=view2%>>
        通过审批
		<% end if %>
		
	</td>
    </tr>
    <tr align="center"> 
      <td colspan="2"> <input type="hidden" name="action_e" value="Edit"> <input type="hidden" name="id" value="<%=Request.QueryString("id")%>"> 
        <input type="submit" name="Submit" value="修改留言" id="Submit" class="input1">
        &nbsp; 
        <input type="reset" name="reset" value="取 消" id="Submit2" class="input1">
        　　<a href="<%=indexfilename%>?action=View_Words">返回</a></td>
    </tr>
  </form>
</table>

<% 
rs.close
set rs=nothing

End Sub %>
<br>
<% Sub Edit_web() %>
<% 

if Request.Form("submit")="修改" then
Set Rs = Server.CreateObject("ADODB.RecordSet")
Sql="Select * From admin"
Rs.Open Sql,Conn,2,3
   rs("title")=Request.Form("webtitle")
   rs("gl")=Request.Form("webggg")
   rs("gbyn")=cint(Request.Form("webyn"))
   rs("webname")=Request.Form("webname")
   rs.update
rs.close
set rs=nothing
response.redirect indexfilename &"?action=Edit_web"
response.end
end if
webyn2=""
if webyn=1 then
webyn2="checked"
end if
%>
<table width="600" border="0" align="center" cellpadding="4" cellspacing="1" class="unnamed1">
  <form name="reply" method="post" action="<%=indexfilename%>">
    <tr> 
      <td colspan="2" align="center"> 编辑留言板属性</td>
    </tr>
    <tr> 
      <td width="202" align="right" valign="top"> 留言板名称</td>
      <td width="379"><input name="webname" type="text" id="webname" value="<%=webname%>" size="50" maxLength=25> 
      </td>
    </tr>
    <tr align="center">
      <td align="right">公告内容：</td>
      <td align="left"><textarea name="webtitle" cols="50" rows="5" id="webtitle"><%=webtitle%></textarea></td>
    </tr>
    <tr align="center"> 
      <td align="right">词语过滤：</td>
      <td align="left"><input name="webggg" type="text" id="webggg" value="<%=webgl%>" size="50">
      </td>
    </tr>
    <tr align="center">
      <td align="right">经过审批才显示留言：</td>
      <td align="left"><input name="webyn" type="checkbox" id="webyn" value="1" <%=webyn2%>>
      是</td>
    </tr>
    <tr align="center"> 
      <td colspan="2"> 
        <input type="hidden" name="action_e" value="Edit_web">
        <input type="submit" name="Submit" value="修改" id="Submit" class="input1">
        　　<a href="<%=indexfilename%>?action=View_Words">返回</a></td>
    </tr>
  </form>
</table>
<% End Sub %>
</body>
</html>

<%  if jd100_fla=1 then
if Request("action")="View_Words" or Request("action")=""  then %>
<div id="Layer1" style="position:absolute; right:1px; top:250px;z-index:1" width="680" height="350">
  <object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,29,0" width="680" height="350">
    <param name="movie" value="<%=imdeximg%>fly.swf">
    <param name="quality" value="high">
    <param name="wmode" value="transparent">
    <embed src="<%=imdeximg%>fly.swf" quality="high" pluginspage="http://www.macromedia.com/go/getflashplayer" type="application/x-shockwave-flash" width="680" height="350"></embed></object>
</div>
<% End if 
   end if
%>

<% sub ubb_jd100()%>

<script language="JavaScript">
<!--
//UBB
function Cbold() {
fontbegin="[B]";
fontend="[/B]";
fontchuli();
}
function Citalic() {
fontbegin="[I]";
fontend="[/I]";
fontchuli();
}
function Cunder() {
fontbegin="[U]";
fontend="[/U]";
fontchuli();
}
function Ccenter() {
fontbegin="[center]";
fontend="[/center]";
fontchuli();
}

function showsize(size){
fontbegin="[size="+size+"]";
fontend="[/size]";
fontchuli();
}
function showcolor(color){
fontbegin="[color="+color+"]";
fontend="[/color]";
fontchuli();
}

function fontchuli(){
if ((document.selection)&&(document.selection.type == "Text")) {
var range = document.selection.createRange();
var ch_text=range.text;
range.text = fontbegin + ch_text + fontend;
} 
else {
document.form1.words.value=fontbegin+document.form1.words.value+fontend;
document.form1.words.focus();
}
}
//-->
</script>

	  <img onclick=Cbold() src="<%=imdeximg%>Ubb_bold.gif" width="23" height="22" alt="粗体" border="0"> 
	  <img onclick=Citalic() src="<%=imdeximg%>Ubb_italicize.gif" width="23" height="22" alt="斜体" border="0"> 
	  <img onclick=Cunder() src="<%=imdeximg%>Ubb_underline.gif" width="23" height="22" alt="下划线" border="0"> 
	  <img onclick=Ccenter() src="<%=imdeximg%>Ubb_center.gif" width="23" height="22" alt="居中" border="0">	
	 
                    字体大小
                   <select name="size" onChange="showsize(this.options[this.selectedIndex].value)">
                      <option value="1">1</option>
                      <option value="2">2</option>
                      <option value="3" selected>3</option>
                      <option value="4">4</option>
        </select>                    &nbsp;<font face="宋体" color=#333333>颜色：</font> 
                    <SELECT onchange=showcolor(this.options[this.selectedIndex].value) name=color>
                      <option style="background-color:#F0F8FF;color: #F0F8FF" value="#F0F8FF">#F0F8FF</option>
                      <option style="background-color:#FAEBD7;color: #FAEBD7" value="#FAEBD7">#FAEBD7</option>
                      <option style="background-color:#00FFFF;color: #00FFFF" value="#00FFFF">#00FFFF</option>
                      <option style="background-color:#7FFFD4;color: #7FFFD4" value="#7FFFD4">#7FFFD4</option>
                      <option style="background-color:#F0FFFF;color: #F0FFFF" value="#F0FFFF">#F0FFFF</option>
                      <option style="background-color:#F5F5DC;color: #F5F5DC" value="#F5F5DC">#F5F5DC</option>
                      <option style="background-color:#FFE4C4;color: #FFE4C4" value="#FFE4C4">#FFE4C4</option>
                      <option style="background-color:#000000;color: #000000" value="#000000">#000000</option>
                      <option style="background-color:#FFEBCD;color: #FFEBCD" value="#FFEBCD">#FFEBCD</option>
                      <option style="background-color:#0000FF;color: #0000FF" value="#0000FF">#0000FF</option>
                      <option style="background-color:#8A2BE2;color: #8A2BE2" value="#8A2BE2">#8A2BE2</option>
                      <option style="background-color:#A52A2A;color: #A52A2A" value="#A52A2A">#A52A2A</option>
                      <option style="background-color:#DEB887;color: #DEB887" value="#DEB887">#DEB887</option>
                      <option style="background-color:#5F9EA0;color: #5F9EA0" value="#5F9EA0">#5F9EA0</option>
                      <option style="background-color:#7FFF00;color: #7FFF00" value="#7FFF00">#7FFF00</option>
                      <option style="background-color:#D2691E;color: #D2691E" value="#D2691E">#D2691E</option>
                      <option style="background-color:#FF7F50;color: #FF7F50" value="#FF7F50">#FF7F50</option>
                      <option style="background-color:#6495ED;color: #6495ED" value="#6495ED" selected>#6495ED</option>
                      <option style="background-color:#FFF8DC;color: #FFF8DC" value="#FFF8DC">#FFF8DC</option>
                      <option style="background-color:#DC143C;color: #DC143C" value="#DC143C">#DC143C</option>
                      <option style="background-color:#00FFFF;color: #00FFFF" value="#00FFFF">#00FFFF</option>
                      <option style="background-color:#00008B;color: #00008B" value="#00008B">#00008B</option>
                      <option style="background-color:#008B8B;color: #008B8B" value="#008B8B">#008B8B</option>
                      <option style="background-color:#B8860B;color: #B8860B" value="#B8860B">#B8860B</option>
                      <option style="background-color:#A9A9A9;color: #A9A9A9" value="#A9A9A9">#A9A9A9</option>
                      <option style="background-color:#006400;color: #006400" value="#006400">#006400</option>
                      <option style="background-color:#BDB76B;color: #BDB76B" value="#BDB76B">#BDB76B</option>
                      <option style="background-color:#8B008B;color: #8B008B" value="#8B008B">#8B008B</option>
                      <option style="background-color:#556B2F;color: #556B2F" value="#556B2F">#556B2F</option>
                      <option style="background-color:#FF8C00;color: #FF8C00" value="#FF8C00">#FF8C00</option>
                      <option style="background-color:#9932CC;color: #9932CC" value="#9932CC">#9932CC</option>
                      <option style="background-color:#8B0000;color: #8B0000" value="#8B0000">#8B0000</option>
                      <option style="background-color:#E9967A;color: #E9967A" value="#E9967A">#E9967A</option>
                      <option style="background-color:#8FBC8F;color: #8FBC8F" value="#8FBC8F">#8FBC8F</option>
                      <option style="background-color:#483D8B;color: #483D8B" value="#483D8B">#483D8B</option>
                      <option style="background-color:#2F4F4F;color: #2F4F4F" value="#2F4F4F">#2F4F4F</option>
                      <option style="background-color:#00CED1;color: #00CED1" value="#00CED1">#00CED1</option>
                      <option style="background-color:#9400D3;color: #9400D3" value="#9400D3">#9400D3</option>
                      <option style="background-color:#FF1493;color: #FF1493" value="#FF1493">#FF1493</option>
                      <option style="background-color:#00BFFF;color: #00BFFF" value="#00BFFF">#00BFFF</option>
                      <option style="background-color:#696969;color: #696969" value="#696969">#696969</option>
                      <option style="background-color:#1E90FF;color: #1E90FF" value="#1E90FF">#1E90FF</option>
                      <option style="background-color:#B22222;color: #B22222" value="#B22222">#B22222</option>
                      <option style="background-color:#FFFAF0;color: #FFFAF0" value="#FFFAF0">#FFFAF0</option>
                      <option style="background-color:#228B22;color: #228B22" value="#228B22">#228B22</option>
                      <option style="background-color:#FF00FF;color: #FF00FF" value="#FF00FF">#FF00FF</option>
                      <option style="background-color:#DCDCDC;color: #DCDCDC" value="#DCDCDC">#DCDCDC</option>
                      <option style="background-color:#F8F8FF;color: #F8F8FF" value="#F8F8FF">#F8F8FF</option>
                      <option style="background-color:#FFD700;color: #FFD700" value="#FFD700">#FFD700</option>
                      <option style="background-color:#DAA520;color: #DAA520" value="#DAA520">#DAA520</option>
                      <option style="background-color:#808080;color: #808080" value="#808080">#808080</option>
                      <option style="background-color:#008000;color: #008000" value="#008000">#008000</option>
                      <option style="background-color:#ADFF2F;color: #ADFF2F" value="#ADFF2F">#ADFF2F</option>
                      <option style="background-color:#F0FFF0;color: #F0FFF0" value="#F0FFF0">#F0FFF0</option>
                      <option style="background-color:#FF69B4;color: #FF69B4" value="#FF69B4">#FF69B4</option>
                      <option style="background-color:#CD5C5C;color: #CD5C5C" value="#CD5C5C">#CD5C5C</option>
                      <option style="background-color:#4B0082;color: #4B0082" value="#4B0082">#4B0082</option>
                      <option style="background-color:#FFFFF0;color: #FFFFF0" value="#FFFFF0">#FFFFF0</option>
                      <option style="background-color:#F0E68C;color: #F0E68C" value="#F0E68C">#F0E68C</option>
                      <option style="background-color:#E6E6FA;color: #E6E6FA" value="#E6E6FA">#E6E6FA</option>
                      <option style="background-color:#FFF0F5;color: #FFF0F5" value="#FFF0F5">#FFF0F5</option>
                      <option style="background-color:#7CFC00;color: #7CFC00" value="#7CFC00">#7CFC00</option>
                      <option style="background-color:#FFFACD;color: #FFFACD" value="#FFFACD">#FFFACD</option>
                      <option style="background-color:#ADD8E6;color: #ADD8E6" value="#ADD8E6">#ADD8E6</option>
                      <option style="background-color:#F08080;color: #F08080" value="#F08080">#F08080</option>
                      <option style="background-color:#E0FFFF;color: #E0FFFF" value="#E0FFFF">#E0FFFF</option>
                      <option style="background-color:#FAFAD2;color: #FAFAD2" value="#FAFAD2">#FAFAD2</option>
                      <option style="background-color:#90EE90;color: #90EE90" value="#90EE90">#90EE90</option>
                      <option style="background-color:#D3D3D3;color: #D3D3D3" value="#D3D3D3">#D3D3D3</option>
                      <option style="background-color:#FFB6C1;color: #FFB6C1" value="#FFB6C1">#FFB6C1</option>
                      <option style="background-color:#FFA07A;color: #FFA07A" value="#FFA07A">#FFA07A</option>
                      <option style="background-color:#20B2AA;color: #20B2AA" value="#20B2AA">#20B2AA</option>
                      <option style="background-color:#87CEFA;color: #87CEFA" value="#87CEFA">#87CEFA</option>
                      <option style="background-color:#778899;color: #778899" value="#778899">#778899</option>
                      <option style="background-color:#B0C4DE;color: #B0C4DE" value="#B0C4DE">#B0C4DE</option>
                      <option style="background-color:#FFFFE0;color: #FFFFE0" value="#FFFFE0">#FFFFE0</option>
                      <option style="background-color:#00FF00;color: #00FF00" value="#00FF00">#00FF00</option>
                      <option style="background-color:#32CD32;color: #32CD32" value="#32CD32">#32CD32</option>
                      <option style="background-color:#FAF0E6;color: #FAF0E6" value="#FAF0E6">#FAF0E6</option>
                      <option style="background-color:#FF00FF;color: #FF00FF" value="#FF00FF">#FF00FF</option>
                      <option style="background-color:#800000;color: #800000" value="#800000">#800000</option>
                      <option style="background-color:#66CDAA;color: #66CDAA" value="#66CDAA">#66CDAA</option>
                      <option style="background-color:#0000CD;color: #0000CD" value="#0000CD">#0000CD</option>
                      <option style="background-color:#BA55D3;color: #BA55D3" value="#BA55D3">#BA55D3</option>
                      <option style="background-color:#9370DB;color: #9370DB" value="#9370DB">#9370DB</option>
                      <option style="background-color:#3CB371;color: #3CB371" value="#3CB371">#3CB371</option>
                      <option style="background-color:#7B68EE;color: #7B68EE" value="#7B68EE">#7B68EE</option>
                      <option style="background-color:#00FA9A;color: #00FA9A" value="#00FA9A">#00FA9A</option>
                      <option style="background-color:#48D1CC;color: #48D1CC" value="#48D1CC">#48D1CC</option>
                      <option style="background-color:#C71585;color: #C71585" value="#C71585">#C71585</option>
                      <option style="background-color:#191970;color: #191970" value="#191970">#191970</option>
                      <option style="background-color:#F5FFFA;color: #F5FFFA" value="#F5FFFA">#F5FFFA</option>
                      <option style="background-color:#FFE4E1;color: #FFE4E1" value="#FFE4E1">#FFE4E1</option>
                      <option style="background-color:#FFE4B5;color: #FFE4B5" value="#FFE4B5">#FFE4B5</option>
                      <option style="background-color:#FFDEAD;color: #FFDEAD" value="#FFDEAD">#FFDEAD</option>
                      <option style="background-color:#000080;color: #000080" value="#000080">#000080</option>
                      <option style="background-color:#FDF5E6;color: #FDF5E6" value="#FDF5E6">#FDF5E6</option>
                      <option style="background-color:#808000;color: #808000" value="#808000">#808000</option>
                      <option style="background-color:#6B8E23;color: #6B8E23" value="#6B8E23">#6B8E23</option>
                      <option style="background-color:#FFA500;color: #FFA500" value="#FFA500">#FFA500</option>
                      <option style="background-color:#FF4500;color: #FF4500" value="#FF4500">#FF4500</option>
                      <option style="background-color:#DA70D6;color: #DA70D6" value="#DA70D6">#DA70D6</option>
                      <option style="background-color:#EEE8AA;color: #EEE8AA" value="#EEE8AA">#EEE8AA</option>
                      <option style="background-color:#98FB98;color: #98FB98" value="#98FB98">#98FB98</option>
                      <option style="background-color:#AFEEEE;color: #AFEEEE" value="#AFEEEE">#AFEEEE</option>
                      <option style="background-color:#DB7093;color: #DB7093" value="#DB7093">#DB7093</option>
                      <option style="background-color:#FFEFD5;color: #FFEFD5" value="#FFEFD5">#FFEFD5</option>
                      <option style="background-color:#FFDAB9;color: #FFDAB9" value="#FFDAB9">#FFDAB9</option>
                      <option style="background-color:#CD853F;color: #CD853F" value="#CD853F">#CD853F</option>
                      <option style="background-color:#FFC0CB;color: #FFC0CB" value="#FFC0CB">#FFC0CB</option>
                      <option style="background-color:#DDA0DD;color: #DDA0DD" value="#DDA0DD">#DDA0DD</option>
                      <option style="background-color:#B0E0E6;color: #B0E0E6" value="#B0E0E6">#B0E0E6</option>
                      <option style="background-color:#800080;color: #800080" value="#800080">#800080</option>
                      <option style="background-color:#FF0000;color: #FF0000" value="#FF0000">#FF0000</option>
                      <option style="background-color:#BC8F8F;color: #BC8F8F" value="#BC8F8F">#BC8F8F</option>
                      <option style="background-color:#4169E1;color: #4169E1" value="#4169E1">#4169E1</option>
                      <option style="background-color:#8B4513;color: #8B4513" value="#8B4513">#8B4513</option>
                      <option style="background-color:#FA8072;color: #FA8072" value="#FA8072">#FA8072</option>
                      <option style="background-color:#F4A460;color: #F4A460" value="#F4A460">#F4A460</option>
                      <option style="background-color:#2E8B57;color: #2E8B57" value="#2E8B57">#2E8B57</option>
                      <option style="background-color:#FFF5EE;color: #FFF5EE" value="#FFF5EE">#FFF5EE</option>
                      <option style="background-color:#A0522D;color: #A0522D" value="#A0522D">#A0522D</option>
                      <option style="background-color:#C0C0C0;color: #C0C0C0" value="#C0C0C0">#C0C0C0</option>
                      <option style="background-color:#87CEEB;color: #87CEEB" value="#87CEEB">#87CEEB</option>
                      <option style="background-color:#6A5ACD;color: #6A5ACD" value="#6A5ACD">#6A5ACD</option>
                      <option style="background-color:#708090;color: #708090" value="#708090">#708090</option>
                      <option style="background-color:#FFFAFA;color: #FFFAFA" value="#FFFAFA">#FFFAFA</option>
                      <option style="background-color:#00FF7F;color: #00FF7F" value="#00FF7F">#00FF7F</option>
                      <option style="background-color:#4682B4;color: #4682B4" value="#4682B4">#4682B4</option>
                      <option style="background-color:#D2B48C;color: #D2B48C" value="#D2B48C">#D2B48C</option>
                      <option style="background-color:#008080;color: #008080" value="#008080">#008080</option>
                      <option style="background-color:#D8BFD8;color: #D8BFD8" value="#D8BFD8">#D8BFD8</option>
                      <option style="background-color:#FF6347;color: #FF6347" value="#FF6347">#FF6347</option>
                      <option style="background-color:#40E0D0;color: #40E0D0" value="#40E0D0">#40E0D0</option>
                      <option style="background-color:#EE82EE;color: #EE82EE" value="#EE82EE">#EE82EE</option>
                      <option style="background-color:#F5DEB3;color: #F5DEB3" value="#F5DEB3">#F5DEB3</option>
                      <option style="background-color:#FFFFFF;color: #FFFFFF" value="#FFFFFF">#FFFFFF</option>
                      <option style="background-color:#F5F5F5;color: #F5F5F5" value="#F5F5F5">#F5F5F5</option>
                      <option style="background-color:#FFFF00;color: #FFFF00" value="#FFFF00">#FFFF00</option>
                      <option style="background-color:#9ACD32;color: #9ACD32" value="#9ACD32">#9ACD32</option>
      </SELECT>
	  <% end sub %>

<%

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'添加新留言到数据库
Sub Add_New_Execute()

    '不良词语过滤
	
	
	If trim(Request.Form("name"))="" Then
	Response.Write("<script>alert(""姓名不能为空"");history.go(-1);</script>")
	Response.End
	End If
	If Len(Request.Form("name"))>20 Then
	Response.Write("<script>alert(""姓名太长"");history.go(-1);</script>")
	Response.End
	End If
	If Request.Form("email")<>"" Then
	If instr(Request.Form("email"),"@")=0 or instr(Request.Form("email"),"@")=1 or instr(Request.Form("email"),"@")=len(email) then
	Response.Write("<script>alert(""电子邮件不正确"");history.go(-1);</script>")
	Response.End
	End If

	End If
	
	If trim(Request.Form("words"))="" Then
	Response.Write("<script>alert(""留言不能为空"");history.go(-1);</script>")
	Response.End
	End If

	Set Rs = Server.CreateObject("ADODB.RecordSet")
	Sql="Select * From words"
	Rs.Open Sql,Conn,2,3
	Rs.AddNew
	Rs("name")=Server.HTMLEncode(Request.Form("name"))
	Rs("sex")=Server.HTMLEncode(Request.Form("sex"))
	Rs("head")=Server.HTMLEncode(Request.Form("head"))
	Rs("web")=Server.HTMLEncode(Request.Form("web"))
	Rs("email")=Server.HTMLEncode(Request.Form("email"))
	Rs("words")=Server.HTMLEncode(Request.Form("words"))
	Rs("qq")=Server.HTMLEncode(Request.Form("qq"))
	Rs("head")=Server.HTMLEncode(Request.Form("Img"))
	
	Rs("date")=Now()
    Rs("ip")=request.servervariables("remote_addr")
    Rs("come")=Server.HTMLEncode(Request.Form("come"))
	Rs.Update
	Rs.Close
	Set Rs = Nothing
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'验证管理员登陆

Sub Admin_Login_Execute()
	username = Server.HTMLEncode(Request.Form("username"))
	password = Server.HTMLEncode(Request.Form("password"))
	if trim(Server.HTMLEncode(Request.Form("jd100rz")))<>session("jd100_rn") then
	Response.Write("<script>alert(""验证码错误"");history.go(-1);</script>")
	Response.End
	end if
	
	session("jd100_rn")=""
	
	If username = "" OR password = "" Then
		Response.Write "用户名或者密码为空"
		Response.End
	End If
	Set Rs = Server.CreateObject("ADODB.RecordSet")
	Sql="Select * From admin"
	Rs.Open Sql,Conn,1,1
	If username = Rs("username") AND password = Rs("password") Then
		Session("Admin") = "Login"
		Else
		Response.Write "用户名或者密码不对，登陆失败"
	End If
	Rs.Close
	Set Rs = Nothing
End Sub
Sub EditPWD_Execute()
    If Session("Admin")="" Then 
		Response.Write "连接超时,请重新登录"
		Response.End
	end if
	
	oldusername=Server.HTMLEncode(Request.Form("oldusername"))
	username = Server.HTMLEncode(Request.Form("username"))
	username_c = Server.HTMLEncode(Request.Form("username_c"))
	oldpwd = Server.HTMLEncode(Request.Form("oldpwd"))
	newpwd = Server.HTMLEncode(Request.Form("newpwd"))
	newpwd_c = Server.HTMLEncode(Request.Form("newpwd_c"))
	If username = "" OR username_c="" Then
		Response.Write "新旧用户名均不能为空"
		Response.End
	End If
	If oldpwd = "" OR newpwd = "" OR newpwd_c="" Then
		Response.Write "新旧密码均不能为空"
		Response.End
	End If
	If username<>username_c Then
		Response.Write "新填写的两个新用户名不一致，请重新填写"
		Response.End
	End If
	If newpwd<>newpwd_c Then
		Response.Write "新填写的两个密码不一致，请重新填写"
		Response.End
	End If
	Set Rs = Server.CreateObject("ADODB.RecordSet")
	Sql="Select * From admin"
	Rs.Open Sql,Conn,2,3
	If Rs("password")=oldpwd And Rs("username")=oldusername Then
		Rs("username")=username
		Rs("password")=newpwd
		Rs.Update
	Else
		Response.Write "你的旧密码填写不对或者旧用户名不对，修改不成功"
		Response.End
	End If
	Rs.Close
	Set Rs = Nothing
End Sub
Sub Exit_Admin()
	Session.Abandon
	response.redirect indexfilename
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'删除数据
Sub Delete()
     If Session("Admin")="" Then 
		Response.Write "连接超时,请重新登录"
		Response.End
	 end if
	'删除数据
	Conn.Execute("Delete * From words Where id="&Request.QueryString("id"))
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'回复留言添加到数据库

Sub Reply_Execute()
    If Session("Admin")="" Then 
		Response.Write "连接超时,请重新登录"
		Response.End
	end if
	
	Set Rs = Server.CreateObject("ADODB.RecordSet")
	Sql="Select reply From words Where id="&Request.Form("id")
	Rs.Open Sql,Conn,2,3
	Rs("reply") = Server.HTMLEncode(Request.Form("reply"))
	Rs.Update
	Rs.Close
	Set Rs=Nothing
End Sub

Sub Edit_Execute()
    If Session("Admin")="" Then 
		Response.Write "连接超时,请重新登录"
		Response.End
	end if
	
	Set Rs = Server.CreateObject("ADODB.RecordSet")
	Sql="Select * From words Where id="&Request.Form("id")
	Rs.Open Sql,Conn,2,3

    if cint(Request.Form("replyedit"))=1 then
	Rs("words") = Server.HTMLEncode(Request.Form("reply"))
	end if

	Rs("reply") = Server.HTMLEncode(Request.Form("words"))
	
	if cint(Request.Form("view"))=1 then
	Rs("view")=1
	else
	Rs("view")=0
	end if
	Rs.Update
	Rs.Close
	Set Rs=Nothing
End Sub

Conn.Close
Set Conn = Nothing
%>
<%
function unHtml(content)
unHtml=content
if content <> "" then
'unHtml=replace(unHtml,"&","&amp;")
unHtml=replace(unHtml,"<","&lt;")
unHtml=replace(unHtml,">","&gt;")
unHtml=replace(unHtml,chr(34),"&quot;")
unHtml=replace(unHtml,chr(13),"<br>")
unHtml=replace(unHtml,chr(32),"&nbsp;") 
   unhtmlgl=split(webgl,"|")
   if IsArray(unhtmlgl) then
   for i=0 to UBound(unhtmlgl)
   unhtml=replace(unhtml,unhtmlgl(i),"***")
   next
   end if
'unHtml=ubb(unHtml)
end if
end function

function ubb(content)
ubb=content
    nowtime=now()
    UBB=Convert(ubb,"code")
    UBB=Convert(ubb,"html")
    UBB=Convert(ubb,"url")
    UBB=Convert(ubb,"color")
    UBB=Convert(ubb,"font")
    UBB=Convert(ubb,"size")
    UBB=Convert(ubb,"quote")
    UBB=Convert(ubb,"email")
    UBB=Convert(ubb,"img")
    UBB=Convert(ubb,"swf")
	ubb=convert(ubb,"cen")
	ubb=convert(ubb,"rig")
    ubb=convert(ubb,"lef")
    ubb=convert(ubb,"center")

    UBB=AutoURL(ubb)
    ubb=replace(ubb,"[b]","<b>",1,-1,1)
    ubb=replace(ubb,"[/b]","</b>",1,-1,1)
    ubb=replace(ubb,"[i]","<i>",1,-1,1)
    ubb=replace(ubb,"[/i]","</i>",1,-1,1)
    ubb=replace(ubb,"[u]","<u>",1,-1,1)
    ubb=replace(ubb,"[/u]","</u>",1,-1,1)
    ubb=replace(ubb,"[blue]","<font color='#000099'>",1,-1,1)
    ubb=replace(ubb,"[/blue]","</font>",1,-1,1)
    ubb=replace(ubb,"[red]","<font color='#990000'>",1,-1,1)
    ubb=replace(ubb,"[/red]","</font>",1,-1,1)
    for i=1 to 28
    ubb=replace(ubb,"{:em"&i&"}","<IMG SRC=emot/emotface/em"&i&".gif ></img>",1,6,1)
    ubb=replace(ubb,"{:em"&i&"}","",1,-1,1)
    next
    ubb=replace(ubb,"["&chr(176),"[",1,-1,1)
    ubb=replace(ubb,chr(176)&"]","]",1,-1,1)
    ubb=replace(ubb,"/"&chr(176),"/",1,-1,1)
'    ubb=replace(ubb,"{;em","{:em",1,-1,1)
end function


function Convert(ubb,CovT)
cText=ubb
startubb=1
do while Covt="url" or Covt="color" or Covt="font" or Covt="size"
startubb=instr(startubb,cText,"["&CovT&"=",1)
if startubb=0 then exit do
endubb=instr(startubb,cText,"]",1)
if endubb=0 then exit do
Lcovt=Covt
startubb=startubb+len(lCovT)+2
text=mid(cText,startubb,endubb-startubb)
codetext=replace(text,"[","["&chr(176),1,-1,1)
codetext=replace(codetext,"]",chr(176)&"]",1,-1,1)
'codetext=replace(codetext,"{:em","{;em",1,-1,1)
codetext=replace(codetext,"/","/"&chr(176),1,-1,1)
select case CovT
    case "color"
	cText=replace(cText,"[color="&text&"]","<font color='"&text&"'>",1,1,1)
	cText=replace(cText,"[/color]","</font>",1,1,1)
    case "font"
	cText=replace(cText,"[font="&text&"]","<font face='"&text&"'>",1,1,1)
	cText=replace(cText,"[/font]","</font>",1,1,1)
    case "size"
	if IsNumeric(text) then
	if text>6 then text=6
	if text<1 then text=1
	cText=replace(cText,"[size="&text&"]","<font size='"&text&"'>",1,1,1)
	cText=replace(cText,"[/size]","</font>",1,1,1)
	end if
    case "url"
	cText=replace(cText,"[url="&text&"]","<a href='"&codetext&"' target=_blank>",1,1,1)
	cText=replace(cText,"[/url]","</a>",1,1,1)
    case "email"
	cText=replace(cText,"["&CovT&"="&text&"]","<a href=mailto:"&text&">",1,1,1)
	cText=replace(cText,"[/"&CovT&"]","</a>",1,1,1)
end select
loop

startubb=1
do
startubb=instr(startubb,cText,"["&CovT&"]",1)
if startubb=0 then exit do
endubb=instr(startubb,cText,"[/"&CovT&"]",1)
if endubb=0 then exit do
Lcovt=Covt
startubb=startubb+len(lCovT)+2
text=mid(cText,startubb,endubb-startubb)
codetext=replace(text,"[","["&chr(176),1,-1,1)
codetext=replace(codetext,"]",chr(176)&"]",1,-1,1)
'codetext=replace(codetext,"{:em","{;em",1,-1,1)
codetext=replace(codetext,"/","/"&chr(176),1,-1,1)
select case CovT
    case "center"
    cText=replace(cText,"[center]","<div align='center'>",1,1,1)
	cText=replace(cText,"[/center]","</div>",1,1,1)

    case "url"
	cText=replace(cText,"["&CovT&"]"&text,"<a href='"&codetext&"' target=_blank>"&codetext,1,1,1)
	cText=replace(cText,"<a href='"&codetext&"' target=_blank>"&codetext&"[/"&CovT&"]","<a href="&codetext&" target=_blank>"&codetext&"</a>",1,1,1)
    case "email"
	cText=replace(cText,"["&CovT&"]","<a href=mailto:"&text&">",1,1,1)
	cText=replace(cText,"[/"&CovT&"]","</a>",1,1,1)
    case "html"
	codetext=replace(codetext,"<br>",chr(13),1,-1,1)
	codetext=replace(codetext,"&nbsp;",chr(32),1,-1,1)
	Randomize
	rid="temp"&Int(100000 * Rnd)
	cText=replace(cText,"[html]"&text,"代码片断如下：<TEXTAREA id="&rid&" rows=15 style='width:100%' class='bk'>"&codetext,1,1,1)
	cText=replace(cText,"代码片断如下：<TEXTAREA id="&rid&" rows=15 style='width:100%' class='bk'>"&codetext&"[/html]","代码片断如下：<TEXTAREA id="&rid&" rows=15 style='width:100%' class='bk'>"&codetext&"</TEXTAREA><INPUT onclick=runEx('"&rid&"') type=button value=运行此段代码 name=Button1 class='Tips_bo'> <INPUT onclick=JM_cc('"&rid&"') type=button value=复制到我的剪贴板 name=Button2 class='Tips_bo'>",1,1,1)
    case "img" '一般显示的图片
	cText=replace(cText,"[img]"&text,"<a href="&chr(34)&"about:<img src="&codetext&" border=0>"&chr(34)&" target=_blank><img src="&codetext ,1,1,1 )
	cText=replace(cText,"[/img]"," vspace=2 hspace=2 border=0 alt=::点击图片在新窗口中打开:: onload='javascript:if(this.width>580)this.width=580'></a>",1,1,1)
    
	case "cen" '图片居中
	cText=replace(cText,"[cen]"&text,"<table border='0' align='center' cellpadding='0' cellspacing='0'><tr><td > <a href="&chr(34)&"about:<img src="&codetext&" border=0>"&chr(34)&" target=_blank><img src="&codetext ,1,1,1 )

	cText=replace(cText,"[/cen]"," vspace=2 hspace=2 border=0 alt=::点击图片在新窗口中打开:: onload='javascript:if(this.width>580)this.width=580'></a></td></tr></table>",1,1,1)
	
	case "rig" '图片居右,文字绕排
	cText=replace(cText,"[rig]"&text,"<a href="&chr(34)&"about:<img src="&codetext&" border=0>"&chr(34)&" target=_blank><img src="&codetext ,1,1,1 )
	cText=replace(cText,"[/rig]"," vspace=2 hspace=2 border=0 align='right' alt=::点击图片在新窗口中打开:: onload='javascript:if(this.width>580)this.width=580'></a>",1,1,1)
   
    case "lef" '图片居左,文字绕排
	cText=replace(cText,"[lef]"&text,"<a href="&chr(34)&"about:<img src="&codetext&" border=0>"&chr(34)&" target=_blank><img src="&codetext ,1,1,1 )
	cText=replace(cText,"[/lef]"," vspace=2 hspace=2 border=0 align='left' alt=::点击图片在新窗口中打开:: onload='javascript:if(this.width>580)this.width=580'></a>",1,1,1)

	case "code"
	cText=replace(cText,"[code]"&text,"以下内容为程序代码<hr noshade>"&codetext,1,1,1)
	cText=replace(cText,"以下内容为程序代码<hr noshade>"&codetext&"[/code]","以下内容为程序代码<hr noshade>"&codetext&"<hr noshade>",1,1,1)
    case "quote"
    atext=replace(text,"[cen]","",1,-1,1)
	atext=replace(text,"[/cen]","",1,-1,1)

	atext=replace(text,"[img]","",1,-1,1)
	atext=replace(atext,"[/img]","",1,-1,1)
	atext=replace(atext,"[swf]","",1,-1,1)
	atext=replace(atext,"[/swf]","",1,-1,1)
	atext=replace(atext,"[html]","",1,-1,1)
	atext=replace(atext,"[/html]","",1,-1,1)
'	atext=replace(atext,"{:em","{;em",1,-1,1)
	atext=SplitWords(atext,350)
	atext=replace(atext,chr(32),"&nbsp;",1,-1,1)
	cText=replace(cText,"[quote]"&text,"<blockquote><hr noshade>"&atext,1,1,1)
	cText=replace(cText,"<blockquote><hr noshade>"&atext&"[/quote]","<blockquote><hr noshade>"&atext&"<hr noshade></blockquote>",1,1,1)
    case "swf"
	
	cText=replace(cText,"[swf]"&text,"<object classid='clsid:D27CDB6E-AE6D-11cf-96B8-444553540000' codebase='http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0' width='500' height='400'><param name=movie value='"&codetext&"'><param name=quality value=high><embed src='"&codetext&"' quality=high pluginspage='http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash' type='application/x-shockwave-flash' width='500' height='400'>",1,1,1)

	cText=replace(cText,"<object classid='clsid:D27CDB6E-AE6D-11cf-96B8-444553540000' codebase='http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0' width='500' height='400'><param name=movie value='"&codetext&"'><param name=quality value=high><embed src='"&codetext&"' quality=high pluginspage='http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash' type='application/x-shockwave-flash' width='500' height='400'>"&"[/swf]","<object classid='clsid:D27CDB6E-AE6D-11cf-96B8-444553540000' codebase='http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0' width='500' height='400'><param name=movie value='"&codetext&"'><param name=quality value=high><embed src='"&codetext&"' quality=high pluginspage='http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash' type='application/x-shockwave-flash' width='500' height='400'>"&"</embed></object>",1,1,1)
end select
loop
Convert=cText
end function

function AutoURL(ubb)
cText=ubb
startubb=1
do
startubb=1
endubb_a=0
endubb_b=0
endubb=0
startubb=instr(startubb,cText,"http://",1)
if startubb=0 then exit do
endubb_b=instr(startubb,cText,"<",1)
endubb_a=instr(startubb,cText,"&nbsp;",1)

endubb=endubb_a

if endubb=0 then
endubb=endubb_b
end if

if endubb_b<endubb and endubb_b>0 then
endubb=endubb_b
end if

if endubb=0 then
lenc=ctext
endubb=len(lenc)+1
end if

'response.write startubb&","&endubb
if startubb>endubb then exit do
text=mid(cText,startubb,endubb-startubb)
'response.write text
'codetext=replace(text,"/","/"&chr(176),1,-1,1)
codetext=text
'response.write text&","
urllink="<a href='"&codetext&"' target=_blank>"&codetext&"</a> "
'response.write urllink
urllink=replace(urllink,"/","/"&chr(176),1,-1,1)
cText=replace(cText,text,urllink,1,1,1)
loop
AutoURL=cText
end function
%>
