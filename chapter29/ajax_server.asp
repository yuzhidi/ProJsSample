<% @Language="JavaScript" %>
<%
function OpenDB(sdbname)
{
/*
*--------------- OpenDB(sdbname) -----------------
* OpenDB(sdbname) 
* 功能:打开数据库sdbname,返回conn对象.
* 参数:sdbname,字符串,数据库名称.
* 实例:var conn = OpenDB("database.mdb");
*--------------- OpenDB(sdbname) -----------------
*/
var connstr = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source="+Server.MapPath(sdbname);
var conn = Server.CreateObject("ADODB.Connection");
conn.Open(connstr);
return conn;
}
var oConn = OpenDB("ajax_data.mdb");
var sel = Request("sel");
var classid = Request("classid")
var fieldname = Request("fieldname")
var arrResult = new Array();
//var sql = "select "+fieldname+" from Demo where parentid='"+sel+"' and classid="+classid;
var sql = "select id,"+fieldname+" from Demo where parentid='"+sel+"'";
//Response.Write("alert("+sql+")")
var rs = Server.CreateObject("ADODB.Recordset");
rs.Open(sql,oConn,1,1);
while(!rs.EOF)
{
//遍历所有适合的数据放入arrResult数组中.
arrResult[arrResult.length] = rs(0).Value+"|"+rs(1).Value;
rs.MoveNext();
}
//escape解决了XMLHTTP。中文处理的问题.
//数组组合成字符串.由","字符串连接.
Response.Write(escape(arrResult.join(",")));
%>