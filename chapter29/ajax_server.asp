<% @Language="JavaScript" %>
<%
function OpenDB(sdbname)
{
/*
*--------------- OpenDB(sdbname) -----------------
* OpenDB(sdbname) 
* ����:�����ݿ�sdbname,����conn����.
* ����:sdbname,�ַ���,���ݿ�����.
* ʵ��:var conn = OpenDB("database.mdb");
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
//���������ʺϵ����ݷ���arrResult������.
arrResult[arrResult.length] = rs(0).Value+"|"+rs(1).Value;
rs.MoveNext();
}
//escape�����XMLHTTP�����Ĵ��������.
//������ϳ��ַ���.��","�ַ�������.
Response.Write(escape(arrResult.join(",")));
%>