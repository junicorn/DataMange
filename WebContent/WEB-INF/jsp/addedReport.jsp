<%@ page language="java" import="java.util.*" pageEncoding="UTF-8"%>
<%
String path = request.getContextPath();
String basePath = request.getScheme()+"://"+request.getServerName()+":"+request.getServerPort()+path+"/";
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
  <head>
    <base href="<%=basePath%>">
    
    <title>Insert title here<</title>
    <meta http-equiv="Content-Type" content="text/html; charset=UTF-8" /> 
	

  </head>
  
  <body>
    <h1>导入成功!</h1>  
  
<a href="/DataManage/report">返回</a>  
  </body>
</html>
