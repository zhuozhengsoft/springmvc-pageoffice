<%@ page language="java"
         pageEncoding="utf-8" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
    <title>获取index.jsp页面传递过来的参数</title>
</head>
<body>
<script type="text/javascript">
    function Save() {
        document.getElementById("PageOfficeCtrl1").WebSave();
    }
</script>
(1)获取index.jsp用session传递过来的userName的值：</br>
&nbsp;&nbsp;&nbsp;代码：String userName=(String)session.getAttribute("userName");</br>
&nbsp;&nbsp;&nbsp;输出：UserName=${userName} </br></br>

(2)获取index.jsp用？传递过来的id的值：</br>
&nbsp;&nbsp;&nbsp;代码：String id=request.getParameter("id");</br>
&nbsp;&nbsp;&nbsp;输出：id=${id}</br>
<form id="form1">
    <div style=" width:auto; height:700px;">
        ${pageoffice}
    </div>
</form>
</body>
</html>
