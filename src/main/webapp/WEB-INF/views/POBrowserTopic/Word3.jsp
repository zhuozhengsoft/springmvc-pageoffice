<%@ page language="java"
         pageEncoding="utf-8" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
    <title>主页面传递参数到子页面</title>
    <script type="text/javascript">
        function Save() {
            document.getElementById("PageOfficeCtrl1").WebSave();
            window.external.close();
        }
    </script>
</head>
<body>
获取index.jsp中第一个文本框中的值：</br>
代码：String txt=(String)session.getAttribute("txt");</br>
输出：${txt} </br></br>
<form id="form1">
    <div style=" width:auto; height:700px;">
       ${pageoffice}
    </div>
</form>
</body>
</html>
