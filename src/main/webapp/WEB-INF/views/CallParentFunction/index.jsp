<%@ page language="java" pageEncoding="utf-8" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
    <title>POBrowser方式打开Word文档</title>
    <meta http-equiv="keywords" content="keyword1,keyword2,keyword3">
    <meta http-equiv="description" content="this is my page">
    <meta http-equiv="content-type" content="text/html; charset=utf-8">
    <!--pageoffice.js文件一定要引用-->
    <script type="text/javascript" src="../pageoffice.js"></script>
    <script type="text/javascript">
        var count = 0;
        function updateCount(value) {
            count = count + value;
            document.getElementById("Text1").value = count;
            return count.toString();
        }
    </script>
</head>
<body>
<a href="javascript:POBrowser.openWindowModeless('Word?id=123', 'width=1200px;height=800px;scroll=no;');">POBrowser方式打开Word文档</a><br><br><br>
<div>Count=<input id="Text1" type="text" value="0"/></div>
</body>
</html>



