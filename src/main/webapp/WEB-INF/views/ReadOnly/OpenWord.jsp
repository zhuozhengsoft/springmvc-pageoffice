<%@ page language="java" pageEncoding="utf-8" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
    <title>演示：文件在线安全浏览</title>
    <meta http-equiv="pragma" content="no-cache">
    <meta http-equiv="cache-control" content="no-cache">
    <meta http-equiv="expires" content="0">
    <meta http-equiv="keywords" content="keyword1,keyword2,keyword3">
    <meta http-equiv="description" content="This is my page">
    <script type="text/javascript">
        function AfterDocumentOpened() {
            document.getElementById("PageOfficeCtrl1").SetEnableFileCommand(4, false); //禁止另存
            document.getElementById("PageOfficeCtrl1").SetEnableFileCommand(5, false); //禁止打印
            document.getElementById("PageOfficeCtrl1").SetEnableFileCommand(6, false); //禁止页面设置
            document.getElementById("PageOfficeCtrl1").SetEnableFileCommand(8, false); //禁止打印预览
        }
    </script>
</head>
<body>

<div style=" width:auto; height:700px;">
   ${pageoffice}
</div>
</body>
</html>
