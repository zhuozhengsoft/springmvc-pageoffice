<%@ page language="java" pageEncoding="utf-8" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
    <title></title>
    <script type="text/javascript">
        function Save() {
            document.getElementById("PageOfficeCtrl1").WebSave();
        }
    </script>
</head>
<body>
<a href="#" onclick="window.external.close();">返回列表页</a>
<form id="form1">
    <div style="width: auto; height: 600px;">
        ${pageoffice}
    </div>
</form>
</body>
</html>
