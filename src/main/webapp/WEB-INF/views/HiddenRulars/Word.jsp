<%@ page language="java" pageEncoding="utf-8" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
    <title>显示/隐藏标尺</title>
    <meta http-equiv="pragma" content="no-cache">
    <meta http-equiv="cache-control" content="no-cache">
    <meta http-equiv="expires" content="0">
    <meta http-equiv="keywords" content="keyword1,keyword2,keyword3">
    <meta http-equiv="description" content="This is my page">
</head>
<body>
<div style=" width:1200px; height:700px;">
    <!-- *********************************PageOffice组件的使用************************************ -->
    <script type="text/javascript" language="javascript">
        //显示/隐藏标尺
        function Hidden() {
            document.getElementById("PageOfficeCtrl1").Document.ActiveWindow.ActivePane.DisplayRulers =
                !document.getElementById("PageOfficeCtrl1").Document.ActiveWindow.ActivePane.DisplayRulers;
        }

    </script>
    ${pageoffice}
    <!-- *********************************PageOffice组件的使用************************************ -->
</div>
</body>
</html>
