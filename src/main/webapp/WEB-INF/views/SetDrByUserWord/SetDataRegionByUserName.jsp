<%@ page language="java"  pageEncoding="utf-8" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
    <meta http-equiv="Content-Type" content="text/html; charset=gb2312"/>
    <title></title>
    <link href="images/csstg.css" rel="stylesheet" type="text/css"/>
</head>
<body>
<div id="content">
    <div id="textcontent" style="width: 1000px; height: 800px;">
        <div class="flow4">
            <a href="Default"> 返回登录页</a>
            <strong>当前用户：</strong>
            <span style="color: Red;">${userName}</span>
        </div>
        <script type="text/javascript">
            //保存页面
            function Save() {
                document.getElementById("PageOfficeCtrl1").WebSave();
            }

            //全屏/还原
            function IsFullScreen() {
                document.getElementById("PageOfficeCtrl1").FullScreen = !document
                    .getElementById("PageOfficeCtrl1").FullScreen;
            }
        </script>
        <!--**************   卓正 PageOffice组件 ************************-->
        ${pageoffice}
    </div>
</div>

</body>
</html>

