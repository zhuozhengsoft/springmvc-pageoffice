﻿<%@ page language="java" pageEncoding="utf-8" %>


<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
    <title>4.不弹出用户名、密码输入框(手写)签字到模板中的指定位置</title>
    <meta http-equiv="pragma" content="no-cache">
    <meta http-equiv="cache-control" content="no-cache">
    <meta http-equiv="expires" content="0">
    <meta http-equiv="keywords" content="keyword1,keyword2,keyword3">
    <meta http-equiv="description" content="This is my page">
</head>
<body>
<script type="text/javascript">
    function Save() {
        document.getElementById("PageOfficeCtrl1").WebSave();
    }

    function AddHandSign() {
        try {
            //定位到印章位置
            document.getElementById("PageOfficeCtrl1").ZoomSeal.LocateSealPosition("Seal1");
            /**第一个参数，可选项，签字的用户名，为空字符串时，将弹出用户名密+密码框，如果为指定的签章用户名，则直接弹出签字框；
             * 第二个参数，可选项，标识是否保护文档，为null时保护文档，为空字符串时不保护文档;
             * 第三个参数，可选项，标识盖章指定位置名称，须为英文或数字，不区分大小写。
             */
            document.getElementById("PageOfficeCtrl1").ZoomSeal.AddHandSign("李志", null, "Seal1");
        } catch (e) {
        }
    }
</script>
<div style="width: auto; height: 700px;">
    ${pageoffice}
</div>
</body>
</html>