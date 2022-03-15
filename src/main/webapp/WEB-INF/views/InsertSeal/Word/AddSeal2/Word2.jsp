<%@ page language="java" pageEncoding="utf-8" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
    <title> 2.不弹出用户名、密码输入框盖章</title>

    <meta http-equiv="pragma" content="no-cache">
    <meta http-equiv="cache-control" content="no-cache">
    <meta http-equiv="expires" content="0">
    <meta http-equiv="keywords" content="keyword1,keyword2,keyword3">
    <meta http-equiv="description" content="This is my page">
</head>
<body>
<script type="text/javascript">
    //保存
    function Save() {
        document.getElementById("PageOfficeCtrl1").WebSave();
    }

    //加盖印章
    function InsertSeal() {
        try {
            document.getElementById("PageOfficeCtrl1").ZoomSeal.AddSeal("李志");
        } catch (e) {
        }
    }
</script>
<div style="width: auto; height: 700px;">
    ${pageoffice}
</div>
</body>

</html>