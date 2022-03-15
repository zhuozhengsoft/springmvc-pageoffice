<%@ page language="java" pageEncoding="utf-8" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
    <title>4.编辑模版 - 在模版中添加盖章位置</title>
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

    function InsertSealPos() {
        try {
            document.getElementById("PageOfficeCtrl1").ZoomSeal.AddSealPosition();
        } catch (e) {
        }
        ;
    }
</script>
<div style="width: auto; height: 700px;">
    ${pageoffice}
</div>
</body>
</html>