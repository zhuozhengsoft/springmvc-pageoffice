<%@ page language="java" pageEncoding="utf-8" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
    <title>保存时获取word文档中的图片</title>
</head>
<body>
<script type="text/javascript">
    function Save() {
        document.getElementById("PageOfficeCtrl1").WebSave();
        var value = document.getElementById("PageOfficeCtrl1").CustomSaveResult;
        document.getElementById("PageOfficeCtrl1").Alert("保存成功,文件保存到：" + value);
    }
</script>
<div id="div1" style="width: auto; height: 700px;">
   ${pageoffice}
</div>
</body>
</html>
