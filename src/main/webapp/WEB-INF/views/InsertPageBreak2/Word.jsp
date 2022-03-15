<%@ page language="java" pageEncoding="utf-8" %>


<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
    <title>在word文档中光标处插入分页符</title>
</head>
<body>
<script type="text/javascript">
    function Save() {
        document.getElementById("PageOfficeCtrl1").WebSave();
        if (document.getElementById("PageOfficeCtrl1").CustomSaveResult == "ok") {
            alert("保存成功！请在/Samples4/InsertPageBreak2/doc目录下查看合并后的新文档\"test3.doc\"。");
        }
    }
</script>
<div style=" width:auto; height:700px;">
    ${pageoffice}
</div>
</body>
</html>
