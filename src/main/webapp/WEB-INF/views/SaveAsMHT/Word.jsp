<%@ page language="java"
         import="com.zhuozhengsoft.pageoffice.OpenModeType,com.zhuozhengsoft.pageoffice.PageOfficeCtrl"
         pageEncoding="utf-8" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
    <title>Word文件另存为MHT</title>
</head>
<body>
<script type="text/javascript">
    function saveAsMHT() {
        document.getElementById("PageOfficeCtrl1").WebSaveAsMHT();
        document.getElementById("PageOfficeCtrl1").Alert("MHT格式的文件已经保存到 doc\\SaveAsMHT 目录下。");
        document.getElementById("div1").innerHTML = "<a href='../../doc/SaveAsMHT/test.mht?r=" + Math.random() + "'> 查看另存的 MHT 文件<a><br><br>";
    }
</script>
<form id="form1">
    <div id="div1"></div>
    <div style=" width:1000px; height:800px;">
        ${pageoffice}
    </div>
</form>
</body>
</html>
