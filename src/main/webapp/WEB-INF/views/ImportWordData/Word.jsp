<%@ page language="java" pageEncoding="utf-8" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
    <title>导入文件并提交数据</title>
    <meta http-equiv="pragma" content="no-cache">
    <meta http-equiv="cache-control" content="no-cache">
    <meta http-equiv="expires" content="0">
    <meta http-equiv="keywords" content="keyword1,keyword2,keyword3">
    <meta http-equiv="description" content="This is my page">
</head>
<body>
<!-- ***************************************PageOffice组件的使用****************************************** -->
<script type="text/javascript">
    function importData() {
        document.getElementById("PageOfficeCtrl1").WordImportDialog();
    }

    function submitData() {
        document.getElementById("PageOfficeCtrl1").WebSave();

    }
</script>
<div style="color:Red">请导入“/doc/ImportWordData”下的ImportWord.doc文档查看演示效果。</div>
<div style="width: auto; height: 600px;">
    ${pageoffice}
</div>
<!-- ***************************************PageOffice组件的使用****************************************** -->
</body>
</html>
