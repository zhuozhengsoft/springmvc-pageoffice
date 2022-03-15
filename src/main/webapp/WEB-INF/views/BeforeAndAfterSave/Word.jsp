<%@ page language="java"
         import="com.zhuozhengsoft.pageoffice.OpenModeType,com.zhuozhengsoft.pageoffice.PageOfficeCtrl"
         pageEncoding="utf-8" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
    <title>文档保存前和保存后执行的事件</title>
</head>
<body>
<script type="text/javascript">
    function BeforeDocumentSaved() {
        document.getElementById("PageOfficeCtrl1").Alert("BeforeDocumentSaved事件：文件就要开始保存了.");
        return true;
    }

    function AfterDocumentSaved(IsSaved) {
        if (IsSaved) {
            document.getElementById("PageOfficeCtrl1").Alert("AfterDocumentSaved事件：文件保存成功了.");
        }
    }
</script>
<form id="form1">
    演示: 文档保存前和保存后执行的事件。<br/><br/>
    <div style=" width:auto; height:700px;">
        ${pageoffice}
    </div>
</form>
</body>
</html>
