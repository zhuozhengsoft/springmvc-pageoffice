<%@ page language="java"
         import="com.zhuozhengsoft.pageoffice.OpenModeType,com.zhuozhengsoft.pageoffice.PageOfficeCtrl"
         pageEncoding="utf-8" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
    <title>文件打开后触发的事件</title>
</head>
<body>
<script type="text/javascript">
    function AfterDocumentOpened() {
        // 打开文件的时候，给word中当前光标位置赋值一个文本值
        document.getElementById("PageOfficeCtrl1").Document.Application.Selection.Range.Text = "文件打开了";
    }
</script>
<form id="form1">
    Word中的"<span style=" color:Red;"> 文件打开了</span>" 是在文档打开的事件中用程序添加进去的。<br/><br/>
    <div style=" width:auto; height:700px;">
       ${pageoffice}
    </div>
</form>
</body>
</html>
