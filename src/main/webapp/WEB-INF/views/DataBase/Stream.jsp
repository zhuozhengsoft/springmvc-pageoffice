<%@ page language="java" pageEncoding="utf-8" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
    <title></title>
</head>
<body>
<form id="form1">
    <div style="width: auto; height: 700px;">
        <!-- *********************PageOffice组件客户端JS代码*************************** -->
        <script type="text/javascript">
            function Save() {
                document.getElementById("PageOfficeCtrl1").WebSave();
                if (document.getElementById("PageOfficeCtrl1").CustomSaveResult == "ok") {
                    document.getElementById("PageOfficeCtrl1").Alert('保存成功！');
                }
            }
        </script>
        <!-- *********************PageOffice组件的引用*************************** -->
        ${pageoffice}
    </div>
</form>
</body>
</html>
