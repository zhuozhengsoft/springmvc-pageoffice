<%@ page language="java"
         import="com.zhuozhengsoft.pageoffice.OpenModeType,com.zhuozhengsoft.pageoffice.PageOfficeCtrl"
         pageEncoding="utf-8" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
    <title>Word文档保存后获取返回值</title>
    <script type="text/javascript">
        function Save() {
            document.getElementById("PageOfficeCtrl1").WebSave();
            //document.getElementById("PageOfficeCtrl1").CustomSaveResult获取的是保存页面的返回值
            document.getElementById("PageOfficeCtrl1").Alert("保存成功，返回值为：" + document.getElementById("PageOfficeCtrl1").CustomSaveResult);
        }

        function Close() {
            window.external.close();
        }
    </script>
</head>
<body>
<form id="form1">
    <div style=" font-size:small; color:Red;">
        <label>键代码：点右键，选择“查看源文件”，看js函数“Save()</label>
        <br/>document.getElementById("PageOfficeCtrl1").WebSave()//执行保存操作"
        <br/>document.getElementById("PageOfficeCtrl1").CustomSaveResult//获取返回值保存页面SaveFile.jsp代码fs.setCustomSaveResult("ok");定义的返回值
        <br/>
    </div>
    <div style=" width:auto; height:700px;">
        ${pageoffice}
    </div>
</form>
</body>
</html>
