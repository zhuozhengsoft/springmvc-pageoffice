<%@ page language="java"
         import="com.zhuozhengsoft.pageoffice.OpenModeType,com.zhuozhengsoft.pageoffice.PageOfficeCtrl,com.zhuozhengsoft.pageoffice.wordwriter.DataRegion,com.zhuozhengsoft.pageoffice.wordwriter.WordDocument"
         pageEncoding="utf-8" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html >
<head>
    <title>最简单的提交Word中的数据</title>
    <meta http-equiv="Content-Type" content="text/html; charset=UTF-8"/>
    <script type="text/javascript">
        function Save() {
            document.getElementById("PageOfficeCtrl1").WebSave();
            //alert(document.getElementById("PageOfficeCtrl1").CustomSaveResult);
        }
    </script>
</head>
<body>
<form id="form1">
    <div>
        <span style="color: Red; font-size: 14px;">请输入公司名称、年龄、部门等信息后，单击工具栏上的保存按钮</span>
        <br/>
        <span style="color: Red; font-size: 14px;">请输入公司名称：</span>
        <input id="txtName" name="txtCompany" type="text"/>
        <br/>
    </div>
    <div style="width: auto; height: 700px;">
       ${pageoffice}
    </div>
</form>
</body>
</html>
