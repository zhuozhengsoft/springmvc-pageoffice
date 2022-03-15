<%@ page language="java"
         import="com.zhuozhengsoft.pageoffice.OpenModeType,com.zhuozhengsoft.pageoffice.PageOfficeCtrl"
         pageEncoding="utf-8" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
    <title>在自定义工具条添加按钮</title>
</head>
<body>
<script type="text/javascript">
    function myTest() {
        document.getElementById("PageOfficeCtrl1").Alert("测试成功！");
    }
</script>
<form id="form1">
    点击自定义工具栏中的“测试按钮”查看效果。<br/>
    <img src="../images/CustomToolButton/addbutton.jpg"/>
    <div style=" width:auto; height:700px;">
       ${pageoffice}
    </div>
</form>
</body>
</html>

