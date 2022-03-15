<%@ page language="java" pageEncoding="utf-8" %>


<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
    <title>My JSP 'index.jsp' starting page</title>
</head>
<body>
<!-- *********************pageoffice组件的使用 **************************-->
<script language="javascript" type="text/javascript">
    //全屏
    function SetFullScreen() {
        document.getElementById("PageOfficeCtrl1").FullScreen = !document.getElementById("PageOfficeCtrl1").FullScreen;
    }
</script>
${pageoffice}
<!-- *********************pageoffice组件的使用 **************************-->
</body>
</html>
