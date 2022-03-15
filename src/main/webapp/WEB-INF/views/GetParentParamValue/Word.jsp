<%@ page language="java" pageEncoding="utf-8" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
    <title>父页面给子页面传递参数</title>
</head>
<body>
<script type="text/javascript">
    function AfterDocumentOpened() {
        var userName = window.external.UserParams;
        document.getElementById("userName").value =  decodeURI(userName);
    }
</script>
<div>
    <font color="red">父页面传递过来的参数:</font><input type="text" id="userName" name="userName"/>
</div>
<div style="width:auto; height:670px;">
    ${pageoffice}
</div>
</body>
</html>
