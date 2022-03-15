<%@ page language="java" pageEncoding="utf-8" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
    <title>Excel只读模式下调整行高和列宽</title>
</head>
<body>
<script type="text/javascript">
    function Save() {
        document.getElementById("PageOfficeCtrl1").WebSave();
    }
</script>
<form id="form1">
    <div>
        设置当工作表只读时，允许用户手动调整行列。</br>
        <div style="color:Red;">sheet1.setAllowAdjustRC(true);</div>
    </div>
    <div style=" width:100%; height:700px;">
        ${pageoffice}
    </div>
</form>
</body>
</html>
