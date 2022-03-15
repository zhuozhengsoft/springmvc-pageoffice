<%@ page contentType="text/html;charset=UTF-8" language="java" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
    <title>OnWordSelectionChange的使用</title>
</head>
<body>
<div style="font-size: 12px; line-height: 20px; border-bottom: dotted 1px #ccc; border-top: dotted 1px #ccc; padding: 5px;">
    <span style="color: red;">操作说明：请选中文档中一段内容进行测试。</span>
</div>
<br/>
<script type="text/javascript">
    function OnWordSelectionChange() {
        var obj = document.getElementById("PageOfficeCtrl1").Document.Application.Selection;
        if (obj.Range.Text != "") {
            alert(obj.Range.Text);
        }
    }
</script>
<div style="width:auto; height:800px;">
    ${pageoffice}
</div>
</body>
</html>
