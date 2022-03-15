<%@ page contentType="text/html;charset=UTF-8" language="java" %>
<html>
<head>
    <meta charset="UTF-8">
    <title>CommentOnly</title>
</head>
<body style="overflow:hidden">
<script type="text/javascript">
    function Save() {
        document.getElementById("PageOfficeCtrl1").WebSave();
    }

    function newComment() {
        var docSel = document.getElementById("PageOfficeCtrl1").Document.Application.Selection;
        docSel.Comments.Add(docSel.Range);
    }
</script>
<div style=" width:1100px; height:800px;">
    ${pageoffice}
</div>
</body>
</html>
