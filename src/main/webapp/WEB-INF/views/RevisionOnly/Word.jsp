<%@ page language="java" pageEncoding="utf-8" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html>
<head>
    <meta charset="utf-8">
    <title>XX文档系统</title>
    <style>
        #main {
            width: 1040px;
            height: 890px;
            border: #83b3d9 2px solid;
            background: #f2f7fb;

        }
        #shut {
            width: 45px;
            height: 30px;
            float: right;
            margin-right: -1px;
        }
        #shut:hover {
        }
    </style>
</head>
<body style="margin:0; padding:0;border:0px; overflow:hidden" scroll="no">
<script type="text/javascript">
    function Save() {
        document.getElementById("PageOfficeCtrl1").WebSave();
    }

    function showRevision() {
        document.getElementById("PageOfficeCtrl1").ShowRevisions = true;
    }

    function hideRevision() {
        document.getElementById("PageOfficeCtrl1").ShowRevisions = false;
    }
</script>
<div id="main">
    <div id="content" style="height:850px;width:1036px;overflow:hidden;">
        ${pageoffice}
    </div>
</div>
</body>
</html>

