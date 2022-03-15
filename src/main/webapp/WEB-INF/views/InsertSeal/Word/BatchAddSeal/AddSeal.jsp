<%@ page language="java" pageEncoding="utf-8" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
    <title>My JSP 'FileMaker.jsp' starting page</title>
    <meta http-equiv="pragma" content="no-cache">
    <meta http-equiv="cache-control" content="no-cache">
    <meta http-equiv="expires" content="0">
    <meta http-equiv="keywords" content="keyword1,keyword2,keyword3">
    <meta http-equiv="description" content="This is my page">
    <script type="text/javascript">
        //标识盖章是否成功
        var isAddSealSuc = false;

        function AfterDocumentOpened() {
            try {
                //先定位到印章位置,再在印章位置上盖章
                document.getElementById("FileMakerCtrl1").ZoomSeal.LocateSealPosition("Seal1");
                isAddSealSuc = document.getElementById("FileMakerCtrl1").ZoomSeal.AddSealByName("测试公章", null, "Seal1"); //位置名称不区分大小写
            } catch (e) {
            }
        }

        function OnProgressComplete() {
            window.parent.convert(isAddSealSuc); //调用父页面的js函数
        }
    </script>
</head>
<body>
<div>
    ${pageoffice}
</div>
</body>
</html>