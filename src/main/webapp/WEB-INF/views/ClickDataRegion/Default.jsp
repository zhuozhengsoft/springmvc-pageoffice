<%@ page language="java" pageEncoding="utf-8" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
    <meta http-equiv="Content-Type" content="text/html; charset=gb2312"/>
    <title></title>
    <link href="images/csstg.css" rel="stylesheet" type="text/css"/>
</head>
<body>
<div id="content">
    <div id="textcontent" style="width: 1000px; height: 800px;">
        <script type="text/javascript">
            //***********************************PageOffice组件调用的js函数**************************************
            //保存页面
            function Save() {
                document.getElementById("PageOfficeCtrl1").WebSave();
            }

            //全屏/还原
            function IsFullScreen() {
                document.getElementById("PageOfficeCtrl1").FullScreen = !document.getElementById("PageOfficeCtrl1").FullScreen;
            }

            function OnWordDataRegionClick(Name, Value, Left, Bottom) {

                if (Name == "PO_deptName") {
                    var strRet = document.getElementById("PageOfficeCtrl1").ShowHtmlModalDialog("HTMLPage", Value, "left=" + Left + "px;top=" + Bottom + "px;width=400px;height=300px;frame=no;");
                    if (strRet != "") {
                        return (strRet);
                    }
                    else {
                        if ((Value == undefined) || (Value == ""))
                            return " ";
                        else
                            return Value;
                    }
                }
            }
        </script>
        <!--**************   卓正 PageOffice组件 ************************-->
        ${pageoffice}
    </div>
</div>
</body>
</html>
