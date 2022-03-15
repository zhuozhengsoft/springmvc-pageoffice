<%@ page language="java" pageEncoding="utf-8" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <title></title>
    <link href="../css/TaoHong/csstg.css" rel="stylesheet" type="text/css"/>
</head>
<body>
<form id="form2">
    <div id="header">
        <div style="float: left; margin-left: 20px;">
            <img src="../images/TaoHong/logo.jpg" height="30"/></div>
        <ul>
            <li><a target="_blank" href="http://www.zhuozhengsoft.com">卓正网站</a></li>
            <li><a target="_blank" href="http://www.zhuozhengsoft.com/poask/index.asp">客户问吧</a></li>
            <li><a href="#">在线帮助</a></li>
            <li><a target="_blank" href="http://www.zhuozhengsoft.com/about/about/">联系我们</a></li>
        </ul>
    </div>
    <div id="content">
        <div id="textcontent" style="width: 1000px; height: 800px;">
            <div class="flow4">
                <a href="#" onclick="window.external.close();">
                    <img alt="返回" src="../images/TaoHong/return.gif" border="0"/>文件列表</a> <span style="width: 100px;">
                    </span><strong>文档主题：</strong> <span style="color: Red;">测试文件</span>

            </div>
            <!--**************   卓正 PageOffice组件 ************************-->
            <script language="javascript">
                function ShowDialog1() {
                    document.getElementById("PageOfficeCtrl1").ShowDialog(2);
                }

                function ShowDialog2() {
                    document.getElementById("PageOfficeCtrl1").ShowDialog(5);
                }

                function ShowDialog3() {
                    document.getElementById("PageOfficeCtrl1").ShowDialog(4);
                }

                //全屏/还原
                function IsFullScreen() {
                    document.getElementById("PageOfficeCtrl1").FullScreen = !document.getElementById("PageOfficeCtrl1").FullScreen;
                }
            </script>
           ${pageoffice}
        </div>
    </div>
    <div id="footer">
        <hr width="1000"/>
        <div>
            Copyright (c) 2019 北京卓正志远软件有限公司
        </div>
    </div>
</form>
</body>
</html>
