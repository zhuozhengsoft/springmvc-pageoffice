﻿<%@ page language="java" pageEncoding="utf-8" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
</head>
<body style="overflow:hidden">
<!--**************   卓正 PageOffice 客户端代码开始    ************************-->
<script language="javascript" type="text/javascript">
    function AfterDocumentOpened() {
        //alert(document.getElementById("PDFCtrl1").Caption);
    }

    function SetBookmarks() {
        document.getElementById("PDFCtrl1").BookmarksVisible = !document.getElementById("PDFCtrl1").BookmarksVisible;
    }

    function PrintFile() {
        document.getElementById("PDFCtrl1").ShowDialog(4);
    }

    function SwitchFullScreen() {
        document.getElementById("PDFCtrl1").FullScreen = !document.getElementById("PDFCtrl1").FullScreen;
    }

    function SetPageReal() {
        document.getElementById("PDFCtrl1").SetPageFit(1);
    }

    function SetPageFit() {
        document.getElementById("PDFCtrl1").SetPageFit(2);
    }

    function SetPageWidth() {
        document.getElementById("PDFCtrl1").SetPageFit(3);
    }

    function ZoomIn() {
        document.getElementById("PDFCtrl1").ZoomIn();
    }

    function ZoomOut() {
        document.getElementById("PDFCtrl1").ZoomOut();
    }

    function FirstPage() {
        document.getElementById("PDFCtrl1").GoToFirstPage();
    }

    function PreviousPage() {
        document.getElementById("PDFCtrl1").GoToPreviousPage();
    }

    function NextPage() {
        document.getElementById("PDFCtrl1").GoToNextPage();
    }

    function LastPage() {
        document.getElementById("PDFCtrl1").GoToLastPage();
    }

    function SetRotateRight() {
        document.getElementById("PDFCtrl1").RotateRight();
    }

    function SetRotateLeft() {
        document.getElementById("PDFCtrl1").RotateLeft();
    }
</script>
<div style="height:850px;width:auto;">
    ${pageoffice}
</div>
</body>
</html>