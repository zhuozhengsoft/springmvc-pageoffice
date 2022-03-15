<%@ page language="java" pageEncoding="utf-8" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
    <title>后台编程插入Word文件到数据区域</title>
</head>
<body>
<div style="font-size: 12px; line-height: 20px; border-bottom: dotted 1px #ccc; border-top: dotted 1px #ccc;
        padding: 5px;">
    关键代码：<span style="background-color: Yellow;"> <br/>DataRegion dataRegion
            = worddoc.openDataRegion("PO_开头的书签名称");
            <br/>
             dataRegion.setValue("[word]/doc/WordResWord/1.doc[/word]");</span><br/>
</div>
<br/>
<form id="form1">
    <div style="width: auto; height: 700px;">
        <!--**************   PageOffice 客户端代码开始    ************************-->
        ${pageoffice}
        <!--**************   PageOffice 客户端代码结束    ************************-->
    </div>
</form>
</body>
</html>