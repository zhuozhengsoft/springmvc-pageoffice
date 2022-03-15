<%@ page language="java"
         import="com.zhuozhengsoft.pageoffice.OpenModeType,com.zhuozhengsoft.pageoffice.PageOfficeCtrl,com.zhuozhengsoft.pageoffice.ThemeType"
         pageEncoding="utf-8" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
    <title>演示：修改Word文档的主题样式</title>
    <style>
        html, body {
            height: 100%;
        }

        .main {
            height: 100%;
        }
    </style>
</head>
<body>
<form id="form1">
    <div style="font-size: 12px; line-height: 20px; border-bottom: dotted 1px #ccc; border-top: dotted 1px #ccc;
        padding: 5px;">
        操作：在Html页面添加PageOfficeCtrl控件，再在后台设置PageOfficeCtrl的Theme属性<br/>
        关键代码：<span style="background-color:Yellow;">poCtrl.setTheme(ThemeType.Office2007);//poCtrl为PageOfficeCtrl对象，ThemeType为枚举类型</span>
    </div>
    <div style="height: 600px; width: auto;">
        ${pageoffice}
    </div>
</form>
</body>
</html>
