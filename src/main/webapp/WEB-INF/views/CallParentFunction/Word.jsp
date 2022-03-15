<%@ page language="java" pageEncoding="utf-8" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html>
<head>
</head>
<body>
<script type="text/javascript">
    function Close() {
        window.external.close();
    }

    function increaseCount(value) {
        var sResult = window.external.CallParentFunc("updateCount(" + value + ");");
        // if (sResult == 'poerror:parentlost') {
        //     alert('父页面关闭或跳转刷新了，导致父页面函数没有调用成功！');
        //     return false;
        // }
        document.getElementById("PageOfficeCtrl1").Alert("现在父窗口Count的值为：" + sResult);
    }

    function increaseCountAndClose(value) {
        var sResult = window.external.CallParentFunc("updateCount(" + value + ");");
        if (sResult == 'poerror:parentlost') {
            alert('父页面关闭或跳转刷新了，导致父页面函数没有调用成功！');
        }
        window.external.close();
    }
</script>
<input type="button" value="设置父窗口Count的值加1" onclick="increaseCount(1);"/>
<input type="button" value="设置父窗口Count的值加5，并关闭窗口" onclick="increaseCountAndClose(5);"/></br>
<div style=" width:auto; height:700px;">
    ${pageoffice}
</div>
</body>
</html>

