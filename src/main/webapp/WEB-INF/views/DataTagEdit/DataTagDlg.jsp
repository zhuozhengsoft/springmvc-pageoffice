<%@ page language="java" pageEncoding="utf-8" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN"
        "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
    <title></title>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <style type="text/css">
        #tagTable td {
            height: 25px;
            border-bottom: dotted 1px gray;
        }
    </style>
    <script type="text/javascript">
        // 方法: window.external.CallParentFunc 
        // 作用: 调用父窗口中的js函数, 目前只支持传递一个参数. 
        var names = window.external.CallParentFunc("getTagNames", "");
        var tagArr = names.split(";");

        //首次加载数据
        function load() {
            searchBookMark('');
            return;
        }

        //加载数据列表
        function searchBookMark(s) {
            //删除所有行
            var tb1 = document.getElementById("tagTable");
            var rCount = tb1.rows.length;
            for (var i = 0; i < rCount; i++) {
                tb1.deleteRow(0);
            }

            var oTable = document.getElementById("tagTable");
            for (var i = 0; i < tagArr.length; i++) {
                if (tagArr[i] != null && tagArr[i] != "" && 0 == tagArr[i].toLocaleLowerCase().indexOf(s.toLocaleLowerCase())) {
                    var oTr = oTable.insertRow();
                    var oTd = oTr.insertCell();
                    oTd.innerHTML = tagArr[i];
                    oTd = oTr.insertCell();
                    oTd.innerHTML = "&nbsp;&nbsp;<a href=\"javascript:void(0);\"  onclick= \"add('" + tagArr[i] + "','aaaa')\">&nbsp;添加</a>";
                    oTd = oTr.insertCell();
                    oTd.innerHTML = "&nbsp;&nbsp;<a href=\"javascript:void(0);\"  onclick= \"locate('" + tagArr[i] + "')\">&nbsp;定位</a>";
                    oTd = oTr.insertCell();
                    oTd.innerHTML = "&nbsp;&nbsp;<a href=\"javascript:void(0);\"  onclick= \"del('" + tagArr[i] + "')\">&nbsp;删除</a>";
                }
            }
        }


        function Button1_onclick() {
            var s = document.getElementById("Text1").value.toLocaleLowerCase();
            var tb1 = document.getElementById("tagTable");
            var rCount = tb1.rows.length;
            for (var i = 0; i < rCount; i++) {
                tb1.deleteRow(0);
            }

            var oTable = document.getElementById("tagTable");
            for (var i = 0; i < tagArr.length; i++) {
                if (tagArr[i] != null && tagArr[i] != "" && tagArr[i].toLocaleLowerCase().indexOf(s) >= 0) {

                    var oTr = oTable.insertRow();
                    var oTd = oTr.insertCell();
                    oTd.innerHTML = tagArr[i];
                    oTd = oTr.insertCell();
                    oTd.innerHTML = "&nbsp;&nbsp;<a href=\"javascript:void(0);\"  onclick= \"add('" + tagArr[i] + "')\">&nbsp;添加</a>";
                    oTd = oTr.insertCell();
                    oTd.innerHTML = "&nbsp;&nbsp;<a href=\"javascript:void(0);\"  onclick= \"locate('" + tagArr[i] + "')\">&nbsp;定位</a>";
                    oTd = oTr.insertCell();
                    oTd.innerHTML = "&nbsp;&nbsp;<a href=\"javascript:void(0);\"  onclick= \"del('" + tagArr[i] + "')\">&nbsp;删除</a>";
                }
            }
        }

        //******** Tag 操作 ************************************************************
        function add(name) {
            if ("true" == window.external.CallParentFunc("addTag", name)) {

            }
        }

        function del(name) {//alert(name);
            if ("false" == window.external.CallParentFunc("delTag", name)) {
                alert("请先执行\"定位\"操作，然后再删除。");
            }
        }

        function locate(name) {

            window.external.CallParentFunc("locateTag", name);
        }
    </script>
</head>
<body>
<div style="width: 380px; height: 320px;">
    <div style="float: left;font-size:12px;">
        <span>待添加数据标签：</span><br/>
        <input id="Text1" type="text" onpropertychange="searchBookMark(this.value);"
               oninput="searchBookMark(this.value);"/>
        <input id="Button1" type="button" value="搜索" onclick="return Button1_onclick()"/>
        <div style="width: 380px; height: 300px; border: solid 1px #ccc; overflow-y: scroll;  ">
            <table id="tagTable" style=" font-size:12px;">
            </table>
        </div>
    </div>
</div>
</body>
<script type="text/javascript">
    load();
    //alert(2);
</script>
</html>
