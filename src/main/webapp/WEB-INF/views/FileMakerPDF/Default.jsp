<%@ page language="java" pageEncoding="utf-8" %>
<%
    String url = request.getSession().getServletContext().getRealPath("FileMakerPDF/doc/" + "/");
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
    <meta charset="UTF-8">
    <title>word转pdf</title>
    <!--pageoffice.js一定要引用--->
    <script type="text/javascript" src="../pageoffice.js"></script>
    <script type="text/javascript">
        function  myFunc(type){
            alert("转换完成！");
            document.getElementById("pgImg").style.display="none";
            if("1"==type){
                //下载生成的pdf文件
                location.href="DownPDF";
            }else if("2"==type){
                //打开pdf文件
                POBrowser.openWindowModeless('OpenPDF' , 'width=1200px;height=800px;');
            }

        }
        function convert(type) {
            document.getElementById("pgImg").style.display="block";
            POBrowser.openWindowModeless('FileMakerPDF?type='+type , 'width=1px;height=1px;frame=no;');
        }
    </script>
</head>
<body style="text-align: center;">
<a  href="javascript:;" onclick="convert(1)">1.Word文件转PDF文件并下载PDF文件</a><br><br>
<a  href="javascript:;" onclick="convert(2)">2.Word文件转PDF文件并打开PDF文件</a>
<div id="pgImg" style="with:100px;height:100px;margin-top:20px;display: none;"  >
    正在转换文件，请稍等：<img src="../images/FileMakerPDF/pg.gif">
</div>
</body>
</html>