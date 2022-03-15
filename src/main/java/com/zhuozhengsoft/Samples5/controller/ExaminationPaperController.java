package com.zhuozhengsoft.Samples5.controller;

import com.zhuozhengsoft.pageoffice.*;
import com.zhuozhengsoft.pageoffice.wordwriter.DataRegion;
import com.zhuozhengsoft.pageoffice.wordwriter.DataRegionInsertType;
import com.zhuozhengsoft.pageoffice.wordwriter.WordDocument;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.servlet.ModelAndView;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.OutputStream;
import java.sql.*;
import java.util.Map;

@Controller
@RequestMapping("ExaminationPaper")
public class ExaminationPaperController {

    @RequestMapping("/Default")
    public ModelAndView openWord1(HttpServletRequest request, Map<String, Object> map) throws Exception{
        Class.forName("org.sqlite.JDBC");
        String strUrl = "jdbc:sqlite:"
                + request.getServletContext().getRealPath("demodata/") + "/ExaminationPaper.db";
        Connection conn = DriverManager.getConnection(strUrl);
        Statement stmt = conn.createStatement();
        ResultSet rs = stmt.executeQuery("Select * from stream");
        boolean flg = false;//标识是否有数据
        StringBuilder strHtmls = new StringBuilder();
        strHtmls.append("<tr  style='background-color:#FEE;'>");
        strHtmls.append("<td style='text-align:center;width=10%' >选择</td>");
        strHtmls.append("<td style='text-align:center;width=30%'>题库编号</td>");
        strHtmls.append("<td style='text-align:center;width=60%'>操作</td>");
        strHtmls.append("</tr>");
        while (rs.next()) {
            flg = true;
            String pID = rs.getString("ID");
            strHtmls.append("<tr  style='background-color:white;'>");
            strHtmls.append("<td style='text-align:center'><input id='check" + pID + "'  type='checkbox' /></td>");
            strHtmls.append("<td style='text-align:center'>选择题-" + pID + "</td>");
            strHtmls.append("<td style='text-align:center'><a href='javascript:POBrowser.openWindowModeless(\"Edit?id=" + pID + "\" ,\"width=1200px;height=800px;\");'>编辑</a></td>");
            strHtmls.append("</tr>");
        }

        if (!flg) {
            strHtmls.append("<tr>\r\n");
            strHtmls.append("<td width='100%' height='100' align='center'>对不起，暂时没有可以操作的数据。\r\n");
            strHtmls.append("</td></tr>\r\n");
        }
        request.setAttribute("strHtmls",strHtmls);
        ModelAndView mv = new ModelAndView("ExaminationPaper/Default");
        return mv;
    }
    @RequestMapping("/Openfile")
    public void Openfile(HttpServletRequest request, HttpServletResponse response) throws Exception{
        String err = "";
        if (request.getParameter("id") != null
                && request.getParameter("id").trim().length() > 0) {
            String id = request.getParameter("id");
            Class.forName("org.sqlite.JDBC");
            String strUrl = "jdbc:sqlite:"
                    + request.getServletContext().getRealPath("demodata/") + "\\ExaminationPaper.db";
            Connection conn = DriverManager.getConnection(strUrl);
            Statement stmt = conn.createStatement();
            String strSql = "select * from stream where id =" + id;
            ResultSet rs = stmt.executeQuery(strSql);
            if (rs.next()) {
                //******读取磁盘文件，输出文件流 开始*******************************
                byte[] imageBytes = rs.getBytes("Word");
                int fileSize = imageBytes.length;

                response.reset();
                response.setContentType("application/msword"); // application/x-excel, application/ms-powerpoint, application/pdf
                response.setHeader("Content-Disposition", "attachment; filename=down.doc"); //fileN应该是编码后的(utf-8)
                response.setContentLength(fileSize);
                OutputStream outputStream = response.getOutputStream();
                outputStream.write(imageBytes);
                outputStream.flush();
                outputStream.close();
                outputStream = null;

                //下面两句代码解决response.getWriter()和response.getOutputStream()冲突问题
                /*out.clear();
                out = pageContext.pushBody();*/
                //******读取磁盘文件，输出文件流 结束*******************************
            } else {
                err = "未获得文件的信息";
            }
            rs.close();
            stmt.close();
            conn.close();
        } else {
            err = "未获得文件的ID";
            //out.print(err);
        }
        if (err.length() > 0) {
            response.setContentType("text/plain; charset=utf-8");
            response.getWriter().write("<script>alert(" + err + ");</script>");
            //err = "<script>alert(" + err + ");</script>";
        }
    }
    @RequestMapping("/Edit")
    public ModelAndView openWord(HttpServletRequest request, Map<String, Object> map) throws Exception{
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage(request.getContextPath() + "/poserver.zz"); //此行必须
        //隐藏菜单栏
        poCtrl.setMenubar(false);
        poCtrl.addCustomToolButton("保存", "Save()", 1);
        //设置保存页面
        String id = request.getParameter("id");
        poCtrl.setSaveFilePage("save?id=" + id);
        //打开Word文件
        poCtrl.webOpen("Openfile?id=" + id, OpenModeType.docNormalEdit, "张三");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        ModelAndView mv = new ModelAndView("ExaminationPaper/Edit");
        return mv;
    }

    @RequestMapping("/save")
    public void save(HttpServletRequest request, HttpServletResponse response) throws Exception{
        FileSaver fs = new FileSaver(request, response);
        String err = "";
        if (request.getParameter("id") != null
                && request.getParameter("id").trim().length() > 0) {
            String id = request.getParameter("id").trim();
            Class.forName("org.sqlite.JDBC");
            String strUrl = "jdbc:sqlite:"
                    + request.getServletContext().getRealPath("demodata/") + "\\ExaminationPaper.db";
            Connection conn = DriverManager.getConnection(strUrl);
            String sql = "UPDATE  Stream SET Word=?  where ID=" + id;
            PreparedStatement pstmt = null;
            pstmt = conn.prepareStatement(sql);
            pstmt.setBytes(1, fs.getFileBytes());
            //pstmt.setBinaryStream(1,fs.getFileStream(),fs.getFileSize());
            pstmt.executeUpdate();
            pstmt.close();
            conn.close();
            fs.setCustomSaveResult("ok");
        } else {
            //err = "<script>alert('未获得文件的ID，保存失败');</script>";
            response.setContentType("text/plain; charset=utf-8");
            response.getWriter().write("<script>alert('未获得文件的ID，保存失败');</script>");
        }
        fs.close();
    }
    @RequestMapping("/Compose")
    public ModelAndView Compose(HttpServletRequest request, Map<String, Object> map) throws Exception{
        if (request.getParameter("ids").equals(null)
                || request.getParameter("ids").equals("")) {
            ModelAndView mv = new ModelAndView("ExaminationPaper/Default");
            return mv;
        }
        String idlist = request.getParameter("ids").trim();
        String[] ids = idlist.split(",");//将idlist按照","截取后存到ids数组中，然后遍历数组用js插入文件即可

        int pNum = 1;
        String operateStr = "";
        operateStr += "function Create(){\n";
        // document.getElementById('PageOfficeCtrl1').Document.Application 微软office VBA对象的根Application对象
        operateStr += "var obj = document.getElementById('PageOfficeCtrl1').Document.Application;\n";
        operateStr += "obj.Selection.EndKey(6);\n"; // 定位光标到文档末尾

        for (int i = 0; i < ids.length; i++) {
            operateStr += "obj.Selection.TypeParagraph();"; //用来换行
            operateStr += "obj.Selection.Range.Text = '" + pNum + ".';\n"; // 用来生成题号
            // 下面两句代码用来移动光标位置
            operateStr += "obj.Selection.EndKey(5,1);\n";
            operateStr += "obj.Selection.MoveRight(1,1);\n";
            // 插入指定的题到文档中
            operateStr += "document.getElementById('PageOfficeCtrl1').InsertDocumentFromURL('Openfile?id="
                    + ids[i] + "');\n";
            pNum++;

        }
        operateStr += "\n}\n";
        //******************************卓正PageOffice组件的使用*******************************
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage(request.getContextPath() + "/poserver.zz"); //此行必须
        //隐藏菜单栏
        poCtrl.setMenubar(false);
        poCtrl.setCustomToolbar(false);
        poCtrl.setCaption("生成试卷");
        poCtrl.setJsFunction_AfterDocumentOpened("Create()");
        //打开Word文件
        poCtrl.webOpen("/doc/ExaminationPaper/test.doc", OpenModeType.docNormalEdit, "张三");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        map.put("operateStr", operateStr);
        ModelAndView mv = new ModelAndView("ExaminationPaper/Compose");
        return mv;
    }
    @RequestMapping("/Compose2")
    public ModelAndView Compose2(HttpServletRequest request, Map<String, Object> map) throws Exception{
        if (request.getParameter("ids").equals(null)
                || request.getParameter("ids").equals("")) {
            ModelAndView mv = new ModelAndView("ExaminationPaper/Default");
            return mv;
        }
        String idlist = request.getParameter("ids").trim();
        String[] ids = idlist.split(","); //将idlist按照","截取后存到ids数组中，然后遍历数组用js插入文件
        String temp = "PO_begin";//存储数据区域名称
        int num = 1;//试题编号
        WordDocument doc = new WordDocument();
        for (int i = 0; i < ids.length; i++) {
            DataRegion dataNum = doc.createDataRegion("PO_" + num,
                    DataRegionInsertType.After, temp);
            dataNum.setValue(num + ".\t");
            DataRegion dataRegion = doc.createDataRegion("PO_begin"
                    + (i + 1), DataRegionInsertType.After, "PO_" + num);
            dataRegion.setValue("[word]Openfile?id=" + ids[i]
                    + "[/word]");
            temp = "PO_begin" + (i + 1);
            num++;
        }

        //******************************卓正PageOffice组件的使用*******************************
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage(request.getContextPath() + "/poserver.zz"); //此行必须
        //隐藏菜单栏
        poCtrl.setMenubar(false);
        poCtrl.setCustomToolbar(false);
        poCtrl.setCaption("生成试卷");
        poCtrl.setWriter(doc);
        //打开Word文件
        poCtrl.webOpen("/doc/ExaminationPaper/test.doc", OpenModeType.docNormalEdit, "张三");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        ModelAndView mv = new ModelAndView("ExaminationPaper/Compose2");
        return mv;
    }
}