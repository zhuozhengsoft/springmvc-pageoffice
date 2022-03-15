package com.zhuozhengsoft.Samples5.controller;

import com.zhuozhengsoft.pageoffice.FileSaver;
import com.zhuozhengsoft.pageoffice.OpenModeType;
import com.zhuozhengsoft.pageoffice.PageOfficeCtrl;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.servlet.ModelAndView;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.OutputStream;
import java.sql.*;
import java.util.Map;

@Controller
@RequestMapping("DataBase")
public class DataBaseController {

    @RequestMapping("/Openstream")
    public void Openstream(HttpServletRequest request, HttpServletResponse response,Map<String, Object> map) throws Exception{
        String id = "2";
        if (request.getParameter("id") != null
                && request.getParameter("id").trim().length() > 0) {
            id = request.getParameter("id");
        }
        Class.forName("org.sqlite.JDBC");
        String strUrl = "jdbc:sqlite:"
                + request.getServletContext().getRealPath("demodata/") + "/DataBase.db";
        Connection conn = DriverManager.getConnection(strUrl);
        Statement stmt = conn.createStatement();
        ResultSet rs = stmt.executeQuery("select * from stream where id = "
                + id);
        int newID = 1;
        if (rs.next()) {
            //******读取磁盘文件，输出文件流 开始*******************************
            byte[] imageBytes = rs.getBytes("Word");
            int fileSize = imageBytes.length;

            response.reset();
            response.setContentType("application/msword"); // application/x-excel, application/ms-powerpoint, application/pdf
            response.setHeader("Content-Disposition",
                    "attachment; filename=down.doc"); //fileN应该是编码后的(utf-8)
            response.setContentLength(fileSize);

            OutputStream outputStream = response.getOutputStream();
            outputStream.write(imageBytes);

            outputStream.flush();
            outputStream.close();
            outputStream = null;
            //******读取磁盘文件，输出文件流 结束*******************************
        }
        rs.close();
        conn.close();
    }
    @RequestMapping("/Stream")
    public ModelAndView openWord(HttpServletRequest request, Map<String, Object> map) throws Exception{
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage(request.getContextPath() + "/poserver.zz"); //此行必须
        //隐藏菜单栏
        poCtrl.setMenubar(false);
        poCtrl.addCustomToolButton("保存", "Save()", 1);
        poCtrl.setSaveFilePage("save?id=1");
        //打开Word文件
        poCtrl.webOpen("Openstream?id=1", OpenModeType.docNormalEdit, "张三");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        ModelAndView mv = new ModelAndView("DataBase/Stream");
        return   mv;
    }

    @RequestMapping("/save")
    public void save(HttpServletRequest request, HttpServletResponse response) throws Exception{
        FileSaver fs = new FileSaver(request, response);
        if (request.getParameter("id") != null
                && request.getParameter("id").trim().length() > 0) {
            String id = request.getParameter("id").trim();
            Class.forName("org.sqlite.JDBC");
            String strUrl = "jdbc:sqlite:"
                    + request.getServletContext().getRealPath("demodata/") + "\\DataBase.db";
            Connection conn = DriverManager.getConnection(strUrl);
            String sql = "UPDATE  Stream SET Word=?  where ID=" + id;
            PreparedStatement pstmt = null;
            pstmt = conn.prepareStatement(sql);
            pstmt.setBytes(1, fs.getFileBytes());
            //pstmt.setBinaryStream(1, fs.getFileStream(),fs.getFileSize());
            pstmt.executeUpdate();
            pstmt.close();
            conn.close();

            fs.setCustomSaveResult("ok");
        } else {
            response.setContentType("text/plain; charset=utf-8");
            response.getWriter().write("未获得文件的ID，保存失败");
            fs.showPage(500, 400);
        }
        fs.close();
    }
}