package com.zhuozhengsoft.Samples5.controller;

import com.zhuozhengsoft.pageoffice.FileSaver;
import com.zhuozhengsoft.pageoffice.OpenModeType;
import com.zhuozhengsoft.pageoffice.PageOfficeCtrl;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.servlet.ModelAndView;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.Statement;
import java.util.Map;

@Controller
@RequestMapping("SaveAndSearch")
public class SaveAndSearchController {

    @RequestMapping("/FileManage")
    public ModelAndView FileManage(HttpServletRequest request, Map<String, Object> map) throws Exception {
        request.setCharacterEncoding("utf-8");
        String key = request.getParameter("Input_KeyWord");
        String sql = "";
        if (key != null && key.trim().length() > 0) {
            request.setAttribute("key", key);
            sql = "select * from word  where Content like '%" + key
                    + "%' order by ID desc";
            System.out.println("--key--" + key);
        } else {
            sql = "select * from word order by ID desc";
        }
        Class.forName("org.sqlite.JDBC");
        String strUrl = "jdbc:sqlite:"
                + request.getServletContext().getRealPath("demodata/") + "\\SaveAndSearch.db";
        Connection conn = DriverManager.getConnection(strUrl);
        Statement stmt = conn.createStatement();
        ResultSet rs = stmt.executeQuery(sql);
        int id;
        StringBuilder stringBuilder = new StringBuilder();
        String FileName = "";
        String Content = "";
        while (rs.next()) {
            FileName = rs.getString("FileName");
            Content = rs.getString("Content");
            id = rs.getInt("ID");
            stringBuilder.append("<tr onmouseover='onColor(this)' onmouseout='offColor(this)'>\n" +
                    "                <td>").append(FileName).append("</td>");
            stringBuilder.append("<td style='text-align:center;'>");
            if (key != null && key.trim().length() > 0) {
                stringBuilder.append("<a style=\"color:#00217d;\"\n" +
                        "                href=\"javascript:POBrowser.openWindowModeless('Edit?id=" + id + "&key='+encodeURI('" + key + "')+'','width=1200px;height=800px;')\">编辑</a>");
            }
            if (key == null || key == "") {
                stringBuilder.append("<a style=\"color:#00217d;\"\n" +
                        "                href=\"javascript:POBrowser.openWindowModeless('Edit?id=" + id + "' ,'width=1200px;height=800px;');\">编辑</a>");
            }
            stringBuilder.append("</td>");
        }
        stringBuilder.append("</tr>");

        stmt.close();
        conn.close();
        request.setAttribute("list",stringBuilder);
        ModelAndView mv = new ModelAndView("SaveAndSearch/FileManage");
        return mv;
    }

    @RequestMapping("/Edit")
    public ModelAndView Edit(HttpServletRequest request, Map<String, Object> map) throws Exception {
        String key = request.getParameter("key");
        int id = Integer.parseInt(request.getParameter("id"));
        //根据id查询数据库中对应的文档名称
        Class.forName("org.sqlite.JDBC");
        String strUrl = "jdbc:sqlite:"
                + request.getServletContext().getRealPath("demodata/") + "/SaveAndSearch.db";
        Connection conn = DriverManager.getConnection(strUrl);
        Statement stmt = conn.createStatement();
        String sql = "select * from word where id=" + id;
        ResultSet rs = stmt.executeQuery(sql);
        String FileName = "";
        while (rs.next()) {
            FileName = rs.getString("FileName");
        }
        stmt.close();
        conn.close();

        /******************************PageOffice编程开始**************************/
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage(request.getContextPath() + "/poserver.zz"); //此行必须
        //隐藏菜单栏
        poCtrl.setMenubar(false);
        poCtrl.addCustomToolButton("保存", "Save()", 1);
        //设置保存页面
        poCtrl.setSaveFilePage("save?id=" + id);
        //打开Word文件
        String filePath = request.getSession().getServletContext().getRealPath("/doc/SaveAndSearch/" + FileName + ".doc");
        poCtrl.webOpen(filePath, OpenModeType.docNormalEdit, "张三");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        map.put("key", key);
        ModelAndView mv = new ModelAndView("SaveAndSearch/Edit");
        return mv;
    }

    @RequestMapping("/save")
    public void save(HttpServletRequest request, HttpServletResponse response) throws Exception {
        FileSaver fs = new FileSaver(request, response);
        fs.saveToFile(request.getSession().getServletContext().getRealPath("doc/SaveAndSearch/") + "/" + fs.getFileName());
        fs.setCustomSaveResult("ok");
        String strDocumentText = fs.getDocumentText();
        //更新数据库中文档的文本内容
        int id = Integer.parseInt(request.getParameter("id"));
        Class.forName("org.sqlite.JDBC");
        String strUrl = "jdbc:sqlite:"
                + request.getServletContext().getRealPath("demodata/") + "\\SaveAndSearch.db";
        Connection conn = DriverManager.getConnection(strUrl);
        Statement stmt = conn.createStatement();
        String strsql = "update word set Content='" + strDocumentText + "' where id=" + id;
        stmt.executeUpdate(strsql);
        stmt.close();
        conn.close();
        fs.close();

    }
}