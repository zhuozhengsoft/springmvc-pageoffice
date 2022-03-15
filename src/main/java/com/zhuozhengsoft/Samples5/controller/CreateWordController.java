package com.zhuozhengsoft.Samples5.controller;

import com.zhuozhengsoft.pageoffice.DocumentVersion;
import com.zhuozhengsoft.pageoffice.FileSaver;
import com.zhuozhengsoft.pageoffice.OpenModeType;
import com.zhuozhengsoft.pageoffice.PageOfficeCtrl;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.servlet.ModelAndView;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.Statement;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Map;

@Controller
@RequestMapping("CreateWord")
public class CreateWordController {
    @RequestMapping("/word-lists")
    public ModelAndView wordList(HttpServletRequest request, Map<String, Object> map) throws Exception {
        StringBuilder wordlist = new StringBuilder() ;
        Class.forName("org.sqlite.JDBC");
        String strUrl = "jdbc:sqlite:"
                + request.getServletContext().getRealPath("demodata/") + "\\CreateWord.db";
        Connection conn = DriverManager.getConnection(strUrl);
        Statement stmt = conn.createStatement();
        ResultSet rs = stmt.executeQuery("select * from word order by id desc");
        String fileName = "";
        String subject = "";
        String submitTime = "";
        while (rs.next()) {
            int id = rs.getInt("ID");
            fileName = rs.getString("FileName");
            subject = rs.getString("Subject");
            submitTime = rs.getString("SubmitTime");
            wordlist.append(" <tr onmouseover='onColor(this)' onmouseout='offColor(this)'> <td>" +
                    "<a href=\"javascript:POBrowser.openWindowModeless('OpenWord?id="+id+"' , 'width=1200px;height=800px;')\">" + subject +
                    "</a>" +
                    "</td>");

            if (submitTime != null && submitTime.length() > 0) {
                wordlist.append("<td>" + new SimpleDateFormat("yyyy/MM/dd").format(new SimpleDateFormat("yyyy-MM-dd").parse(submitTime)) + "</td>");

            } else {
                wordlist.append("<td>&nbsp;</td>");

            }
            wordlist.append(" </tr>");
        }
        stmt.close();
        conn.close();
        request.setAttribute("wordlist",wordlist);
        ModelAndView mv = new ModelAndView("CreateWord/word-lists");
        return mv;
    }

    @RequestMapping("/OpenWord")
    public ModelAndView openWord(HttpServletRequest request, Map<String, Object> map) throws Exception {
        String subject = "";
        String fileName = "";

        Class.forName("org.sqlite.JDBC");
        String strUrl = "jdbc:sqlite:" + request.getServletContext().getRealPath("demodata/") + "\\CreateWord.db";
        Connection conn = DriverManager.getConnection(strUrl);
        Statement stmt = conn.createStatement();
        String id = request.getParameter("id");
        if (!id.equals("") && !id.equals(null)) {
            ResultSet rs = stmt.executeQuery("select * from word where ID=" + id);
            subject = rs.getString("Subject");
            fileName = rs.getString("FileName");
            System.out.println(fileName);
            rs.close();
        }
        stmt.close();
        conn.close();

        //******************************卓正PageOffice组件的使用*******************************
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage(request.getContextPath() + "/poserver.zz"); //此行必须
        //隐藏菜单栏
        poCtrl.setMenubar(false);
        poCtrl.addCustomToolButton("保存", "Save()", 1);
        //设置保存页面
        poCtrl.setSaveFilePage("save");//设置处理文件保存的请求方法
        //打开word
        poCtrl.webOpen("/doc/CreateWord/"+fileName, OpenModeType.docNormalEdit, "李斯");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        map.put("subject", subject);
        map.put("fileName", fileName);
        ModelAndView mv = new ModelAndView("CreateWord/OpenWord");
        return mv;
    }

    @RequestMapping("/CopyCreatWord")
    public void CopyCreatWord(HttpServletRequest request, HttpServletResponse response) throws Exception {
        if (request.getParameter("op") != null && request.getParameter("op").length() > 0) {
            Class.forName("org.sqlite.JDBC");
            String strUrl = "jdbc:sqlite:"
                    + request.getServletContext().getRealPath("demodata/") + "/CreateWord.db";
            Connection conn = DriverManager.getConnection(strUrl);
            Statement stmt = conn.createStatement();
            ResultSet rs = stmt.executeQuery("select Max(ID) from word");
            int newID = 1;
            if (rs.next()) {
                newID = Integer.parseInt(rs.getString(1)) + 1;
            }
            rs.close();

            String fileName = "aabb" + newID + ".doc";

            String FileSubject = "请输入文档主题";
            String getFile = (String) request.getParameter("FileSubject");
            if (getFile != null && getFile.length() > 0)
                FileSubject = new String(getFile.getBytes("iso-8859-1"), "utf-8");
            //out.print(FileSubject);
            SimpleDateFormat df = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");//设置日期格式
            // new Date()为获取当前系统时间
            String strsql = "Insert into word(ID,FileName,Subject,SubmitTime) values(" + newID
                    + ",'" + fileName + "','" + FileSubject + "','" + df.format(new Date()) + "')";
            stmt.executeUpdate(strsql);
            stmt.close();
            conn.close();

            //拷贝文件
            //if(request.getParameter("action").equals("create")){
            String oldPath = request.getServletContext().getRealPath("doc/CreateWord/template.doc");
            String newPath = request.getServletContext().getRealPath("doc/CreateWord/" + fileName);
            try {
                int bytesum = 0;
                int byteread = 0;
                File oldfile = new File(oldPath);
                if (oldfile.exists()) { //文件存在时
                    InputStream inStream = new FileInputStream(oldPath); //读入原文件
                    FileOutputStream fs = new FileOutputStream(newPath);
                    byte[] buffer = new byte[1444];
                    int length;
                    while ((byteread = inStream.read(buffer)) != -1) {
                        bytesum += byteread; //字节数 文件大小
                        System.out.println(bytesum);
                        fs.write(buffer, 0, byteread);
                    }
                    inStream.close();
                }
            } catch (Exception e) {
                System.out.println("复制单个文件操作出错");
                e.printStackTrace();
            }
            //}
            response.sendRedirect("word-lists");
        }
    }


    @RequestMapping("/CreateWord")
    public ModelAndView CreatWord(HttpServletRequest request, Map<String, Object> map) throws Exception {
        //******************************卓正PageOffice组件的使用*******************************
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage(request.getContextPath() + "/poserver.zz"); //此行必须
        //隐藏菜单栏
        poCtrl.setMenubar(false);
        //隐藏工具栏
        poCtrl.setCustomToolbar(false);
        poCtrl.setJsFunction_BeforeDocumentSaved("BeforeDocumentSaved()");
        //设置保存页面
        poCtrl.setSaveFilePage("saveNewFile");
        //新建Word文件，webCreateNew方法中的两个参数分别指代“操作人”和“新建Word文档的版本号”
        poCtrl.webCreateNew("张佚名", DocumentVersion.Word2003);
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        ModelAndView mv = new ModelAndView("CreateWord/CreateWord");
        return mv;
    }

    @RequestMapping("/save")
    public void save(HttpServletRequest request, HttpServletResponse response) throws Exception {
        FileSaver fs = new FileSaver(request, response);
        fs.saveToFile(request.getSession().getServletContext().getRealPath("doc/CreateWord/") + fs.getFileName());
        fs.close();
    }

    @RequestMapping("/saveNewFile")
    public void SaveNewFile(HttpServletRequest request, HttpServletResponse response) throws Exception {
        //定义保存对象
        FileSaver fs = new FileSaver(request, response);

        Class.forName("org.sqlite.JDBC");
        String strUrl = "jdbc:sqlite:"
                + request.getServletContext().getRealPath("demodata/") + "\\CreateWord.db";
        Connection conn = DriverManager.getConnection(strUrl);
        Statement stmt = conn.createStatement();
        ResultSet rs = stmt.executeQuery("select Max(ID) from word");
        int newID = 1;
        if (rs.next()) {
            newID = Integer.parseInt(rs.getString(1)) + 1;
        }
        rs.close();

        String FileSubject = fs.getFormField("FileSubject").trim();
        String fileName = "aabb" + newID + ".doc";
        String getFile = (String) request.getParameter("FileSubject");
        if (getFile != null && getFile.length() > 0)
            FileSubject = new String(getFile.getBytes("iso-8859-1"));
        //out.print(FileSubject);
        SimpleDateFormat df = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");//设置日期格式
        // new Date()为获取当前系统时间
        String strsql = "Insert into word(ID,FileName,Subject,SubmitTime) values("
                + newID
                + ",'"
                + fileName
                + "','"
                + FileSubject
                + "','"
                + df.format(new Date()) + "')";
        stmt.executeUpdate(strsql);
        stmt.close();
        conn.close();

        //保存文件
        fs.saveToFile(request.getSession().getServletContext().getRealPath("doc/CreateWord/") + "/" + fileName);
        //设置保存结果
        fs.setCustomSaveResult("ok");
        //fs.showPage(300,300);
        fs.close();
    }
}