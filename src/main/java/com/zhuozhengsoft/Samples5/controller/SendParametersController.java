package com.zhuozhengsoft.Samples5.controller;

import com.zhuozhengsoft.pageoffice.FileSaver;
import com.zhuozhengsoft.pageoffice.OpenModeType;
import com.zhuozhengsoft.pageoffice.PageOfficeCtrl;
import org.omg.PortableInterceptor.USER_EXCEPTION;
import org.springframework.context.annotation.ScannedGenericBeanDefinition;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.servlet.ModelAndView;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.net.URLDecoder;
import java.net.URLEncoder;
import java.sql.Connection;
import java.util.Map;

import java.sql.Statement;
import java.sql.Connection;
import java.sql.DriverManager;

@Controller
@RequestMapping("SendParameters")
public class SendParametersController {

    @RequestMapping("/SendParameters")
    public ModelAndView openWord(HttpServletRequest request, Map<String, Object> map) throws Exception{
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        //设置服务器页面
        poCtrl.setServerPage(request.getContextPath() + "/poserver.zz");
        //添加自定义按钮
        poCtrl.addCustomToolButton("保存", "Save()", 1);
        poCtrl.addCustomToolButton("全屏", "SetFullScreen()", 4);
        poCtrl.addCustomToolButton("关闭", "Close()", 21);
        //设置保存页面
        //设置保存页面
        poCtrl.setSaveFilePage("save?id=1");//设置处理文件保存的请求方法
        //打开word
        poCtrl.webOpen("/doc/SendParameters/test.doc", OpenModeType.docNormalEdit, "李斯");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        ModelAndView mv = new ModelAndView("SendParameters/SendParameters");
        return mv;
    }

    @RequestMapping("/save")
    public void save(HttpServletRequest request, HttpServletResponse response) throws Exception{
        FileSaver fs = new FileSaver(request, response);
        fs.saveToFile(request.getSession().getServletContext().getRealPath("doc/SendParameters/")+fs.getFileName());

        int id = 0;
        String userName = "";
        int age = 0;
        String sex = "";
        String content = "";

        //获取通过Url传递过来的值
        if (request.getParameter("id") != null
                && request.getParameter("id").trim().length() > 0)
            id = Integer.parseInt(request.getParameter("id").trim());

        //获取通过网页标签控件传递过来的参数值，注意fs.getFormField("参数名")方法中的参数名是值标签的“name”属性而不是Id

        //获取通过文本框<input type="text" />标签传递过来的值
        if (fs.getFormField("userName") != null
                && fs.getFormField("userName").trim().length() > 0) {
            userName = fs.getFormField("userName");
        }

        //获取通过隐藏域传递过来的值
        if (fs.getFormField("age") != null
                && fs.getFormField("age").trim().length() > 0) {
            age = Integer.parseInt(fs.getFormField("age"));
        }

        //获取通过<select>标签传递过来的值
        if (fs.getFormField("selSex") != null
                && fs.getFormField("selSex").trim().length() > 0) {
            sex = fs.getFormField("selSex");
        }

        Class.forName("org.sqlite.JDBC");//载入驱动程序类别
        String strUrl = "jdbc:sqlite:"
                + request.getServletContext().getRealPath("demodata/") + "\\SendParameters.db";
        Connection conn = DriverManager.getConnection(strUrl);
        Statement stmt = conn.createStatement();
        String strsql = "Update Users set UserName = '" + userName
                + "', age = " + age + ",sex = '" + sex + "' where id = "
                + id;
        stmt.executeUpdate(strsql);
        conn.close();

        content += "id：" + id +"</br>";
        content += "userName：" + userName +"</br>";
        content += "age：" + age +"</br>";
        content += "sex：" + sex +"</br>";
        response.setContentType("text/plain; charset=utf-8");
        response.getWriter().write(content);
        fs.showPage(300, 200);
        fs.close();
    }
}