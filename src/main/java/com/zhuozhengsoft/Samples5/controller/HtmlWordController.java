package com.zhuozhengsoft.Samples5.controller;

import com.zhuozhengsoft.pageoffice.FileSaver;
import com.zhuozhengsoft.pageoffice.OpenModeType;
import com.zhuozhengsoft.pageoffice.PageOfficeCtrl;
import com.zhuozhengsoft.pageoffice.wordwriter.WordDocument;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.servlet.ModelAndView;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.util.Map;

@Controller
@RequestMapping("HtmlWord")
public class HtmlWordController {
    @RequestMapping("/index")
    public ModelAndView index(HttpServletRequest request, Map<String, Object> map) throws Exception{
        ModelAndView mv = new ModelAndView("HtmlWord/index");
        return mv;
    }

    @RequestMapping("/Word")
    public void openWord(HttpServletRequest request, HttpServletResponse response,Map<String, Object> map) throws Exception{
        //获取客户端传递的参数
        response.getWriter().println("param1:" + request.getParameter("param1"));
        response.getWriter().println("<br>");
        response.getWriter().println("param2:" + request.getParameter("param2"));

        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        //设置服务器页面
        poCtrl.setServerPage(request.getContextPath() + "/poserver.zz");
        poCtrl.setJsFunction_AfterDocumentOpened("AfterDocumentOpened");
        poCtrl.addCustomToolButton("保存", "Save", 1);
        //设置保存页面
        poCtrl.setSaveFilePage("save");//设置处理文件保存的请求方法
        //打开Word文档
        poCtrl.webOpen("/doc/HtmlWord/test.doc", OpenModeType.docNormalEdit, "张三");
        response.getWriter().println(poCtrl.getHtmlCode("PageOfficeCtrl1"));
    }
    @RequestMapping("/save")
    public void save(HttpServletRequest request, HttpServletResponse response) throws Exception{
        FileSaver fs = new FileSaver(request, response);
        fs.saveToFile(request.getSession().getServletContext().getRealPath("doc/HtmlWord/")+fs.getFileName());
        fs.close();
    }
}
