package com.zhuozhengsoft.Samples5.controller;

import com.zhuozhengsoft.pageoffice.FileSaver;
import com.zhuozhengsoft.pageoffice.OpenModeType;
import com.zhuozhengsoft.pageoffice.PageOfficeCtrl;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.servlet.ModelAndView;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.util.Map;

@Controller
@RequestMapping("POBrowserTopic")
public class POBrowserTopicController {

    @RequestMapping("/index")
    public ModelAndView openWord(HttpServletRequest request) {
        String userName = "张三";
        request.getSession().setAttribute("userName", userName);
        ModelAndView mv = new ModelAndView("POBrowserTopic/index");
        return mv;
    }

    @RequestMapping("/Word1")
    public ModelAndView openWord(HttpServletRequest request, Map<String, Object> map) throws Exception{
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        //设置服务器页面
        poCtrl.setServerPage(request.getContextPath() + "/poserver.zz");
        //添加自定义按钮
        poCtrl.addCustomToolButton("保存", "Save", 1);
        poCtrl.setJsFunction_AfterDocumentOpened("AfterDocumentOpened()");
        //设置保存页面
        poCtrl.setSaveFilePage("save");//设置处理文件保存的请求方法
        //打开word
        poCtrl.webOpen("/doc/POBrowserTopic/test.doc", OpenModeType.docNormalEdit, "李斯");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        ModelAndView mv = new ModelAndView("POBrowserTopic/Word1");
        return mv;
    }
    @RequestMapping("/Word2")
    public ModelAndView openWord2(HttpServletRequest request, Map<String, Object> map) throws Exception{
        //获取index.jsp页面传递过来参数的值
        String userName = (String) request.getSession().getAttribute("userName");
        //获取index.jsp用？传递过来的id的值
        String id = request.getParameter("id");
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        //设置服务器页面
        poCtrl.setServerPage(request.getContextPath() + "/poserver.zz");
        //添加自定义按钮
        poCtrl.addCustomToolButton("保存", "Save", 1);
        //设置保存页面
        poCtrl.setSaveFilePage("save");//设置处理文件保存的请求方法
        //打开word
        poCtrl.webOpen("/doc/POBrowserTopic/test.doc", OpenModeType.docNormalEdit, "李斯");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        map.put("id", id);
        map.put("userName", userName);
        ModelAndView mv = new ModelAndView("POBrowserTopic/Word2");
        return mv;
    }
    @RequestMapping("/Word3")
    public ModelAndView openWord3(HttpServletRequest request, Map<String, Object> map) throws Exception{
        String txt = (String) request.getSession().getAttribute("txt");

        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        //设置服务器页面
        poCtrl.setServerPage(request.getContextPath() + "/poserver.zz");
        //添加自定义按钮
        poCtrl.addCustomToolButton("保存并关闭", "Save", 1);
        //设置服务器页面
        poCtrl.setServerPage(request.getContextPath() + "/poserver.zz");
        //添加自定义按钮
        poCtrl.addCustomToolButton("保存", "Save", 1);
        //设置保存页面
        poCtrl.setSaveFilePage("save");//设置处理文件保存的请求方法
        //打开word
        poCtrl.webOpen("/doc/POBrowserTopic/test.doc", OpenModeType.docNormalEdit, "李斯");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        map.put("txt", txt);
        ModelAndView mv = new ModelAndView("POBrowserTopic/Word3");
        return mv;
    }
    @RequestMapping("/Result2")
    public void result2(HttpServletRequest request, HttpServletResponse response) throws Exception{
       request.getSession().setAttribute("txt",request.getParameter("param"));
    }
    @RequestMapping("/save")
    public void save(HttpServletRequest request, HttpServletResponse response) throws Exception{
        FileSaver fs = new FileSaver(request, response);
        fs.saveToFile(request.getSession().getServletContext().getRealPath("doc/POBrowserTopic/")+fs.getFileName());
        fs.close();
    }
}