package com.zhuozhengsoft.Samples5.controller.InsertSeal.Word.AddSign;

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
@RequestMapping("InsertSeal/Word")
public class WordAddSignController {

    @RequestMapping("/AddSign1/Word1")
    public ModelAndView AddSign1Word1(HttpServletRequest request, Map<String, Object> map) throws Exception{
        //******************************卓正PageOffice组件的使用*******************************
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage(request.getContextPath() + "/poserver.zz"); //此行必须
        //添加自定义按钮
        poCtrl.addCustomToolButton("保存", "Save", 1);
        poCtrl.addCustomToolButton("签字", "InsertHandSign()",3);
        poCtrl.addCustomToolButton("验证印章", "VerifySeal()", 5);
        poCtrl.addCustomToolButton("修改密码", "ChangePsw()", 0);
        //设置保存页面
        poCtrl.setSaveFilePage("save");
        //打开Word文档
        poCtrl.webOpen("/doc/InsertSeal/Word/AddSign1/word1.doc", OpenModeType.docNormalEdit, "张三");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        ModelAndView mv = new ModelAndView("InsertSeal/Word/AddSign1/Word1");
        return mv;
    }

    @RequestMapping("/AddSign1/save")
    public void AddSign1Word1Save(HttpServletRequest request, HttpServletResponse response) throws Exception{
        FileSaver fs = new FileSaver(request, response);
        fs.saveToFile(request.getSession().getServletContext().getRealPath("doc/InsertSeal/Word/AddSign1/")+fs.getFileName());
        fs.close();
    }

    @RequestMapping("/AddSign2/Word2")
    public ModelAndView AddSign2Word2(HttpServletRequest request, Map<String, Object> map) throws Exception{
        //******************************卓正PageOffice组件的使用*******************************
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        //此行必须，设置PageOffice服务器地址
        poCtrl.setServerPage(request.getContextPath() + "/poserver.zz");
        //添加自定义按钮
        poCtrl.addCustomToolButton("保存", "Save", 1);
        poCtrl.addCustomToolButton("签字", "InsertHandSign()", 2);
        poCtrl.addCustomToolButton("修改密码", "ChangePsw()", 0);
        //设置保存页面
        poCtrl.setSaveFilePage("save");
        //打开Word文档
        poCtrl.webOpen("/doc/InsertSeal/Word/AddSign2/test.doc", OpenModeType.docNormalEdit, "张三");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        ModelAndView mv = new ModelAndView("InsertSeal/Word/AddSign2/Word2");
        return mv;
    }

    @RequestMapping("/AddSign2/save")
    public void AddSeal2Word2Save(HttpServletRequest request, HttpServletResponse response) throws Exception{
        FileSaver fs = new FileSaver(request, response);
        fs.saveToFile(request.getSession().getServletContext().getRealPath("doc/InsertSeal/Word/AddSign2/")+fs.getFileName());
        fs.close();
    }

    @RequestMapping("/AddSign3/Word3")
    public ModelAndView AddSign3Word3(HttpServletRequest request, Map<String, Object> map) throws Exception{
        //******************************卓正PageOffice组件的使用*******************************
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage(request.getContextPath() + "/poserver.zz"); //此行必须
        //添加自定义按钮
        poCtrl.addCustomToolButton("保存", "Save", 1);
        poCtrl.addCustomToolButton("签字", "InsertHandSign()", 3);
        //设置保存页面
        poCtrl.setSaveFilePage("save");
        //打开Word文档
        poCtrl.webOpen("/doc/InsertSeal/Word/AddSign3/test.doc", OpenModeType.docNormalEdit, "张三");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        ModelAndView mv = new ModelAndView("InsertSeal/Word/AddSign3/Word3");
        return mv;
    }

    @RequestMapping("/AddSign3/save")
    public void AddSeal3Word3Save(HttpServletRequest request, HttpServletResponse response) throws Exception{
        FileSaver fs = new FileSaver(request, response);
        fs.saveToFile(request.getSession().getServletContext().getRealPath("doc/InsertSeal/Word/AddSign3/")+fs.getFileName());
        fs.close();
    }

    @RequestMapping("/AddSign4/Word4")
    public ModelAndView AddSign4Word4(HttpServletRequest request, Map<String, Object> map) throws Exception{
        //******************************卓正PageOffice组件的使用*******************************
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage(request.getContextPath() + "/poserver.zz"); //此行必须
        //添加自定义按钮
        poCtrl.addCustomToolButton("保存", "Save", 1);
        poCtrl.addCustomToolButton("签字", "AddHandSign()", 3);
        //设置保存页面
        poCtrl.setSaveFilePage("save");
        //打开Word文档
        poCtrl.webOpen("/doc/InsertSeal/Word/AddSign4/test.doc", OpenModeType.docNormalEdit, "张三");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        ModelAndView mv = new ModelAndView("InsertSeal/Word/AddSign4/Word4");
        return mv;
    }

    @RequestMapping("/AddSign4/save")
    public void AddSign4Word4Save(HttpServletRequest request, HttpServletResponse response) throws Exception{
        FileSaver fs = new FileSaver(request, response);
        fs.saveToFile(request.getSession().getServletContext().getRealPath("doc/InsertSeal/Word/AddSign4/")+fs.getFileName());
        fs.close();
    }
    @RequestMapping("/AddSign5/Word5")
    public ModelAndView AddSign5Word5(HttpServletRequest request, Map<String, Object> map) throws Exception{
        //******************************卓正PageOffice组件的使用*******************************
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage(request.getContextPath() + "/poserver.zz"); //此行必须
        //添加自定义按钮
        poCtrl.addCustomToolButton("保存", "Save", 1);
        poCtrl.addCustomToolButton("签字", "InsertHandSign()", 3);
        //设置保存页面
        poCtrl.setSaveFilePage("save");
        //打开Word文档
        poCtrl.webOpen("/doc/InsertSeal/Word/AddSign5/word2.doc", OpenModeType.docNormalEdit, "张三");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        ModelAndView mv = new ModelAndView("InsertSeal/Word/AddSign5/Word5");
        return mv;
    }

    @RequestMapping("/AddSign5/save")
    public void AddSign5Word4Save(HttpServletRequest request, HttpServletResponse response) throws Exception{
        FileSaver fs = new FileSaver(request, response);
        fs.saveToFile(request.getSession().getServletContext().getRealPath("doc/InsertSeal/Word/AddSign5/")+fs.getFileName());
        fs.close();
    }


}
