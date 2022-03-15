package com.zhuozhengsoft.Samples5.controller.InsertSeal.Word.AddSeal;

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
@RequestMapping("InsertSeal")
public class WordAddSealController {

    @RequestMapping("/Word/AddSeal1/Word1")
    public ModelAndView AddSeal1Word1(HttpServletRequest request, Map<String, Object> map) throws Exception{
        //******************************卓正PageOffice组件的使用*******************************
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage(request.getContextPath() + "/poserver.zz"); //此行必须
        //添加自定义按钮
        poCtrl.addCustomToolButton("保存", "Save", 1);
        poCtrl.addCustomToolButton("加盖印章", "InsertSeal()", 2);
        poCtrl.addCustomToolButton("删除印章", "DeleteSeal()", 21);
        poCtrl.addCustomToolButton("验证印章", "VerifySeal()", 5);
        poCtrl.addCustomToolButton("修改密码", "ChangePsw()", 0);
        //设置保存页面
        poCtrl.setSaveFilePage("save");
        //打开Word文档
        poCtrl.webOpen("/doc/InsertSeal/Word/AddSeal1/word1.doc", OpenModeType.docNormalEdit, "张三");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        ModelAndView mv = new ModelAndView("InsertSeal/Word/AddSeal1/Word1");
        return mv;
    }

    @RequestMapping("/Word/AddSeal1/save")
    public void AddSeal1Word1Save(HttpServletRequest request, HttpServletResponse response) throws Exception{
        FileSaver fs = new FileSaver(request, response);
        fs.saveToFile(request.getSession().getServletContext().getRealPath("doc/InsertSeal/Word/AddSeal1/")+fs.getFileName());
        fs.close();
    }

    @RequestMapping("/Word/AddSeal2/Word2")
    public ModelAndView AddSeal2Word2(HttpServletRequest request, Map<String, Object> map) throws Exception{
        //******************************卓正PageOffice组件的使用*******************************
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage(request.getContextPath() + "/poserver.zz"); //此行必须
        //添加自定义按钮
        poCtrl.addCustomToolButton("保存", "Save", 1);
        poCtrl.addCustomToolButton("加盖印章", "InsertSeal()", 2);
        //设置保存页面
        poCtrl.setSaveFilePage("save");
        //打开Word文档
        poCtrl.webOpen("/doc/InsertSeal/Word/AddSeal2/word2.doc", OpenModeType.docNormalEdit, "张三");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        ModelAndView mv = new ModelAndView("InsertSeal/Word/AddSeal2/Word2");
        return mv;
    }

    @RequestMapping("/Word/AddSeal2/save")
    public void AddSeal2Word2Save(HttpServletRequest request, HttpServletResponse response) throws Exception{
        FileSaver fs = new FileSaver(request, response);
        fs.saveToFile(request.getSession().getServletContext().getRealPath("doc/InsertSeal/Word/AddSeal2/")+fs.getFileName());
        fs.close();
    }

    @RequestMapping("/Word/AddSeal3/Word3")
    public ModelAndView AddSeal3Word3(HttpServletRequest request, Map<String, Object> map) throws Exception{
        //******************************卓正PageOffice组件的使用*******************************
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage(request.getContextPath() + "/poserver.zz"); //此行必须
        //添加自定义按钮
        poCtrl.addCustomToolButton("保存", "Save", 1);
        poCtrl.addCustomToolButton("加盖印章", "InsertSeal()", 2);
        //设置保存页面
        poCtrl.setSaveFilePage("save");
        //打开Word文档
        poCtrl.webOpen("/doc/InsertSeal/Word/AddSeal3/word3.doc", OpenModeType.docNormalEdit, "张三");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        ModelAndView mv = new ModelAndView("InsertSeal/Word/AddSeal3/Word3");
        return mv;
    }

    @RequestMapping("/Word/AddSeal3/save")
    public void AddSeal3Word3Save(HttpServletRequest request, HttpServletResponse response) throws Exception{
        FileSaver fs = new FileSaver(request, response);
        fs.saveToFile(request.getSession().getServletContext().getRealPath("doc/InsertSeal/Word/AddSeal3/")+fs.getFileName());
        fs.close();
    }

    @RequestMapping("/Word/AddSeal4/Word4")
    public ModelAndView AddSeal4Word4(HttpServletRequest request, Map<String, Object> map) throws Exception{
        //******************************卓正PageOffice组件的使用*******************************
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage(request.getContextPath() + "/poserver.zz"); //此行必须
        //添加自定义按钮
        poCtrl.addCustomToolButton("保存", "Save", 1);
        poCtrl.addCustomToolButton("添加印章位置", "InsertSealPos()", 2);
        //设置保存页面
        poCtrl.setSaveFilePage("save");
        //打开Word文档
        poCtrl.webOpen("/doc/InsertSeal/Word/AddSeal4/word4.doc", OpenModeType.docNormalEdit, "张三");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        ModelAndView mv = new ModelAndView("InsertSeal/Word/AddSeal4/Word4");
        return mv;
    }

    @RequestMapping("/Word/AddSeal4/save")
    public void AddSeal4Word4Save(HttpServletRequest request, HttpServletResponse response) throws Exception{
        FileSaver fs = new FileSaver(request, response);
        fs.saveToFile(request.getSession().getServletContext().getRealPath("doc/InsertSeal/Word/AddSeal4/")+fs.getFileName());
        fs.close();
    }

    @RequestMapping("/Word/AddSeal5/Word5")
    public ModelAndView AddSeal5Word5(HttpServletRequest request, Map<String, Object> map) throws Exception{
        //******************************卓正PageOffice组件的使用*******************************
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage(request.getContextPath() + "/poserver.zz"); //此行必须
        //添加自定义按钮
        poCtrl.addCustomToolButton("保存", "Save", 1);
        poCtrl.addCustomToolButton("盖章到印章位置", "AddSealByPos()", 2);
        //设置保存页面
        poCtrl.setSaveFilePage("save");
        //打开Word文档
        poCtrl.webOpen("/doc/InsertSeal/Word/AddSeal5/test.doc", OpenModeType.docNormalEdit, "张三");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        ModelAndView mv = new ModelAndView("InsertSeal/Word/AddSeal5/Word5");
        return mv;
    }

    @RequestMapping("/Word/AddSeal5/save")
    public void AddSeal5Word5Save(HttpServletRequest request, HttpServletResponse response) throws Exception{
        FileSaver fs = new FileSaver(request, response);
        fs.saveToFile(request.getSession().getServletContext().getRealPath("doc/InsertSeal/Word/AddSeal5/")+fs.getFileName());
        fs.close();
    }

    @RequestMapping("/Word/AddSeal6/Word6")
    public ModelAndView AddSeal6Word6(HttpServletRequest request, Map<String, Object> map) throws Exception{
        //******************************卓正PageOffice组件的使用*******************************
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage(request.getContextPath() + "/poserver.zz"); //此行必须
        //添加自定义按钮
        poCtrl.addCustomToolButton("保存", "Save", 1);
        poCtrl.addCustomToolButton("盖章到印章位置", "AddSealByPos()", 2);
        //设置保存页面
        poCtrl.setSaveFilePage("save");
        //打开Word文档
        poCtrl.webOpen("/doc/InsertSeal/Word/AddSeal6/test.doc", OpenModeType.docNormalEdit, "张三");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        ModelAndView mv = new ModelAndView("InsertSeal/Word/AddSeal6/Word6");
        return mv;
    }

    @RequestMapping("/Word/AddSeal6/save")
    public void AddSeal6Word6Save(HttpServletRequest request, HttpServletResponse response) throws Exception{
        FileSaver fs = new FileSaver(request, response);
        fs.saveToFile(request.getSession().getServletContext().getRealPath("doc/InsertSeal/Word/AddSeal6/")+fs.getFileName());
        fs.close();
    }

    @RequestMapping("/Word/AddSeal7/Word7")
    public ModelAndView AddSeal7Word7(HttpServletRequest request, Map<String, Object> map) throws Exception{
        //******************************卓正PageOffice组件的使用*******************************
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage(request.getContextPath() + "/poserver.zz"); //此行必须
        //添加自定义按钮
        poCtrl.addCustomToolButton("保存", "Save", 1);
        poCtrl.addCustomToolButton("盖章到印章位置", "AddSealByPos()", 2);
        //设置保存页面
        poCtrl.setSaveFilePage("save");
        //打开Word文档
        poCtrl.webOpen("/doc/InsertSeal/Word/AddSeal7/test.doc", OpenModeType.docNormalEdit, "张三");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        ModelAndView mv = new ModelAndView("InsertSeal/Word/AddSeal7/Word7");
        return mv;
    }

    @RequestMapping("/Word/AddSeal7/save")
    public void AddSeal7Word7Save(HttpServletRequest request, HttpServletResponse response) throws Exception{
        FileSaver fs = new FileSaver(request, response);
        fs.saveToFile(request.getSession().getServletContext().getRealPath("doc/InsertSeal/Word/AddSeal7/")+fs.getFileName());
        fs.close();
    }

    @RequestMapping("/Word/AddSeal8/Word8")
    public ModelAndView AddSeal8Word8(HttpServletRequest request, Map<String, Object> map) throws Exception{
        //******************************卓正PageOffice组件的使用*******************************
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage(request.getContextPath() + "/poserver.zz"); //此行必须
        //添加自定义按钮
        poCtrl.addCustomToolButton("保存", "Save", 1);
        poCtrl.addCustomToolButton("加盖印章", "InsertSeal()", 2);
        poCtrl.addCustomToolButton("验证印章", "VerifySeal()", 5);
        poCtrl.addCustomToolButton("修改密码", "ChangePsw()", 0);
        //设置保存页面
        poCtrl.setSaveFilePage("save");
        //打开Word文档
        poCtrl.webOpen("/doc/InsertSeal/Word/AddSeal8/word8.doc", OpenModeType.docNormalEdit, "张三");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        ModelAndView mv = new ModelAndView("InsertSeal/Word/AddSeal8/Word8");
        return mv;
    }

    @RequestMapping("/Word/AddSeal8/save")
    public void AddSeal8Word8Save(HttpServletRequest request, HttpServletResponse response) throws Exception{
        FileSaver fs = new FileSaver(request, response);
        fs.saveToFile(request.getSession().getServletContext().getRealPath("doc/InsertSeal/Word/AddSeal8/")+fs.getFileName());
        fs.close();
    }

    @RequestMapping("/Word/AddSeal9/Word9")
    public ModelAndView AddSeal9Word9(HttpServletRequest request, Map<String, Object> map) throws Exception{
        //******************************卓正PageOffice组件的使用*******************************
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage(request.getContextPath() + "/poserver.zz"); //此行必须
        poCtrl.addCustomToolButton("保存", "Save", 1);
        poCtrl.addCustomToolButton("加盖印章", "InsertSeal()", 2);
        poCtrl.addCustomToolButton("删除指定印章", "DeleteSeal()", 21);
        poCtrl.addCustomToolButton("清除所有印章", "DeleteAllSeal()", 21);
        //设置保存页面
        poCtrl.setSaveFilePage("save");
        //打开Word文档
        poCtrl.webOpen("/doc/InsertSeal/Word/AddSeal9/word9.doc", OpenModeType.docNormalEdit, "张三");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        ModelAndView mv = new ModelAndView("InsertSeal/Word/AddSeal9/Word9");
        return mv;
    }

    @RequestMapping("/Word/AddSeal9/save")
    public void AddSeal9Word9Save(HttpServletRequest request, HttpServletResponse response) throws Exception{
        FileSaver fs = new FileSaver(request, response);
        fs.saveToFile(request.getSession().getServletContext().getRealPath("doc/InsertSeal/Word/AddSeal9/")+fs.getFileName());
        fs.close();
    }


}
