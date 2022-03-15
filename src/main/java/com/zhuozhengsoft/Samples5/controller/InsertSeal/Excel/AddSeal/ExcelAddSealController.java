package com.zhuozhengsoft.Samples5.controller.InsertSeal.Excel.AddSeal;

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
@RequestMapping("InsertSeal/Excel")
public class ExcelAddSealController {

    @RequestMapping("/AddSeal1/Excel1")
    public ModelAndView AddSeal1Excel1(HttpServletRequest request, Map<String, Object> map) throws Exception{
        //******************************卓正PageOffice组件的使用*******************************
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        //此行必须，设置PageOffice服务器地址
        poCtrl.setServerPage(request.getContextPath() + "/poserver.zz");
        //添加自定义按钮
        poCtrl.addCustomToolButton("保存", "Save", 1);
        poCtrl.addCustomToolButton("加盖印章", "InsertSeal()", 2);
        poCtrl.addCustomToolButton("删除印章", "DeleteSeal()", 21);
        poCtrl.addCustomToolButton("验证印章", "VerifySeal()", 5);
        poCtrl.addCustomToolButton("修改密码", "ChangePsw()", 0);
        //设置保存页面
        poCtrl.setSaveFilePage("save");
        //打开Word文档
        poCtrl.webOpen("/doc/InsertSeal/Excel/AddSeal1/test.xls", OpenModeType.xlsNormalEdit, "张三");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        ModelAndView mv = new ModelAndView("InsertSeal/Excel/AddSeal1/Excel1");
        return mv;
    }

    @RequestMapping("/AddSeal1/save")
    public void AddSeal1Word1Save(HttpServletRequest request, HttpServletResponse response) throws Exception{
        FileSaver fs = new FileSaver(request, response);
        fs.saveToFile(request.getSession().getServletContext().getRealPath("doc/InsertSeal/Excel/AddSeal1/")+fs.getFileName());
        fs.close();
    }

    @RequestMapping("/AddSeal2/Excel2")
    public ModelAndView AddSeal2Excel2(HttpServletRequest request, Map<String, Object> map) throws Exception{
        //******************************卓正PageOffice组件的使用*******************************
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        //此行必须，设置PageOffice服务器地址
        poCtrl.setServerPage(request.getContextPath() + "/poserver.zz");
        //添加自定义按钮
        poCtrl.addCustomToolButton("保存", "Save", 1);
        poCtrl.addCustomToolButton("加盖印章", "InsertSeal()", 2);
        poCtrl.addCustomToolButton("修改密码", "ChangePsw()", 0);
        //设置保存页面
        poCtrl.setSaveFilePage("save");
        //打开Word文档
        poCtrl.webOpen("/doc/InsertSeal/Excel/AddSeal2/test.xls", OpenModeType.xlsNormalEdit, "张三");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        ModelAndView mv = new ModelAndView("InsertSeal/Excel/AddSeal2/Excel2");
        return mv;

    }

    @RequestMapping("/AddSeal2/save")
    public void AddSeal2Excel2Save(HttpServletRequest request, HttpServletResponse response) throws Exception{
        FileSaver fs = new FileSaver(request, response);
        fs.saveToFile(request.getSession().getServletContext().getRealPath("doc/InsertSeal/Excel/AddSeal2/")+fs.getFileName());
        fs.close();
    }

    @RequestMapping("/AddSeal3/Excel3")
    public ModelAndView AddSeal3Excel3(HttpServletRequest request, Map<String, Object> map) throws Exception{
        //******************************卓正PageOffice组件的使用*******************************
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        //此行必须，设置PageOffice服务器地址
        poCtrl.setServerPage(request.getContextPath() + "/poserver.zz");
        //添加自定义按钮
        poCtrl.addCustomToolButton("保存", "Save", 1);
        poCtrl.addCustomToolButton("加盖印章", "InsertSeal()", 2);
        poCtrl.addCustomToolButton("删除印章", "DeleteSeal()", 21);
        poCtrl.addCustomToolButton("修改密码", "ChangePsw()", 0);
        //设置保存页面
        poCtrl.setSaveFilePage("save");
        //打开Word文档
        poCtrl.webOpen("/doc/InsertSeal/Excel/AddSeal3/test.xls", OpenModeType.xlsNormalEdit, "张三");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        ModelAndView mv = new ModelAndView("InsertSeal/Excel/AddSeal3/Excel3");
        return mv;
    }

    @RequestMapping("/AddSeal3/save")
    public void AddSeal3Word3Save(HttpServletRequest request, HttpServletResponse response) throws Exception{
        FileSaver fs = new FileSaver(request, response);
        fs.saveToFile(request.getSession().getServletContext().getRealPath("doc/InsertSeal/Word/AddSeal3/")+fs.getFileName());
        fs.close();
    }

    @RequestMapping("/AddSeal4/Excel4")
    public ModelAndView AddSeal4Excel4(HttpServletRequest request, Map<String, Object> map) throws Exception{
        //******************************卓正PageOffice组件的使用*******************************
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        //此行必须，设置PageOffice服务器地址
        poCtrl.setServerPage(request.getContextPath() + "/poserver.zz");
        //添加自定义按钮
        poCtrl.addCustomToolButton("保存", "Save", 1);
        poCtrl.addCustomToolButton("加盖印章", "InsertSeal()", 2);
        poCtrl.addCustomToolButton("验证文档", "VerifySeal()", 0);
        //设置保存页面
        poCtrl.setSaveFilePage("save");
        //打开Word文档
        poCtrl.webOpen("/doc/InsertSeal/Excel/AddSeal4/test.xls", OpenModeType.xlsNormalEdit, "张三");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        ModelAndView mv = new ModelAndView("InsertSeal/Excel/AddSeal4/Excel4");
        return mv;
    }

    @RequestMapping("/AddSeal4/save")
    public void AddSeal4Word4Save(HttpServletRequest request, HttpServletResponse response) throws Exception{
        FileSaver fs = new FileSaver(request, response);
        fs.saveToFile(request.getSession().getServletContext().getRealPath("doc/InsertSeal/Excel/AddSeal4/")+fs.getFileName());
        fs.close();
    }

    @RequestMapping("/AddSeal5/Excel5")
    public ModelAndView AddSeal5Excel5(HttpServletRequest request, Map<String, Object> map) throws Exception{
        //******************************卓正PageOffice组件的使用*******************************
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        //此行必须，设置PageOffice服务器地址
        poCtrl.setServerPage(request.getContextPath() + "/poserver.zz");
        //添加自定义按钮
        poCtrl.addCustomToolButton("保存", "Save", 1);
        poCtrl.addCustomToolButton("加盖印章", "InsertSeal()", 2);
        poCtrl.addCustomToolButton("删除指定印章", "DeleteSeal()", 21);
        poCtrl.addCustomToolButton("清除所有印章", "DeleteAllSeal()", 21);
        //设置保存页面
        poCtrl.setSaveFilePage("save");
        //打开Word文档
        poCtrl.webOpen("/doc/InsertSeal/Excel/AddSeal5/test.xls", OpenModeType.xlsNormalEdit, "张三");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        ModelAndView mv = new ModelAndView("InsertSeal/Excel/AddSeal5/Excel5");
        return mv;
    }

    @RequestMapping("/AddSeal5/save")
    public void AddSeal5Excel5Save(HttpServletRequest request, HttpServletResponse response) throws Exception{
        FileSaver fs = new FileSaver(request, response);
        fs.saveToFile(request.getSession().getServletContext().getRealPath("doc/InsertSeal/Excel/AddSeal5/")+fs.getFileName());
        fs.close();
    }


}
