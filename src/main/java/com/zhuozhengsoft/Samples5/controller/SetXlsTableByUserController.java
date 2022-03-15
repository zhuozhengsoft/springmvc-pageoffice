package com.zhuozhengsoft.Samples5.controller;

import com.zhuozhengsoft.pageoffice.FileSaver;
import com.zhuozhengsoft.pageoffice.OpenModeType;
import com.zhuozhengsoft.pageoffice.PageOfficeCtrl;
import com.zhuozhengsoft.pageoffice.excelwriter.Sheet;
import com.zhuozhengsoft.pageoffice.excelwriter.Table;
import com.zhuozhengsoft.pageoffice.excelwriter.Workbook;

import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.servlet.ModelAndView;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.util.Map;

@Controller
@RequestMapping("SetXlsTableByUser")
public class SetXlsTableByUserController {
    @RequestMapping("/Default")
    public ModelAndView openWord(HttpServletRequest request) throws Exception {
        ModelAndView mv = new ModelAndView("SetXlsTableByUser/Default");
        return mv;
    }

    @RequestMapping("/SetXlsTableByUserName")
    public ModelAndView openWord(HttpServletRequest request, Map<String, Object> map) throws Exception {
        String userName = request.getParameter("userName");
        //***************************卓正PageOffice组件的使用********************************
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        Workbook wb = new Workbook();
        Sheet sheet = wb.openSheet("Sheet1");
        Table tableA = sheet.openTable("C4:D6");
        Table tableB = sheet.openTable("C7:D9");
        tableA.setSubmitName("tableA");
        tableB.setSubmitName("tableB");
        //根据登录用户名设置数据区域可编辑性
        String strInfo = "";

        //A部门经理登录后
        if (userName.equals("zhangsan")) {
            strInfo = "A部门经理，所以只能编辑A部门的产品数据";
            tableA.setReadOnly(false);
            tableB.setReadOnly(true);
        }
        //B部门经理登录后
        else {
            strInfo = "B部门经理，所以只能编辑B部门的产品数据";
            tableA.setReadOnly(true);
            tableB.setReadOnly(false);
        }
        poCtrl.setWriter(wb);
        //设置服务器页面
        poCtrl.setServerPage(request.getContextPath() + "/poserver.zz"); //此行必须
        //添加自定义按钮
        poCtrl.addCustomToolButton("保存", "Save", 1);
        //设置保存页
        poCtrl.setSaveFilePage("Save");
        poCtrl.setMenubar(false);
        //打开word
        poCtrl.webOpen("/doc/SetXlsTableByUser/test.xls", OpenModeType.xlsSubmitForm, userName);

        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        map.put("strInfo", strInfo);
        ModelAndView mv = new ModelAndView("SetXlsTableByUser/SetXlsTableByUserName");
        return mv;
    }

    @RequestMapping("/Save")
    public void Save(HttpServletRequest request, HttpServletResponse response) throws Exception {
        FileSaver fs = new FileSaver(request, response);
        fs.saveToFile(request.getSession().getServletContext().getRealPath("doc/SetXlsTableByUser/") + "/" + fs.getFileName());
        fs.close();
    }
}