package com.zhuozhengsoft.Samples5.controller;
import com.zhuozhengsoft.pageoffice.OpenModeType;
import com.zhuozhengsoft.pageoffice.PageOfficeCtrl;
import com.zhuozhengsoft.pageoffice.excelwriter.Sheet;
import com.zhuozhengsoft.pageoffice.excelwriter.Workbook;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.servlet.ModelAndView;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.util.Map;

@Controller
@RequestMapping("DefinedNameCell")
public class DefinedNameCellController {

    @RequestMapping("/ExcelFill")
    public ModelAndView openExcel(HttpServletRequest request, Map<String, Object> map) throws Exception{
        //设置PageOfficeCtrl控件的服务页面
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage(request.getContextPath() + "/poserver.zz"); //此行必须
        poCtrl.setCaption("简单的给Excel赋值");
        //定义Workbook对象
        Workbook workBook = new Workbook();
        //定义Sheet对象，"Sheet1"是打开的Excel表单的名称
        Sheet sheet = workBook.openSheet("Sheet1");
        sheet.openCellByDefinedName("testA1").setValue("Tom");
        sheet.openCellByDefinedName("testB1").setValue("John");

        poCtrl.setWriter(workBook);
        //隐藏菜单栏
        poCtrl.setMenubar(false);
        poCtrl.setSaveDataPage("SaveData");
        poCtrl.addCustomToolButton("保存", "Save()", 1);
        //打开Word文档
        poCtrl.webOpen("/doc/DefinedNameCell/test.xls", OpenModeType.xlsSubmitForm, "李斯");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        ModelAndView mv = new ModelAndView("DefinedNameCell/ExcelFill");
        return mv;
    }

    @RequestMapping("/SaveData")
    public void SaveData(HttpServletRequest request, HttpServletResponse response) throws Exception{
        com.zhuozhengsoft.pageoffice.excelreader.Workbook workBook = new com.zhuozhengsoft.pageoffice.excelreader.Workbook(request, response);
        com.zhuozhengsoft.pageoffice.excelreader.Sheet sheet = workBook.openSheet("Sheet1");
        String content = "";
        content += "testA1：" + sheet.openCellByDefinedName("testA1").getValue() + "<br/>";
        content += "testB1：" + sheet.openCellByDefinedName("testB1").getValue() + "<br/>";
        response.setContentType("text/plain; charset=utf-8");
        response.getWriter().print(content);
        workBook.showPage(500, 400);
        workBook.close();
    }
}