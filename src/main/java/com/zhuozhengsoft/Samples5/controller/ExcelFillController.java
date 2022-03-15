package com.zhuozhengsoft.Samples5.controller;

import com.zhuozhengsoft.pageoffice.FileSaver;
import com.zhuozhengsoft.pageoffice.OpenModeType;
import com.zhuozhengsoft.pageoffice.PageOfficeCtrl;
import com.zhuozhengsoft.pageoffice.excelwriter.Cell;
import com.zhuozhengsoft.pageoffice.excelwriter.Sheet;
import com.zhuozhengsoft.pageoffice.excelwriter.Workbook;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.servlet.ModelAndView;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.text.DecimalFormat;
import java.text.NumberFormat;
import java.util.Map;

@Controller
@RequestMapping("ExcelFill")
public class ExcelFillController {

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
        //定义Cell对象
        Cell cellB4 = sheet.openCell("B4");
        //给单元格赋值
        cellB4.setValue("1月");
        Cell cellC4 = sheet.openCell("C4");
        cellC4.setValue("300");
        Cell cellD4 = sheet.openCell("D4");
        cellD4.setValue("270");
        Cell cellE4 = sheet.openCell("E4");
        cellE4.setValue("270");
        Cell cellF4 = sheet.openCell("F4");
        DecimalFormat df = (DecimalFormat) NumberFormat.getInstance();
        cellF4.setValue(df.format(270.00 / 300 * 100) + "%");

        poCtrl.setWriter(workBook);
        //隐藏菜单栏
        poCtrl.setMenubar(false);
        //隐藏工具栏
        poCtrl.setCustomToolbar(false);
        //打开Word文档
        poCtrl.webOpen("/doc/ExcelFill/test.xls", OpenModeType.xlsNormalEdit, "李斯");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        ModelAndView mv = new ModelAndView("ExcelFill/ExcelFill");
        return mv;
    }
}