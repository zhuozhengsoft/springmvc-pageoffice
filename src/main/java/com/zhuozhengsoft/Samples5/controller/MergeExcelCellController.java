package com.zhuozhengsoft.Samples5.controller;

import com.zhuozhengsoft.pageoffice.FileSaver;
import com.zhuozhengsoft.pageoffice.OpenModeType;
import com.zhuozhengsoft.pageoffice.PageOfficeCtrl;
import com.zhuozhengsoft.pageoffice.excelwriter.*;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.servlet.ModelAndView;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.awt.*;
import java.util.Map;

@Controller
@RequestMapping("MergeExcelCell")
public class MergeExcelCellController {

    @RequestMapping("/Default")
    public ModelAndView openExcel(HttpServletRequest request, Map<String, Object> map) throws Exception{
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage(request.getContextPath() + "/poserver.zz");
        Workbook wb = new Workbook();
        Sheet sheet = wb.openSheet("Sheet1");
        //合并单元格
        sheet.openTable("B2:F2").merge();
        Cell cB2 = sheet.openCell("B2");
        cB2.setValue("北京某公司产品销售情况");
        //设置水平对齐方式
        cB2.setHorizontalAlignment(XlHAlign.xlHAlignCenter);
        cB2.setForeColor(Color.red);
        cB2.getFont().setSize(16);

        sheet.openTable("B4:B6").merge();
        Cell cB4 = sheet.openCell("B4");
        cB4.setValue("A产品");
        //设置水平对齐方式
        cB4.setHorizontalAlignment(XlHAlign.xlHAlignCenter);
        //设置垂直对齐方式
        cB4.setVerticalAlignment(XlVAlign.xlVAlignCenter);

        sheet.openTable("B7:B9").merge();
        Cell cB7 = sheet.openCell("B7");
        cB7.setValue("B产品");
        cB7.setHorizontalAlignment(XlHAlign.xlHAlignCenter);
        cB7.setVerticalAlignment(XlVAlign.xlVAlignCenter);

        poCtrl.setWriter(wb);
        //打开Word文档
        poCtrl.webOpen("/doc/MergeExcelCell/test.xls", OpenModeType.xlsNormalEdit, "李斯");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        ModelAndView mv = new ModelAndView("MergeExcelCell/Default");
        return mv;
    }

}