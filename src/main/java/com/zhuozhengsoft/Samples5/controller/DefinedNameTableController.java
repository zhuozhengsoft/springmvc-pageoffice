package com.zhuozhengsoft.Samples5.controller;

import com.zhuozhengsoft.pageoffice.OpenModeType;
import com.zhuozhengsoft.pageoffice.PageOfficeCtrl;
import com.zhuozhengsoft.pageoffice.excelwriter.Cell;
import com.zhuozhengsoft.pageoffice.excelwriter.Sheet;
import com.zhuozhengsoft.pageoffice.excelwriter.Table;
import com.zhuozhengsoft.pageoffice.excelwriter.Workbook;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.servlet.ModelAndView;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.text.DecimalFormat;
import java.text.NumberFormat;
import java.util.Calendar;
import java.util.GregorianCalendar;
import java.util.Map;

@Controller
@RequestMapping("DefinedNameTable")
public class DefinedNameTableController {
    @RequestMapping("/Default")
    public ModelAndView Default(HttpServletRequest request, Map<String, Object> map) throws Exception{
        ModelAndView mv = new ModelAndView("DefinedNameTable/Default");
        return mv;
    }

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
        Table table = sheet.openTableByDefinedName("report", 10, 5, false);
        table.getDataFields().get(0).setValue("轮胎");
        table.getDataFields().get(1).setValue("100");
        table.getDataFields().get(2).setValue("120");
        table.getDataFields().get(3).setValue("500");
        table.getDataFields().get(4).setValue("120%");
        table.nextRow();
        table.close();

        poCtrl.setWriter(workBook);
        //隐藏菜单栏
        poCtrl.setMenubar(false);
        poCtrl.setSaveDataPage("SaveData");
        poCtrl.addCustomToolButton("保存", "Save()", 1);
        //打开Word文档
        poCtrl.webOpen("/doc/DefinedNameTable/test.xls", OpenModeType.xlsSubmitForm, "李斯");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        ModelAndView mv = new ModelAndView("DefinedNameTable/ExcelFill");
        return mv;
    }

    @RequestMapping("/ExcelFill2")
    public ModelAndView openExcel2(HttpServletRequest request, Map<String, Object> map) throws Exception{
        String tempFileName = request.getParameter("temp");

        //设置PageOfficeCtrl控件的服务页面
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage(request.getContextPath() + "/poserver.zz"); //此行必须
        poCtrl.setCaption("简单的给Excel赋值");

        //定义Workbook对象
        Workbook workBook = new Workbook();
        //定义Sheet对象，"Sheet1"是打开的Excel表单的名称
        Sheet sheet = workBook.openSheet("Sheet1");
        //定义Table对象，参数“report”就是Excel模板中定义的单元格区域的名称
        Table table = sheet.openTableByDefinedName("report", 10, 5, true);
        //给区域中的单元格赋值
        table.getDataFields().get(0).setValue("轮胎");
        table.getDataFields().get(1).setValue("100");
        table.getDataFields().get(2).setValue("120");
        table.getDataFields().get(3).setValue("500");
        table.getDataFields().get(4).setValue("120%");
        //循环下一行
        table.nextRow();
        //关闭table对象
        table.close();
        //定义单元格对象，参数“year”就是Excel模板中定义的单元格的名称
        Cell cellYear = sheet.openCellByDefinedName("year");
        // 给单元格赋值
        Calendar c = new GregorianCalendar();
        int year = c.get(Calendar.YEAR);//获取年份
        cellYear.setValue(year + "年");
        Cell cellName = sheet.openCellByDefinedName("name");
        cellName.setValue("张三");

        poCtrl.setWriter(workBook);
        //隐藏菜单栏
        poCtrl.setMenubar(false);
        //打开Word文档
        poCtrl.webOpen("/doc/DefinedNameTable/"+tempFileName, OpenModeType.xlsSubmitForm, "李斯");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        ModelAndView mv = new ModelAndView("DefinedNameTable/ExcelFill2");
        return mv;
    }
    @RequestMapping("/ExcelFill4")
    public ModelAndView openExcel3(HttpServletRequest request, Map<String, Object> map) throws Exception{
        //设置PageOfficeCtrl控件的服务页面
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage(request.getContextPath() + "/poserver.zz"); //此行必须
        poCtrl.setCaption("简单的给Excel赋值");

        //定义Workbook对象
        Workbook workBook = new Workbook();
        //定义Sheet对象，"Sheet1"是打开的Excel表单的名称
        Sheet sheet = workBook.openSheet("Sheet1");
        //定义Table对象
        Table table = sheet.openTable("B4:F11");
        int rowCount = 12;//假设将要自动填充数据的实际记录条数为12
        for (int i = 1; i <= rowCount; i++) {
            //给区域中的单元格赋值
            table.getDataFields().get(0).setValue(i + "月");
            table.getDataFields().get(1).setValue("100");
            table.getDataFields().get(2).setValue("120");
            table.getDataFields().get(3).setValue("500");
            table.getDataFields().get(4).setValue("120%");
            table.nextRow();//循环下一行，此行必须
        }
        //关闭table对象
        table.close();
        //定义Table对象
        Table table2 = sheet.openTable("B13:F16");
        int rowCount2 = 4;//假设将要自动填充数据的实际记录条数为12
        for (int i = 1; i <= rowCount2; i++) {
            //给区域中的单元格赋值
            table2.getDataFields().get(0).setValue(i + "季度");
            table2.getDataFields().get(1).setValue("300");
            table2.getDataFields().get(2).setValue("300");
            table2.getDataFields().get(3).setValue("300");
            table2.getDataFields().get(4).setValue("100%");
            table2.nextRow();
        }
        //关闭table对象
        table2.close();
        poCtrl.setWriter(workBook);
        //隐藏菜单栏
        poCtrl.setMenubar(false);
        //打开Word文档
        poCtrl.webOpen("/doc/DefinedNameTable/test4.xls", OpenModeType.xlsSubmitForm, "李斯");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));

        ModelAndView mv = new ModelAndView("DefinedNameTable/ExcelFill4");
        return mv;
    }
    @RequestMapping("/ExcelFill5")
    public ModelAndView openExcel5(HttpServletRequest request, Map<String, Object> map) throws Exception{
        //设置PageOfficeCtrl控件的服务页面
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage(request.getContextPath() + "/poserver.zz"); //此行必须
        poCtrl.setCaption("简单的给Excel赋值");
        //定义Workbook对象
        Workbook workBook = new Workbook();
        //定义Sheet对象，"Sheet1"是打开的Excel表单的名称
        Sheet sheet = workBook.openSheet("Sheet1");
        //定义Table对象，参数“report1”为Excel中定义的名称，“4”为名称指定区域的行数，
        //“5”为名称指定区域的列数，“true”表示表格会按实际数据行数自动扩展
        Table table = sheet.openTableByDefinedName("report", 4, 5, true);
        int rowCount = 12;//假设将要自动填充数据的实际记录条数为12
        for (int i = 1; i <= rowCount; i++) {
            //给区域中的单元格赋值
            table.getDataFields().get(0).setValue(i + "月");
            table.getDataFields().get(1).setValue("100");
            table.getDataFields().get(2).setValue("120");
            table.getDataFields().get(3).setValue("500");
            table.getDataFields().get(4).setValue("120%");
            table.nextRow();//循环下一行，此行必须
        }
        //关闭table对象
        table.close();
        //定义Table对象
        Table table2 = sheet.openTableByDefinedName("report2", 4, 5, true);
        int rowCount2 = 4;//假设将要自动填充数据的实际记录条数为12
        for (int i = 1; i <= rowCount2; i++) {
            //给区域中的单元格赋值
            table2.getDataFields().get(0).setValue(i + "季度");
            table2.getDataFields().get(1).setValue("300");
            table2.getDataFields().get(2).setValue("300");
            table2.getDataFields().get(3).setValue("300");
            table2.getDataFields().get(4).setValue("100%");
            table2.nextRow();
        }
        //关闭table对象
        table2.close();
        poCtrl.setWriter(workBook);
        //隐藏菜单栏
        poCtrl.setMenubar(false);
        //打开Word文档
        poCtrl.webOpen("/doc/DefinedNameTable/test4.xls", OpenModeType.xlsSubmitForm, "李斯");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        ModelAndView mv = new ModelAndView("DefinedNameTable/ExcelFill5");
        return mv;
    }
    @RequestMapping("/ExcelFill6")
    public ModelAndView openExcel6(HttpServletRequest request, Map<String, Object> map) throws Exception{
//设置PageOfficeCtrl控件的服务页面
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage(request.getContextPath() + "/poserver.zz"); //此行必须
        poCtrl.setCaption("简单的给Excel赋值");
        //隐藏菜单栏
        poCtrl.setMenubar(false);
        //打开Word文档
        poCtrl.webOpen("/doc/DefinedNameTable/test4.xls", OpenModeType.xlsSubmitForm, "李斯");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        ModelAndView mv = new ModelAndView("DefinedNameTable/ExcelFill6");
        return mv;
    }

    @RequestMapping("/SaveData")
    public void SaveData(HttpServletRequest request, HttpServletResponse response) throws Exception{
        com.zhuozhengsoft.pageoffice.excelreader.Workbook workBook = new com.zhuozhengsoft.pageoffice.excelreader.Workbook(request, response);
        com.zhuozhengsoft.pageoffice.excelreader.Sheet sheet = workBook.openSheet("Sheet1");
        com.zhuozhengsoft.pageoffice.excelreader.Table table = sheet.openTableByDefinedName("report");
        String content = "";
        int result = 0;
        while (!table.getEOF()) {
            //获取提交的数值
            if (!table.getDataFields().getIsEmpty()) {
                content += "<br/>月份名称："
                        + table.getDataFields().get(0).getText();
                content += "<br/>计划完成量："
                        + table.getDataFields().get(1).getText();
                content += "<br/>实际完成量："
                        + table.getDataFields().get(2).getText();
                content += "<br/>累计完成量："
                        + table.getDataFields().get(3).getText();
                //out.print(table.getDataFields().get(2).getText()+"      mmmmmmmmmmmmm          "+table.getDataFields().get(1).getText());

                int colCount = table.getDataFields().size();//获取列数

                if (table.getDataFields().get(2).getText().equals(null)
                        || table.getDataFields().get(2).getText().trim().length() == 0) {
                    content += "<br/>完成率：0%";
                } else {
                    float f = Float.parseFloat(table.getDataFields().get(2)
                            .getText());
                    f = f / Float.parseFloat(table.getDataFields().get(1).getText());
                    DecimalFormat df = (DecimalFormat) NumberFormat.getInstance();
                    content += "<br/>完成率：" + df.format(f * 100) + "%";
                }
                content += "<br/>*********************************************";
            }
            //循环进入下一行
            table.nextRow();
        }
        table.close();
        response.setContentType("text/plain; charset=utf-8");
        response.getWriter().print(content);
        workBook.showPage(500, 400);
        workBook.close();
    }
}