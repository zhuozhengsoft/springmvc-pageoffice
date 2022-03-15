package com.zhuozhengsoft.Samples5.controller;

import com.zhuozhengsoft.pageoffice.OpenModeType;
import com.zhuozhengsoft.pageoffice.PageOfficeCtrl;
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
import java.util.Map;

@Controller
@RequestMapping("ExcelCellClick")
public class ExcelCellClickController {

    @RequestMapping("/select")
    public ModelAndView select() throws Exception{
        ModelAndView mv = new ModelAndView("ExcelCellClick/select");
        return mv;
    }

    @RequestMapping("/SubmitExcel")
    public ModelAndView openExcel(HttpServletRequest request, Map<String, Object> map) throws Exception{
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        //设置服务器页面
        poCtrl.setServerPage(request.getContextPath() + "/poserver.zz");

        //定义Workbook对象
        Workbook workBook = new Workbook();
        //定义Sheet对象，"Sheet1"是打开的Excel表单的名称
        Sheet sheet = workBook.openSheet("Sheet1");
        //定义table对象，设置table对象的设置范围
        Table table = sheet.openTable("B4:D8");
        //设置table对象的提交名称，以便保存页面获取提交的数据
        table.setSubmitName("Info");

        // 设置响应单元格点击事件的js function
        poCtrl.setJsFunction_OnExcelCellClick("OnCellClick()");
        poCtrl.setWriter(workBook);
        //添加自定义按钮
        poCtrl.addCustomToolButton("保存", "Save", 1);
        //设置保存页面
        poCtrl.setSaveDataPage("SaveData");
        //打开Word文档
        poCtrl.webOpen("/doc/ExcelCellClick/test.xls", OpenModeType.xlsSubmitForm, "李斯");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        ModelAndView mv = new ModelAndView("ExcelCellClick/SubmitExcel");
        return mv;
    }

    @RequestMapping("/SaveData")
    public void SaveData(HttpServletRequest request, HttpServletResponse response) throws Exception{
        com.zhuozhengsoft.pageoffice.excelreader.Workbook workBook = new com.zhuozhengsoft.pageoffice.excelreader.Workbook(request, response);
        com.zhuozhengsoft.pageoffice.excelreader.Sheet sheet = workBook.openSheet("Sheet1");
        com.zhuozhengsoft.pageoffice.excelreader.Table table = sheet.openTable("B4:D8");
        String content = "";
        int result = 0;
        while (!table.getEOF()) {
            //获取提交的数值
            //DataFields.Count标识的是table的列数
            if (!table.getDataFields().getIsEmpty()) {
                content += "<br/>月份名称：" + table.getDataFields().get(0).getText();
                content += "<br/>计划完成量：" + table.getDataFields().get(1).getText();
                content += "<br/>实际完成量：" + table.getDataFields().get(2).getText();

                content += "<br/>*********************************************";
            }
            //循环进入下一行
            table.nextRow();
        }

        String html = "<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.01 Transitional//EN\">\n" +
                "<html>\n" +
                "<head>\n" +
                "    <title></title>\n" +
                "</head>\n" +
                "<body>\n" +
                "<form id=\"form1\">\n" +
                "    <div style=\"border: solid 1px gray;\">\n" +
                "        <div class=\"errTopArea\"\n" +
                "             style=\"text-align: left; border-bottom: solid 1px gray;\">\n" +
                "            [提示标题：这是一个开发人员可自定义的对话框]\n" +
                "        </div>\n" +
                "        <div class=\"errTxtArea\" style=\"height: 88%; text-align: left\">\n" +
                "            <b class=\"txt_title\">\n" +
                "                <div style=\" color:#FF0000;\">\n" +
                "                    提交的信息如下：\n" +
                "                </div>\n" + content +
                "                \n" +
                "            </b>\n" +
                "        </div>\n" +
                "        <div class=\"errBtmArea\" style=\"text-align: center;\">\n" +
                "            <input type=\"button\" class=\"btnFn\" value=\" 关闭 \"\n" +
                "                   onclick=\"window.opener=null;window.open('','_self');window.close();\"/>\n" +
                "        </div>\n" +
                "    </div>\n" +
                "</form>\n" +
                "</body>\n" +
                "</html>";
        response.setContentType("text/plain; charset=utf-8");
        response.setContentType("text/plain; charset=utf-8");
        response.getWriter().print(html);
        workBook.showPage(500, 400);
        workBook.close();
    }
}