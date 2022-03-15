package com.zhuozhengsoft.Samples5.controller;

import com.zhuozhengsoft.pageoffice.FileSaver;
import com.zhuozhengsoft.pageoffice.OpenModeType;
import com.zhuozhengsoft.pageoffice.PageOfficeCtrl;
import com.zhuozhengsoft.pageoffice.wordwriter.*;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.servlet.ModelAndView;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.awt.*;
import java.util.Map;

@Controller
@RequestMapping("DataRegionTable")
public class DataRegionTableController {

    @RequestMapping("/Default")
    public ModelAndView openWord(HttpServletRequest request, Map<String, Object> map) throws Exception{
        WordDocument doc = new WordDocument();
        //打开数据区域
        DataRegion dTable = doc.openDataRegion("PO_table");
        //设置数据区域可编辑性
        dTable.setEditing(true);
        //打开数据区域中的表格，OpenTable(index)方法中的index为word文档中表格的下标，从1开始
        Table table1 = doc.openDataRegion("PO_Table").openTable(1);
        //设置表格边框样式
        table1.getBorder().setLineColor(Color.green);
        table1.getBorder().setLineWidth(WdLineWidth.wdLineWidth050pt);
        // 设置表头单元格文本居中
        table1.openCellRC(1, 2).getParagraphFormat().setAlignment(WdParagraphAlignment.wdAlignParagraphCenter);
        table1.openCellRC(1, 3).getParagraphFormat().setAlignment(WdParagraphAlignment.wdAlignParagraphCenter);
        table1.openCellRC(2, 1).getParagraphFormat().setAlignment(WdParagraphAlignment.wdAlignParagraphCenter);
        table1.openCellRC(3, 1).getParagraphFormat().setAlignment(WdParagraphAlignment.wdAlignParagraphCenter);
        // 给表头单元格赋值
        table1.openCellRC(1, 2).setValue("产品1");
        table1.openCellRC(1, 3).setValue("产品2");
        table1.openCellRC(2, 1).setValue("A部门");
        table1.openCellRC(3, 1).setValue("B部门");

        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        //设置服务器页面
        poCtrl.setServerPage(request.getContextPath() + "/poserver.zz");
        poCtrl.setWriter(doc);

        //添加自定义按钮
        poCtrl.addCustomToolButton("保存", "Save", 1);
        poCtrl.addCustomToolButton("全屏/还原", "IsFullScreen", 4);

        //设置保存数据的页面
        poCtrl.setSaveDataPage("SaveData");

        //打开Word文档,当需要同时保存数据和保存文档时,OpenModeType必须是docSubmitForm模式。
        poCtrl.webOpen("/doc/DataRegionTable/test.doc", OpenModeType.docSubmitForm, "李斯");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        ModelAndView mv = new ModelAndView("DataRegionTable/Default");
        return mv;
    }

    @RequestMapping("/SaveData")
    public void SaveData(HttpServletRequest request, HttpServletResponse response) throws Exception{
        com.zhuozhengsoft.pageoffice.wordreader.WordDocument doc = new com.zhuozhengsoft.pageoffice.wordreader.WordDocument(request, response);
        com.zhuozhengsoft.pageoffice.wordreader.DataRegion dataReg = doc.openDataRegion("PO_table");
        com.zhuozhengsoft.pageoffice.wordreader.Table table = dataReg.openTable(1);
        StringBuilder dataStr = new StringBuilder();
        dataStr.append("表格中的各个单元的格数据为：<br/><br/>");
        for (int i = 1; i <= table.getRowsCount(); i++) {
            dataStr.append("<div style='width:220px;'>");
            for (int j = 1; j <= table.getColumnsCount(); j++) {
                dataStr.append("<div style='float:left;width:70px;border:1px solid red;'>" + table.openCellRC(i, j).getValue() + "</div>");
            }
            dataStr.append("</div>");
        }
        response.setContentType("text/plain; charset=utf-8");
        response.getWriter().print(dataStr.toString());
        //向客户端显示提交的数据
        doc.showPage(300, 300);
        doc.close();
    }
}