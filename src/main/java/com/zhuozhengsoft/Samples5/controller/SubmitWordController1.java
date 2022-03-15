package com.zhuozhengsoft.Samples5.controller;

import com.zhuozhengsoft.pageoffice.FileSaver;
import com.zhuozhengsoft.pageoffice.OpenModeType;
import com.zhuozhengsoft.pageoffice.PageOfficeCtrl;
import com.zhuozhengsoft.pageoffice.wordwriter.DataRegion;
import com.zhuozhengsoft.pageoffice.wordwriter.WordDocument;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.servlet.ModelAndView;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.PrintWriter;
import java.util.Map;

@Controller
//@RequestMapping("SubmitWord")
public class SubmitWordController1 {

    @RequestMapping(value = "/SubmitWord",method = RequestMethod.GET)
    public ModelAndView openWord(HttpServletRequest request, Map<String, Object> map) throws Exception{
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        //设置服务器页面
        poCtrl.setServerPage(request.getContextPath() + "/poserver.zz");

        WordDocument wordDoc = new WordDocument();
        //打开数据区域，openDataRegion方法的参数代表Word文档中的书签名称
        DataRegion dataRegion1 = wordDoc.openDataRegion("PO_userName");
        //设置DataRegion的可编辑性
        dataRegion1.setEditing(true);
        //为DataRegion赋值,此处的值可在页面中打开Word文档后自己进行修改
        dataRegion1.setValue("");
        DataRegion dataRegion2 = wordDoc.openDataRegion("PO_deptName");
        dataRegion2.setEditing(true);
        dataRegion2.setValue("");

        poCtrl.setWriter(wordDoc);
        //添加自定义按钮
        poCtrl.addCustomToolButton("保存", "Save", 1);
        //设置保存页面
        poCtrl.setSaveDataPage("SaveData");

        //打开word
        poCtrl.webOpen("/doc/SubmitWord/test.doc", OpenModeType.docSubmitForm, "李斯");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        ModelAndView mv = new ModelAndView("SubmitWord/SubmitWord");
        return mv;
    }

    @RequestMapping("/SaveData")
    public ModelAndView SaveData(HttpServletRequest request, HttpServletResponse response,Map<String, Object> map) {
response.setCharacterEncoding("utf-8");
    //    com.zhuozhengsoft.pageoffice.wordreader.WordDocument doc = new com.zhuozhengsoft.pageoffice.wordreader.WordDocument(request, response);
//        //获取提交的数值
//        com.zhuozhengsoft.pageoffice.wordreader.DataRegion dataUserName = doc.openDataRegion("PO_userName");
//        com.zhuozhengsoft.pageoffice.wordreader.DataRegion dataDeptName = doc.openDataRegion("PO_deptName");
        String content = "";
//        content += "公司名称：" + doc.getFormField("txtCompany");
//        content += "<br/>员工姓名：" + dataUserName.getValue();
//        content += "<br/>部门名称：" + dataDeptName.getValue();

        //doc.setCustomSaveResult(content);
//        doc.showPage(500,400);
//        doc.close();
        map.put("content",content);
        ModelAndView modelAndView = new ModelAndView("SubmitWord/SaveData");
        return modelAndView;

    }
}