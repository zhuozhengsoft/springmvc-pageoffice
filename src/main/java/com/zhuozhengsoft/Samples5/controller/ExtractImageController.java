package com.zhuozhengsoft.Samples5.controller;

import com.zhuozhengsoft.pageoffice.OpenModeType;
import com.zhuozhengsoft.pageoffice.PageOfficeCtrl;
import com.zhuozhengsoft.pageoffice.wordwriter.DataRegion;
import com.zhuozhengsoft.pageoffice.wordwriter.WordDocument;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.servlet.ModelAndView;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.util.Map;

@Controller
@RequestMapping("ExtractImage")
public class ExtractImageController {

    @RequestMapping("/Word")
    public ModelAndView openWord(HttpServletRequest request, Map<String, Object> map) throws Exception{
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        //设置服务器页面
        poCtrl.setServerPage(request.getContextPath() + "/poserver.zz");
        poCtrl.addCustomToolButton("保存图片", "Save", 1);
        WordDocument wordDoc = new WordDocument();
        //打开数据区域，openDataRegion方法的参数代表Word文档中的书签名称
        DataRegion dataRegion1 = wordDoc.openDataRegion("PO_image");
        dataRegion1.setEditing(true);//放图片的数据区域是可以编辑的，其它部分不可编辑
        poCtrl.setWriter(wordDoc);
        //设置保存页面
        poCtrl.setSaveDataPage("SaveData");

        //打开word
        poCtrl.webOpen("/doc/ExtractImage/test.doc", OpenModeType.docSubmitForm, "李斯");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        ModelAndView mv = new ModelAndView("ExtractImage/Word");
        return mv;
    }

    @RequestMapping("/SaveData")
    public void SaveData(HttpServletRequest request, HttpServletResponse response) throws Exception {
        com.zhuozhengsoft.pageoffice.wordreader.WordDocument doc = new com.zhuozhengsoft.pageoffice.wordreader.WordDocument(request, response);
        com.zhuozhengsoft.pageoffice.wordreader.DataRegion dr = doc.openDataRegion("PO_image");
        //将提取的图片保存到服务器上，图片的名称为:a.jpg
        dr.openShape(1).saveAsJPG(request.getSession().getServletContext().getRealPath("doc/ExtractImage/") + "/a.jpg");
        doc.setCustomSaveResult(request.getSession().getServletContext().getRealPath("doc/ExtractImage/") + "\\a.jpg");
        doc.close();
    }
}