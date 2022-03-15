package com.zhuozhengsoft.Samples5.controller;

import com.zhuozhengsoft.pageoffice.OpenModeType;
import com.zhuozhengsoft.pageoffice.PageOfficeCtrl;
import com.zhuozhengsoft.pageoffice.wordwriter.DataRegion;
import com.zhuozhengsoft.pageoffice.wordwriter.WordDocument;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.servlet.ModelAndView;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.FileOutputStream;
import java.util.Map;

@Controller
@RequestMapping("SplitWord")
public class SplitWordController {

    @RequestMapping("/Word")
    public ModelAndView openWord(HttpServletRequest request, Map<String, Object> map) throws Exception{
        //******************************卓正PageOffice组件的使用*******************************
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage(request.getContextPath() + "/poserver.zz"); //此行必须
        WordDocument wordDoc = new WordDocument();
        //打开数据区域，openDataRegion方法的参数代表Word文档中的书签名称
        DataRegion dataRegion1 = wordDoc.openDataRegion("PO_test1");
        dataRegion1.setSubmitAsFile(true);
        DataRegion dataRegion2 = wordDoc.openDataRegion("PO_test2");
        dataRegion2.setSubmitAsFile(true);
        dataRegion2.setEditing(true);
        DataRegion dataRegion3 = wordDoc.openDataRegion("PO_test3");
        dataRegion3.setSubmitAsFile(true);

        poCtrl.setWriter(wordDoc);
        poCtrl.addCustomToolButton("保存", "Save()", 1);
        //设置保存页面
        poCtrl.setSaveDataPage("SaveData");

        //打开word
        poCtrl.webOpen("/doc/SplitWord/test.doc", OpenModeType.docSubmitForm, "李斯");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        ModelAndView mv = new ModelAndView("SplitWord/Word");
        return mv;
    }

    @RequestMapping("/SaveData")
    public void SaveData(HttpServletRequest request, HttpServletResponse response) throws Exception {
        com.zhuozhengsoft.pageoffice.wordreader.WordDocument doc = new com.zhuozhengsoft.pageoffice.wordreader.WordDocument(request, response);
        String filePath = request.getSession().getServletContext().getRealPath("doc/SplitWord/") + "/";
        System.out.println(filePath);
        byte[] bWord;
        
        com.zhuozhengsoft.pageoffice.wordreader.DataRegion dr1 = doc.openDataRegion("PO_test1");
        bWord = dr1.getFileBytes();
        FileOutputStream fos1 = new FileOutputStream(filePath + "new1.doc");
        fos1.write(bWord);
        fos1.flush();
        fos1.close();

        com.zhuozhengsoft.pageoffice.wordreader.DataRegion dr2 = doc.openDataRegion("PO_test2");
        bWord = dr2.getFileBytes();
        FileOutputStream fos2 = new FileOutputStream(filePath + "new2.doc");
        fos2.write(bWord);
        fos2.flush();
        fos2.close();

        com.zhuozhengsoft.pageoffice.wordreader.DataRegion dr3 = doc.openDataRegion("PO_test3");
        bWord = dr3.getFileBytes();
        FileOutputStream fos3 = new FileOutputStream(filePath + "new3.doc");
        fos3.write(bWord);
        fos3.flush();
        fos3.close();
        
        doc.close();
    }
}