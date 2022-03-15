package com.zhuozhengsoft.Samples5.controller;

import com.zhuozhengsoft.pageoffice.OpenModeType;
import com.zhuozhengsoft.pageoffice.PageOfficeCtrl;
import com.zhuozhengsoft.pageoffice.wordwriter.Table;
import com.zhuozhengsoft.pageoffice.wordwriter.WordDocument;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.servlet.ModelAndView;

import javax.servlet.http.HttpServletRequest;
import java.util.Map;

@Controller
@RequestMapping("WordCompare")
public class WordCompareController {

    @RequestMapping("/compare")
    public ModelAndView openWord(HttpServletRequest request, Map<String, Object> map) throws Exception{
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage(request.getContextPath() + "/poserver.zz");
        // Create custom toolbar
        poCtrl.addCustomToolButton("保存", "SaveDocument()", 1);
        poCtrl.addCustomToolButton("显示A文档", "ShowFile1View()", 0);
        poCtrl.addCustomToolButton("显示B文档", "ShowFile2View()", 0);
        poCtrl.addCustomToolButton("显示比较结果", "ShowCompareView()", 0);
        poCtrl.wordCompare("/doc/WordCompare/aaa1.doc", "/doc/WordCompare/aaa2.doc", OpenModeType.docAdmin, "张三");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        ModelAndView mv = new ModelAndView("WordCompare/compare");
        return mv;
    }
}
