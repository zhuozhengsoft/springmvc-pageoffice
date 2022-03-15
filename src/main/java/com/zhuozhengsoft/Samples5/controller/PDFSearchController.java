package com.zhuozhengsoft.Samples5.controller;

import com.zhuozhengsoft.pageoffice.PDFCtrl;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.servlet.ModelAndView;

import javax.servlet.http.HttpServletRequest;
import java.util.Map;

@Controller
@RequestMapping("PDFSearch")
public class PDFSearchController {

    @RequestMapping("/PDF")
    public ModelAndView openWord(HttpServletRequest request, Map<String, Object> map) throws Exception{
        PDFCtrl pdfCtrl = new PDFCtrl(request);
        pdfCtrl.setServerPage(request.getContextPath() + "/poserver.zz"); //此行必须
        // Create custom toolbar
        pdfCtrl.addCustomToolButton("搜索", "SearchText()", 0);
        pdfCtrl.addCustomToolButton("搜索下一个", "SearchTextNext()", 0);
        pdfCtrl.addCustomToolButton("搜索上一个", "SearchTextPrev()", 0);
        pdfCtrl.addCustomToolButton("实际大小", "SetPageReal()", 16);
        pdfCtrl.addCustomToolButton("适合页面", "SetPageFit()", 17);
        pdfCtrl.addCustomToolButton("适合宽度", "SetPageWidth()", 18);
        pdfCtrl.webOpen("/doc/PDFSearch/test.pdf");
        //打开Word文档
        map.put("pageoffice", pdfCtrl.getHtmlCode("PDFCtrl1"));
        ModelAndView mv = new ModelAndView("PDFSearch/PDF");
        return mv;
    }
}
