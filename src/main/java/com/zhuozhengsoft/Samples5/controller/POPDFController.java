package com.zhuozhengsoft.Samples5.controller;

import com.zhuozhengsoft.pageoffice.OpenModeType;
import com.zhuozhengsoft.pageoffice.PDFCtrl;
import com.zhuozhengsoft.pageoffice.PageOfficeCtrl;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.servlet.ModelAndView;

import javax.servlet.http.HttpServletRequest;
import java.util.Map;

@Controller
@RequestMapping("POPDF")
public class POPDFController {

    @RequestMapping("/PDF")
    public ModelAndView openWord(HttpServletRequest request, Map<String, Object> map) throws Exception{
        PDFCtrl pdfCtrl = new PDFCtrl(request);
        pdfCtrl.setServerPage(request.getContextPath() + "/poserver.zz"); //此行必须
        // Create custom toolbar
        pdfCtrl.addCustomToolButton("打印", "PrintFile()", 6);
        pdfCtrl.addCustomToolButton("隐藏/显示书签", "SetBookmarks()", 0);
        pdfCtrl.addCustomToolButton("-", "", 0);
        pdfCtrl.addCustomToolButton("实际大小", "SetPageReal()", 16);
        pdfCtrl.addCustomToolButton("适合页面", "SetPageFit()", 17);
        pdfCtrl.addCustomToolButton("适合宽度", "SetPageWidth()", 18);
        pdfCtrl.addCustomToolButton("-", "", 0);
        pdfCtrl.addCustomToolButton("首页", "FirstPage()", 8);
        pdfCtrl.addCustomToolButton("上一页", "PreviousPage()", 9);
        pdfCtrl.addCustomToolButton("下一页", "NextPage()", 10);
        pdfCtrl.addCustomToolButton("尾页", "LastPage()", 11);
        pdfCtrl.addCustomToolButton("-", "", 0);
        pdfCtrl.addCustomToolButton("向左旋转90度", "SetRotateLeft()", 12);
        pdfCtrl.addCustomToolButton("向右旋转90度", "SetRotateRight()", 13);
        pdfCtrl.webOpen("/doc/POPDF/test.pdf");
        //打开Word文档
        map.put("pageoffice", pdfCtrl.getHtmlCode("PDFCtrl1"));
        ModelAndView mv = new ModelAndView("POPDF/PDF");
        return mv;
    }
}
