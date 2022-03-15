package com.zhuozhengsoft.Samples5.controller;

import com.zhuozhengsoft.pageoffice.PDFCtrl;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.servlet.ModelAndView;

import javax.servlet.http.HttpServletRequest;
import java.util.Map;

@Controller
@RequestMapping("OpenImage")
public class OpenImageController {

    @RequestMapping("/Image")
    public ModelAndView openWord(HttpServletRequest request, Map<String, Object> map) throws Exception{
        PDFCtrl poCtrl1 = new PDFCtrl(request);
        poCtrl1.setServerPage(request.getContextPath() + "/poserver.zz"); //此行必须
        // Create custom toolbar
        poCtrl1.addCustomToolButton("打印", "Print()", 6);
        poCtrl1.addCustomToolButton("-", "", 0);
        poCtrl1.addCustomToolButton("实际大小", "SetPageReal()", 16);
        poCtrl1.addCustomToolButton("适合页面", "SetPageFit()", 17);
        poCtrl1.addCustomToolButton("适合宽度", "SetPageWidth()", 18);
        poCtrl1.addCustomToolButton("-", "", 0);
        poCtrl1.addCustomToolButton("左转", "RotateLeft()", 12);
        poCtrl1.addCustomToolButton("右转", "RotateRight()", 13);
        poCtrl1.addCustomToolButton("-", "", 0);
        poCtrl1.addCustomToolButton("放大", "ZoomIn()", 14);
        poCtrl1.addCustomToolButton("缩小", "ZoomOut()", 15);
        poCtrl1.addCustomToolButton("-", "", 0);
        poCtrl1.addCustomToolButton("全屏", "SwitchFullScreen()", 4);
        poCtrl1.webOpen("/doc/OpenImage/test.jpg");
        //打开Word文档
        map.put("pageoffice", poCtrl1.getHtmlCode("PDFCtrl1"));
        ModelAndView mv = new ModelAndView("OpenImage/Image");
        return mv;
    }
}
