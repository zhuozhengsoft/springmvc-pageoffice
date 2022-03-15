package com.zhuozhengsoft.Samples5.controller;

import com.zhuozhengsoft.pageoffice.FileSaver;
import com.zhuozhengsoft.pageoffice.OpenModeType;
import com.zhuozhengsoft.pageoffice.PageOfficeCtrl;
import com.zhuozhengsoft.pageoffice.wordwriter.DataRegion;
import com.zhuozhengsoft.pageoffice.wordwriter.WdParagraphAlignment;
import com.zhuozhengsoft.pageoffice.wordwriter.WordDocument;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.servlet.ModelAndView;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.awt.*;
import java.util.Map;

@Controller
@RequestMapping("DataRegionText")
public class DataRegionTextController {

    @RequestMapping("/Default")
    public ModelAndView openWord(HttpServletRequest request, Map<String, Object> map) throws Exception{
        WordDocument doc = new WordDocument();
        DataRegion d1 = doc.openDataRegion("d1");
        d1.getFont().setColor(Color.BLUE);//设置数据区域文本字体颜色
        d1.getFont().setName("华文彩云");//设置数据区域文本字体样式
        d1.getFont().setSize(16);//设置数据区域文本字体大小
        d1.getParagraphFormat().setAlignment(WdParagraphAlignment.wdAlignParagraphCenter);//设置数据区域文本对齐方式
        DataRegion d2 = doc.openDataRegion("d2");
        d2.getFont().setColor(Color.orange);//设置数据区域文本字体颜色
        d2.getFont().setName("黑体");//设置数据区域文本字体样式
        d2.getFont().setSize(14);//设置数据区域文本字体大小
        d2.getParagraphFormat().setAlignment(WdParagraphAlignment.wdAlignParagraphLeft);//设置数据区域文本对齐方式
        DataRegion d3 = doc.openDataRegion("d3");
        d3.getFont().setColor(Color.magenta);//设置数据区域文本字体颜色
        d3.getFont().setName("华文行楷");//设置数据区域文本字体样式
        d3.getFont().setSize(12);//设置数据区域文本字体大小
        d3.getParagraphFormat().setAlignment(WdParagraphAlignment.wdAlignParagraphRight);//设置数据区域文本对齐方式

        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setWriter(doc);
        //设置服务器页面
        poCtrl.setServerPage(request.getContextPath() + "/poserver.zz"); //此行必须
        //打开word
        poCtrl.webOpen("/doc/DataRegionText/test.doc", OpenModeType.docNormalEdit, "李斯");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        ModelAndView mv = new ModelAndView("DataRegionText/Default");
        return mv;
    }
    @RequestMapping("/Default2")
    public ModelAndView Default2(HttpServletRequest request, Map<String, Object> map) throws Exception{
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        //设置服务器页面
        poCtrl.setServerPage(request.getContextPath() + "/poserver.zz"); //此行必须
        //打开word
        poCtrl.webOpen("/doc/DataRegionText/test.doc", OpenModeType.docNormalEdit, "李斯");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        ModelAndView mv = new ModelAndView("DataRegionText/Default");
        return mv;
    }

}