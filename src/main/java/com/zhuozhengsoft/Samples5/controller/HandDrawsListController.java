package com.zhuozhengsoft.Samples5.controller;

import com.zhuozhengsoft.pageoffice.FileSaver;
import com.zhuozhengsoft.pageoffice.OpenModeType;
import com.zhuozhengsoft.pageoffice.PageOfficeCtrl;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.servlet.ModelAndView;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.util.Map;

@Controller
@RequestMapping("HandDrawsList")
public class HandDrawsListController {

    @RequestMapping("/Word")
    public ModelAndView openWord(HttpServletRequest request, Map<String, Object> map) throws Exception{
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage(request.getContextPath() + "/poserver.zz"); //此行必须
        poCtrl.setJsFunction_AfterDocumentOpened("AfterDocumentOpened()");
        poCtrl.addCustomToolButton("保存", "Save()", 1);
        poCtrl.addCustomToolButton("开始手写", "StartHandDraw()", 3);
        poCtrl.addCustomToolButton("设置线宽", "SetPenWidth()", 5);
        poCtrl.addCustomToolButton("设置颜色", "SetPenColor()", 5);
        poCtrl.addCustomToolButton("设置笔型", "SetPenType()", 5);
        poCtrl.addCustomToolButton("设置缩放", "SetPenZoom()", 5);
        poCtrl.setOfficeToolbars(false);//隐藏office工具栏
        //设置保存页面
        poCtrl.setSaveFilePage("save");//设置处理文件保存的请求方法
        //打开word
        poCtrl.webOpen("/doc/HandDrawsList/test.doc", OpenModeType.docHandwritingOnly, "John");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        ModelAndView mv = new ModelAndView("HandDrawsList/Word");
        return mv;
    }


    @RequestMapping("/save")
    public void save(HttpServletRequest request, HttpServletResponse response) throws Exception{
        FileSaver fs = new FileSaver(request, response);
        fs.saveToFile(request.getSession().getServletContext().getRealPath("doc/HandDrawsList/")+fs.getFileName());
        fs.close();
    }
}