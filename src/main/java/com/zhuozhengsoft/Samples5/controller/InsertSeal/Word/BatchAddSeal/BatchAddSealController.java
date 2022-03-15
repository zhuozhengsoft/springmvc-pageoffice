package com.zhuozhengsoft.Samples5.controller.InsertSeal.Word.BatchAddSeal;

import com.zhuozhengsoft.pageoffice.*;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.servlet.ModelAndView;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.util.Map;

@Controller
@RequestMapping("InsertSeal/Word/BatchAddSeal")
public class BatchAddSealController {
    @RequestMapping("/Default")
    public ModelAndView Default(HttpServletRequest request, Map<String, Object> map) throws Exception{
        String url = request.getSession().getServletContext().getRealPath("doc/InsertSeal/Word/BatchAddSeal/" + "/");
        request.setAttribute("url",url);
        ModelAndView mv = new ModelAndView("InsertSeal/Word/BatchAddSeal/Default");
        return mv;
    }

    @RequestMapping("/AddSeal")
    public ModelAndView AddSeal(HttpServletRequest request, Map<String, Object> map) throws Exception{
       // String filePath = request.getSession().getServletContext().getRealPath("InsertSeal/doc/");
        String filePath = "";
        String id = request.getParameter("id").trim();
        if ("1".equals(id)) {
            filePath = "/doc/InsertSeal/Word/BatchAddSeal/test1.doc";
        }
        if ("2".equals(id)) {
            filePath = "/doc/InsertSeal/Word/BatchAddSeal/test2.doc";
        }
        if ("3".equals(id)) {
            filePath = "/doc/InsertSeal/Word/BatchAddSeal/test3.doc";
        }
        if ("4".equals(id)) {
            filePath = "/doc/InsertSeal/Word/BatchAddSeal/test4.doc";
        }

        FileMakerCtrl fmCtrl = new FileMakerCtrl(request);
        fmCtrl.setServerPage(request.getContextPath() + "/poserver.zz");
        fmCtrl.setJsFunction_AfterDocumentOpened("AfterDocumentOpened()");
        fmCtrl.setJsFunction_OnProgressComplete("OnProgressComplete()");
        fmCtrl.setSaveFilePage("save");
        fmCtrl.fillDocument(filePath, DocumentOpenType.Word);
        //打开Word文档
        map.put("pageoffice", fmCtrl.getHtmlCode("FileMakerCtrl1"));
        ModelAndView mv = new ModelAndView("InsertSeal/Word/BatchAddSeal/AddSeal");
        return mv;
    }

    @RequestMapping("/save")
    public void AddSeal1Word1Save(HttpServletRequest request, HttpServletResponse response) throws Exception{
        FileSaver fs = new FileSaver(request, response);
        fs.saveToFile(request.getSession().getServletContext().getRealPath("doc/InsertSeal/Word/BatchAddSeal/")+fs.getFileName());
        fs.close();
    }

    @RequestMapping("/Edit")
    public ModelAndView Edit(HttpServletRequest request, Map<String, Object> map) throws Exception{
        String path = request.getContextPath();
        String filePath = "";
        String id = request.getParameter("id").trim();
        if ("1".equals(id)) {
            filePath = "doc/test1.doc";
        }
        if ("2".equals(id)) {
            filePath = "doc/test2.doc";
        }
        if ("3".equals(id)) {
            filePath = "doc/test3.doc";
        }
        if ("4".equals(id)) {
            filePath = "doc/test4.doc";
        }

        //******************************卓正PageOffice组件的使用*******************************
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage(request.getContextPath() + "/poserver.zz"); //此行必须
        //添加自定义按钮
        poCtrl.addCustomToolButton("保存", "Save", 1);
        //设置保存页面
        poCtrl.setSaveFilePage("save");
        //打开Word文档
        poCtrl.webOpen("/doc/InsertSeal/Word/AddSeal2/word2.doc", OpenModeType.docNormalEdit, "张三");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        ModelAndView mv = new ModelAndView("InsertSeal/Word/BatchAddSeal/Edit");
        return mv;
    }


}
