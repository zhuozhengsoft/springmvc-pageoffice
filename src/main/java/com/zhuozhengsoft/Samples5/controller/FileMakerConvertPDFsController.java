package com.zhuozhengsoft.Samples5.controller;

import com.zhuozhengsoft.pageoffice.*;
import com.zhuozhengsoft.pageoffice.wordwriter.WordDocument;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.servlet.ModelAndView;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.util.Map;

@Controller
@RequestMapping("FileMakerConvertPDFs")
public class FileMakerConvertPDFsController {

    @RequestMapping("/index")
    public ModelAndView openWord1(HttpServletRequest request, Map<String, Object> map) throws Exception {
        String url = request.getSession().getServletContext().getRealPath("/doc/PrintFiles/" + "/");
        request.setAttribute("url", url);
        ModelAndView mv = new ModelAndView("FileMakerConvertPDFs/index");
        return mv;
    }

    @RequestMapping("/Edit")
    public ModelAndView edit(HttpServletRequest request, Map<String, Object> map) throws Exception {
        String filePath = "";
        String id = request.getParameter("id").trim();
        if ("1".equals(id)) {
            filePath = request.getSession().getServletContext().getRealPath("/doc/FileMakerConvertPDFs/PageOffice产品简介.doc");
        }
        if ("2".equals(id)) {
            filePath = request.getSession().getServletContext().getRealPath("/doc/FileMakerConvertPDFs/Pageoffice客户端安装步骤.doc");
        }
        if ("3".equals(id)) {
            filePath = request.getSession().getServletContext().getRealPath("/doc/FileMakerConvertPDFs/PageOffice的应用领域.doc");
        }
        if ("4".equals(id)) {
            filePath = request.getSession().getServletContext().getRealPath("/doc/FileMakerConvertPDFs/PageOffice产品对客户端环境要求.doc");
        }
        //******************************卓正PageOffice组件的使用*******************************
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage(request.getContextPath() + "/poserver.zz"); //此行必须
        poCtrl.setSaveFilePage("save");//如要保存文件，此行必须
        poCtrl.addCustomToolButton("保存", "Save()", 1);//添加自定义工具栏按钮
        //打开文件
        poCtrl.webOpen(filePath, OpenModeType.docNormalEdit, "张三");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        ModelAndView mv = new ModelAndView("FileMakerConvertPDFs/Edit");
        return mv;
    }

    @RequestMapping("/Convert")
    public ModelAndView Convert(HttpServletRequest request, Map<String, Object> map) throws Exception {
        String filePath = "";
        String id = request.getParameter("id").trim();
        if ("1".equals(id)) {
            filePath = request.getSession().getServletContext().getRealPath("/doc/FileMakerConvertPDFs/PageOffice产品简介.doc");
        }
        if ("2".equals(id)) {
            filePath = request.getSession().getServletContext().getRealPath("/doc/FileMakerConvertPDFs/Pageoffice客户端安装步骤.doc");
        }
        if ("3".equals(id)) {
            filePath = request.getSession().getServletContext().getRealPath("/doc/FileMakerConvertPDFs/PageOffice的应用领域.doc");
        }
        if ("4".equals(id)) {
            filePath = request.getSession().getServletContext().getRealPath("/doc/FileMakerConvertPDFs/PageOffice产品对客户端环境要求.doc");
        }
        FileMakerCtrl fmCtrl = new FileMakerCtrl(request);
        fmCtrl.setServerPage(request.getContextPath() + "/poserver.zz");
        fmCtrl.setJsFunction_OnProgressComplete("OnProgressComplete()");
        fmCtrl.setSaveFilePage("save");
        fmCtrl.fillDocumentAsPDF(filePath, DocumentOpenType.Word, "a.pdf");
        map.put("pageoffice", fmCtrl.getHtmlCode("FileMakerCtrl1"));
        ModelAndView mv = new ModelAndView("FileMakerConvertPDFs/Convert");
        return mv;
    }

    @RequestMapping("/save")
    public void save(HttpServletRequest request, HttpServletResponse response) throws Exception {

        FileSaver fs = new FileSaver(request, response);
        fs.saveToFile(request.getSession().getServletContext().getRealPath("doc/FileMakerConvertPDFs/") + "/" + fs.getFileName());
        fs.close();

    }
}