package com.zhuozhengsoft.Samples5.controller;

import com.zhuozhengsoft.pageoffice.DocumentOpenType;
import com.zhuozhengsoft.pageoffice.FileMakerCtrl;
import com.zhuozhengsoft.pageoffice.FileSaver;
import com.zhuozhengsoft.pageoffice.PDFCtrl;
import com.zhuozhengsoft.pageoffice.wordwriter.WordDocument;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.servlet.ModelAndView;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.Map;

@Controller
@RequestMapping("FileMakerPDF")
public class FileMakerPDFController {

    @RequestMapping("/Default")
    public ModelAndView openWord1(HttpServletRequest request, Map<String, Object> map) throws Exception{
        ModelAndView mv = new ModelAndView("FileMakerPDF/Default");
        return mv;
    }

    @RequestMapping("/FileMakerPDF")
    public ModelAndView openWord(HttpServletRequest request, Map<String, Object> map) throws Exception{
        FileMakerCtrl fmCtrl = new FileMakerCtrl(request);
        fmCtrl.setServerPage(request.getContextPath() + "/poserver.zz");
        WordDocument doc = new WordDocument();
        //禁用右击事件
        doc.setDisableWindowRightClick(true);
        //给数据区域赋值，即把数据填充到模板中相应的位置
        doc.openDataRegion("PO_company").setValue("北京卓正志远软件有限公司  ");
        fmCtrl.setSaveFilePage("save");
        fmCtrl.setWriter(doc);
        fmCtrl.setJsFunction_OnProgressComplete("OnProgressComplete()");
        fmCtrl.fillDocumentAsPDF("/doc/FileMakerPDF/template.doc", DocumentOpenType.Word, "template.pdf");
        map.put("pageoffice", fmCtrl.getHtmlCode("FileMakerCtrl1"));
        map.put("type", request.getParameter("type"));
        ModelAndView mv = new ModelAndView("FileMakerPDF/FileMakerPDF");
        return mv;
    }

    @RequestMapping("/save")
    public void save(HttpServletRequest request, HttpServletResponse response) throws Exception{
        FileSaver fs = new FileSaver(request, response);
        fs.saveToFile(request.getSession().getServletContext().getRealPath("doc/FileMakerPDF/") + "/" + fs.getFileName());
        fs.close();
    }

    @RequestMapping("/OpenPDF")
    public ModelAndView OpenPDF(HttpServletRequest request, Map<String, Object> map) throws Exception{
        PDFCtrl pdfCtrl1 = new PDFCtrl(request);
        pdfCtrl1.setServerPage(request.getContextPath() + "/poserver.zz"); //此行必须
        // Create custom toolbar
        pdfCtrl1.addCustomToolButton("打印", "PrintFile()", 6);
        pdfCtrl1.addCustomToolButton("隐藏/显示书签", "SetBookmarks()", 0);
        pdfCtrl1.addCustomToolButton("-", "", 0);
        pdfCtrl1.addCustomToolButton("实际大小", "SetPageReal()", 16);
        pdfCtrl1.addCustomToolButton("适合页面", "SetPageFit()", 17);
        pdfCtrl1.addCustomToolButton("适合宽度", "SetPageWidth()", 18);
        pdfCtrl1.addCustomToolButton("-", "", 0);
        pdfCtrl1.addCustomToolButton("首页", "FirstPage()", 8);
        pdfCtrl1.addCustomToolButton("上一页", "PreviousPage()", 9);
        pdfCtrl1.addCustomToolButton("下一页", "NextPage()", 10);
        pdfCtrl1.addCustomToolButton("尾页", "LastPage()", 11);
        pdfCtrl1.addCustomToolButton("-", "", 0);
        pdfCtrl1.addCustomToolButton("向左旋转90度", "SetRotateLeft()", 12);
        pdfCtrl1.addCustomToolButton("向右旋转90度", "SetRotateRight()", 13);
        pdfCtrl1.webOpen("/doc/FileMakerPDF/template.pdf");
        map.put("pageoffice", pdfCtrl1.getHtmlCode("FileMakerCtrl1"));
        ModelAndView mv = new ModelAndView("FileMakerPDF/OpenPDF");
        return mv;
    }

    @RequestMapping("/DownPDF")
    public void DownPDF(HttpServletRequest request,HttpServletResponse response, Map<String, Object> map) throws Exception{
        String filePath =request.getSession().getServletContext().getRealPath("doc/FileMakerPDF/template.pdf");
        int fileSize =(int)new File(filePath).length();
        response.reset();
        response.setContentType("application/pdf"); // application/x-excel, application/ms-powerpoint, application/pdf
        response.setHeader("Content-Disposition", "attachment; filename=template.pdf");
        response.setContentLength(fileSize);

        OutputStream outputStream = response.getOutputStream();
        InputStream inputStream = new FileInputStream(filePath);
        byte[] buffer = new byte[10240];
        int i = -1;
        while ((i = inputStream.read(buffer)) != -1) {
            outputStream.write(buffer, 0, i);
        }
        outputStream.flush();
        outputStream.close();
        inputStream.close();
    }
}