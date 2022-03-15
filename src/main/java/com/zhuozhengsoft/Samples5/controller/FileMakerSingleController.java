package com.zhuozhengsoft.Samples5.controller;

import com.zhuozhengsoft.pageoffice.*;
import com.zhuozhengsoft.pageoffice.wordwriter.WordDocument;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
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
@RequestMapping("FileMakerSingle")
public class FileMakerSingleController {

    @RequestMapping("/Default")
    public ModelAndView openWord1(HttpServletRequest request, Map<String, Object> map) throws Exception{
        ModelAndView mv = new ModelAndView("FileMakerSingle/Default");
        return mv;
    }

    @RequestMapping("/FileMakerSingle")
    public ModelAndView fileMakerSingle(HttpServletRequest request, Map<String, Object> map) throws Exception{
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
        fmCtrl.setFileTitle("newfilename.doc");
        fmCtrl.fillDocument("/doc/FileMakerSingle/template.doc", DocumentOpenType.Word);
        map.put("pageoffice", fmCtrl.getHtmlCode("FileMakerCtrl1"));
        map.put("type" ,request.getParameter("type"));
        ModelAndView mv = new ModelAndView("FileMakerSingle/FileMakerSingle");
        return mv;
    }

    @RequestMapping("/save")
    public void save(HttpServletRequest request, HttpServletResponse response) throws Exception{
        FileSaver fs = new FileSaver(request, response);
        String fileName = "\\maker" + fs.getFileExtName();
        fs.saveToFile(request.getSession().getServletContext().getRealPath("doc/FileMakerSingle/") + "/" + "maker.doc");
        fs.close();
    }

    @RequestMapping("/DownWord")
    public void downWord(HttpServletRequest request, HttpServletResponse response) throws Exception{
        String filePath =request.getSession().getServletContext().getRealPath("doc/FileMakerSingle/maker.doc");
        int fileSize =(int)new File(filePath).length();

        response.reset();
        response.setContentType("application/msword"); // application/x-excel, application/ms-powerpoint, application/pdf
        response.setHeader("Content-Disposition", "attachment; filename=maker.doc");
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

    @RequestMapping("/OpenWord")
    public ModelAndView OpenWord(HttpServletRequest request, Map<String, Object> map) throws Exception{
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        //设置服务器页面
        poCtrl.setCustomToolbar(false);
        poCtrl.setServerPage(request.getContextPath() + "/poserver.zz");
        poCtrl.webOpen("/doc/FileMakerSingle/maker.doc", OpenModeType.docReadOnly, "张佚名");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        ModelAndView mv = new ModelAndView("FileMakerSingle/OpenWord");
        return mv;
    }
}