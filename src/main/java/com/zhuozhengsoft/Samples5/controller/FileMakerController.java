package com.zhuozhengsoft.Samples5.controller;

import com.zhuozhengsoft.pageoffice.DocumentOpenType;
import com.zhuozhengsoft.pageoffice.FileMakerCtrl;
import com.zhuozhengsoft.pageoffice.FileSaver;
import com.zhuozhengsoft.pageoffice.wordwriter.WordDocument;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.servlet.ModelAndView;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.util.Map;

@Controller
@RequestMapping("FileMaker")
public class FileMakerController {

    @RequestMapping("/Default")
    public ModelAndView openWord1(HttpServletRequest request, Map<String, Object> map) throws Exception {
        String url = request.getSession().getServletContext().getRealPath("/doc/FileMaker/" + "/");
        request.setAttribute("url", url);
        ModelAndView mv = new ModelAndView("FileMaker/Default");
        return mv;
    }

    @RequestMapping("/FileMaker")
    public ModelAndView openWord(HttpServletRequest request, Map<String, Object> map) throws Exception {
        FileMakerCtrl fmCtrl = new FileMakerCtrl(request);
        fmCtrl.setServerPage(request.getContextPath() + "/poserver.zz");
        String id = request.getParameter("id");

        if (id != null && id.length() > 0) {
            WordDocument doc = new WordDocument();
            //禁用右击事件
            doc.setDisableWindowRightClick(true);
            //给数据区域赋值，即把数据填充到模板中相应的位置
            doc.openDataRegion("PO_company").setValue("北京卓正志远软件有限公司  " + id);
            fmCtrl.setSaveFilePage("SaveMaker?id=" + id);
            fmCtrl.setWriter(doc);
            fmCtrl.setJsFunction_OnProgressComplete("OnProgressComplete()");
            fmCtrl.setFileTitle("newfilename.doc");
            fmCtrl.fillDocument("/doc/FileMaker/template.doc", DocumentOpenType.Word);
        }
        map.put("pageoffice", fmCtrl.getHtmlCode("FileMakerCtrl1"));
        ModelAndView mv = new ModelAndView("FileMaker/FileMaker");
        return mv;
    }

    @RequestMapping("/SaveMaker")
    public void save(HttpServletRequest request, HttpServletResponse response) throws Exception {

        FileSaver fs = new FileSaver(request, response);
        String id = request.getParameter("id");
        String err = "aaaa";
        if (id != null && id.length() > 0) {
            String fileName = "\\maker" + id + fs.getFileExtName();
            fs.saveToFile(request.getSession().getServletContext()
                    .getRealPath("doc/FileMaker/")
                    + fileName);
        } else {
            err = "<script>alert('未获得文件名称');</script>";
        }

        if (err.length()>0 ){
            System.out.println(err);
            response.getWriter().write("<script>alert('未获得文件名称');</script>");
        }
        fs.close();
    }
}