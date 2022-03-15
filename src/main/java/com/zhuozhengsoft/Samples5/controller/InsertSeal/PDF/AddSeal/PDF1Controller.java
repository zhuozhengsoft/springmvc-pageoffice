package com.zhuozhengsoft.Samples5.controller.InsertSeal.PDF.AddSeal;

import com.zhuozhengsoft.pageoffice.FileSaver;
import com.zhuozhengsoft.pageoffice.PDFCtrl;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.servlet.ModelAndView;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.util.Map;

@Controller
@RequestMapping("InsertSeal/PDF/AddSeal1")
public class PDF1Controller {

    @RequestMapping("/PDF1")
    public ModelAndView openWord(HttpServletRequest request, Map<String, Object> map) throws Exception{
        PDFCtrl pdfCtrl = new PDFCtrl(request);
        pdfCtrl.setServerPage(request.getContextPath() + "/poserver.zz"); //此行必须
        //设置保存页面
        pdfCtrl.setSaveFilePage("save");

        // Create custom toolbar
        pdfCtrl.addCustomToolButton("保存", "Save()", 1);
        pdfCtrl.addCustomToolButton("加盖印章", "InsertSeal()", 2);
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
        pdfCtrl.webOpen("/doc/InsertSeal/PDF/AddSeal1/test.pdf");
        //打开Word文档
        map.put("pageoffice", pdfCtrl.getHtmlCode("PDFCtrl1"));
        ModelAndView mv = new ModelAndView("InsertSeal/PDF/AddSeal1/PDF1");
        return mv;
    }

    @RequestMapping("/save")
    public void AddSeal6Word6Save(HttpServletRequest request, HttpServletResponse response) throws Exception{
        FileSaver fs = new FileSaver(request, response);
        fs.saveToFile(request.getSession().getServletContext().getRealPath("doc/InsertSeal/PDF/AddSeal1/")+fs.getFileName());
        fs.close();
    }
}
