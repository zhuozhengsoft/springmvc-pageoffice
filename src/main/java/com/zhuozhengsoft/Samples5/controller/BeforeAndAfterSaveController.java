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
@RequestMapping("BeforeAndAfterSave")
public class BeforeAndAfterSaveController {

    @RequestMapping("/Word")
    public ModelAndView openWord(HttpServletRequest request, Map<String, Object> map) throws  Exception{
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage(request.getContextPath() + "/poserver.zz");
        // 设置文件保存之前执行的事件
        poCtrl.setJsFunction_BeforeDocumentSaved("BeforeDocumentSaved()");
        // 设置文件保存之后执行的事件
        poCtrl.setJsFunction_AfterDocumentSaved("AfterDocumentSaved()");
        poCtrl.setSaveFilePage("save");
        //打开Word文档
        poCtrl.webOpen("/doc/BeforeAndAfterSave/test.doc", OpenModeType.docNormalEdit, "张三");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        ModelAndView mv = new ModelAndView("BeforeAndAfterSave/Word");
        return mv;
    }
    @RequestMapping("/save")
    public void save(HttpServletRequest request, HttpServletResponse response)throws Exception {
        FileSaver fs = new FileSaver(request, response);
        fs.saveToFile(request.getSession().getServletContext().getRealPath("doc/BeforeAndAfterSave/")+fs.getFileName());
        fs.close();
    }
}
