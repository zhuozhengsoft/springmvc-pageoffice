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
@RequestMapping("SaveFirstPageAsImg")
public class SaveFirstPageAsImgController {

    @RequestMapping("/Word")
    public ModelAndView openWord(HttpServletRequest request, Map<String, Object> map) throws Exception{
        // 设置PageOffice组件服务页面
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage(request.getContextPath() + "/poserver.zz"); //此行必须
        poCtrl.addCustomToolButton("保存", "Save()", 1);
        poCtrl.addCustomToolButton("保存首页为图片", "SaveFirstAsImg()", 1);
        //设置保存页面
        poCtrl.setSaveFilePage("save");//设置处理文件保存的请求方法
        //打开word
        poCtrl.webOpen("/doc/SaveFirstPageAsImg/test.doc", OpenModeType.docNormalEdit, "李斯");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        ModelAndView mv = new ModelAndView("SaveFirstPageAsImg/Word");
        return mv;

    }

    @RequestMapping("/save")
    public void save(HttpServletRequest request, HttpServletResponse response) throws Exception{

        FileSaver fs = new FileSaver(request, response);
        //String aa=fs.getFileExtName();
        if (fs.getFileExtName().equals(".jpg")) {
            fs.saveToFile(request.getSession().getServletContext().getRealPath("images/SaveFirstPageAsImg/")  + fs.getFileName());
        } else {
            fs.saveToFile(request.getSession().getServletContext().getRealPath("doc/SaveFirstPageAsImg/")+fs.getFileName());
        }
        fs.setCustomSaveResult("ok");
        fs.close();




    }
}