package com.zhuozhengsoft.Samples5.controller;

import com.zhuozhengsoft.pageoffice.FileSaver;
import com.zhuozhengsoft.pageoffice.OpenModeType;
import com.zhuozhengsoft.pageoffice.PageOfficeCtrl;
import com.zhuozhengsoft.pageoffice.wordwriter.DataRegion;
import com.zhuozhengsoft.pageoffice.wordwriter.WordDocument;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.servlet.ModelAndView;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.util.Map;

@Controller
@RequestMapping("SetHandDrawByUser")
public class SetHandDrawByUserController {
    @RequestMapping("/Default")
    public ModelAndView openWord(HttpServletRequest request) throws Exception{
        ModelAndView mv = new ModelAndView("SetHandDrawByUser/Default");
        return mv;
    }

    @RequestMapping("/SetHandDrawByUserName")
    public ModelAndView openWord(HttpServletRequest request, Map<String, Object> map) throws Exception{
        String userName = request.getParameter("userName");

        if (userName.equals("zhangsan")) userName = "张三";
        if (userName.equals("lisi")) userName = "李四";
        if (userName.equals("wangwu")) userName = "王五";
        //***************************卓正PageOffice组件的使用********************************
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.addCustomToolButton("保存", "Save", 1);
        poCtrl.addCustomToolButton("领导圈阅", "StartHandDraw", 3);
        poCtrl.addCustomToolButton("全屏/还原", "IsFullScreen", 4);
        poCtrl.setJsFunction_AfterDocumentOpened("ShowByUserName");
        poCtrl.setServerPage(request.getContextPath() + "/poserver.zz"); //此行必须
        //设置保存页面
        poCtrl.setSaveFilePage("save");//设置处理文件保存的请求方法
        poCtrl.setMenubar(false);
        //打开word
        poCtrl.webOpen("/doc/SetHandDrawByUser/test.doc", OpenModeType.docSubmitForm, userName);

        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        map.put("userName", userName);
        ModelAndView mv = new ModelAndView("SetHandDrawByUser/SetHandDrawByUserName");
        return mv;
    }

    @RequestMapping("/save")
    public void save(HttpServletRequest request, HttpServletResponse response) throws Exception{
        FileSaver fs = new FileSaver(request, response);
        fs.saveToFile(request.getSession().getServletContext().getRealPath("doc/SetHandDrawByUser/")+fs.getFileName());
        fs.close();
    }
}