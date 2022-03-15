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
@RequestMapping("ConcurrencyCtrl")
public class ConcurrencyCtrlController {

    @RequestMapping("/Default")
    public ModelAndView defaultPage(HttpServletRequest request, Map<String, Object> map) {
        ModelAndView mv = new ModelAndView("ConcurrencyCtrl/Default");
        return mv;
    }

    @RequestMapping("/Word")
    public ModelAndView openWord(HttpServletRequest request, Map<String, Object> map) throws Exception{
        String userName = "somebody";
        String userId = request.getParameter("userid").toString();
        if (userId.equals("1")) {
            userName = "张三";
        } else {
            userName = "李四";
        }

        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage(request.getContextPath() + "/poserver.zz");
        poCtrl.addCustomToolButton("保存", "Save", 1);
        poCtrl.setSaveFilePage("save");
        //设置并发控制时间
        poCtrl.setTimeSlice(20);
        //打开word
        poCtrl.webOpen("/doc/ConcurrencyCtrl/test.doc", OpenModeType.docNormalEdit, userName);
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        map.put("userName", userName);
        ModelAndView mv = new ModelAndView("ConcurrencyCtrl/Word");
        return mv;
    }

    @RequestMapping("/save")
    public void save(HttpServletRequest request, HttpServletResponse response) throws Exception{
        FileSaver fs = new FileSaver(request, response);
        fs.saveToFile(request.getSession().getServletContext().getRealPath("doc/ConcurrencyCtrl/")+fs.getFileName());
        fs.close();
    }
}