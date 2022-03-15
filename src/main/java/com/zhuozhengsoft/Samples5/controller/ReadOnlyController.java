package com.zhuozhengsoft.Samples5.controller;

import com.zhuozhengsoft.pageoffice.OpenModeType;
import com.zhuozhengsoft.pageoffice.PageOfficeCtrl;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.servlet.ModelAndView;

import javax.servlet.http.HttpServletRequest;
import java.util.Map;

@Controller
@RequestMapping("ReadOnly")
public class ReadOnlyController {

    @RequestMapping("/OpenWord")
    public ModelAndView openWord(HttpServletRequest request, Map<String, Object> map) throws Exception{
        //设置PageOffice服务器组件
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage(request.getContextPath() + "/poserver.zz"); //此行必须

        poCtrl.setAllowCopy(false);//禁止拷贝
        poCtrl.setMenubar(false);//隐藏菜单栏
        poCtrl.setOfficeToolbars(false);//隐藏Office工具条
        poCtrl.setCustomToolbar(false);//隐藏自定义工具栏
        poCtrl.setJsFunction_AfterDocumentOpened("AfterDocumentOpened");
        //设置页面的显示标题
        poCtrl.setCaption("演示：文件在线安全浏览");
        //打开文件
        //打开文件
        poCtrl.webOpen("/doc/ReadOnly/template.doc", OpenModeType.docReadOnly, "张三");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        ModelAndView mv = new ModelAndView("ReadOnly/OpenWord");
        return mv;
    }
}