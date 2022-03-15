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
@RequestMapping("OpenWord")
public class OpenWordController {

    @RequestMapping("/OpenWord")
    public ModelAndView openWord(HttpServletRequest request, Map<String, Object> map) throws Exception{
        //设置PageOffice服务器组件
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage(request.getContextPath() + "/poserver.zz"); //此行必须
        //隐藏Office工具条
        poCtrl.setOfficeToolbars(false);
        //隐藏自定义工具栏
        poCtrl.setCustomToolbar(false);
        //设置页面的显示标题
        poCtrl.setCaption("演示：最简单的以只读模式打开Word文档");
        //打开文件
        poCtrl.webOpen("/doc/OpenWord/template.doc", OpenModeType.docReadOnly, "张三");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        ModelAndView mv = new ModelAndView("OpenWord/OpenWord");
        return mv;
    }
}