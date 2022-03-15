package com.zhuozhengsoft.Samples5.controller;

import com.zhuozhengsoft.pageoffice.OpenModeType;
import com.zhuozhengsoft.pageoffice.PageOfficeCtrl;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.servlet.ModelAndView;

import javax.servlet.http.HttpServletRequest;
import java.util.Map;

@Controller
@RequestMapping("WordRibbonCtrl")
public class WordRibbonCtrlController {

    @RequestMapping("/Word")
    public ModelAndView openWord(HttpServletRequest request, Map<String, Object> map) throws Exception{
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        //设置服务器页面
        poCtrl.setServerPage(request.getContextPath() + "/poserver.zz");
        //添加自定义按钮
        poCtrl.addCustomToolButton("保存", "Save", 1);
        poCtrl.getRibbonBar().setTabVisible("TabHome", true);//开始
        poCtrl.getRibbonBar().setTabVisible("TabPageLayoutWord", false);//页面布局
        poCtrl.getRibbonBar().setTabVisible("TabReferences", false);//引用
        poCtrl.getRibbonBar().setTabVisible("TabMailings", false);//邮件
        poCtrl.getRibbonBar().setTabVisible("TabReviewWord", false);//审阅
        poCtrl.getRibbonBar().setTabVisible("TabInsert", false);//插入
        poCtrl.getRibbonBar().setTabVisible("TabView", false);//视图
        poCtrl.getRibbonBar().setSharedVisible("FileSave", false);//office自带的保存按钮
        poCtrl.getRibbonBar().setGroupVisible("GroupClipboard", false);//分组剪贴板
        //打开Word文档
        poCtrl.webOpen("/doc/WordRibbonCtrl/test.doc", OpenModeType.docNormalEdit, "张三");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        ModelAndView mv = new ModelAndView("WordRibbonCtrl/Word");
        return mv;
    }
}
