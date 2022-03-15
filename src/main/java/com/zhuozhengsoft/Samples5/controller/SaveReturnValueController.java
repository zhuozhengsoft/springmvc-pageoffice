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
@RequestMapping("SaveReturnValue")
public class SaveReturnValueController {

    @RequestMapping("/SaveReturnValue")
    public ModelAndView openWord(HttpServletRequest request, Map<String, Object> map) throws Exception{
        //设置PageOffice服务器组件
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage(request.getContextPath() + "/poserver.zz"); //此行必须
        //设置PageOfficeCtrl控件的标题栏名称
        poCtrl.setCaption("Word文档保存后获取返回值");
        //隐藏菜单栏
        poCtrl.setMenubar(false);
        //添加自定义按钮
        poCtrl.addCustomToolButton("保存", "Save()", 1);
        poCtrl.addCustomToolButton("关闭", "Close()", 21);
        //设置保存页面
        poCtrl.setSaveFilePage("save");//设置处理文件保存的请求方法
        //打开word
        poCtrl.webOpen("/doc/SaveReturnValue/test.doc", OpenModeType.docNormalEdit, "李斯");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        ModelAndView mv = new ModelAndView("SaveReturnValue/SaveReturnValue");
        return mv;
    }

    @RequestMapping("/save")
    public void save(HttpServletRequest request, HttpServletResponse response) throws Exception{
        FileSaver fs = new FileSaver(request, response);
        fs.saveToFile(request.getSession().getServletContext().getRealPath("doc/SaveReturnValue/")+fs.getFileName());
        fs.setCustomSaveResult("ok");//设置保存结果
        fs.close();
    }
}