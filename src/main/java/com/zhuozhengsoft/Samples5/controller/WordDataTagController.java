package com.zhuozhengsoft.Samples5.controller;

import com.zhuozhengsoft.pageoffice.OpenModeType;
import com.zhuozhengsoft.pageoffice.PageOfficeCtrl;
import com.zhuozhengsoft.pageoffice.wordwriter.DataTag;
import com.zhuozhengsoft.pageoffice.wordwriter.WordDocument;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.servlet.ModelAndView;

import javax.servlet.http.HttpServletRequest;
import java.awt.*;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Map;

@Controller
@RequestMapping("WordDataTag")
public class WordDataTagController {

    @RequestMapping("/DataTag")
    public ModelAndView openWord(HttpServletRequest request, Map<String, Object> map) throws Exception{
        //PageOffice组件的使用
        //设置服务器页面
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage(request.getContextPath() + "/poserver.zz");
        //定义WordDocument对象
        WordDocument doc = new WordDocument();
        //定义DataTag对象
        DataTag deptTag = doc.openDataTag("{部门名}");
        //给DataTag对象赋值
        deptTag.setValue("B部门");
        deptTag.getFont().setColor(Color.GREEN);

        DataTag userTag = doc.openDataTag("{姓名}");
        userTag.setValue("李四");
        userTag.getFont().setColor(Color.GREEN);

        DataTag dateTag = doc.openDataTag("【时间】");
        dateTag.setValue(new SimpleDateFormat("yyyy-MM-dd").format(new Date()).toString());
        dateTag.getFont().setColor(Color.BLUE);

        poCtrl.setWriter(doc);
        //打开Word文件
        poCtrl.webOpen("/doc/WordDataTag/test2.doc", OpenModeType.docNormalEdit, "张三");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        ModelAndView mv = new ModelAndView("WordDataTag/DataTag");
        return mv;
    }
}
