package com.zhuozhengsoft.Samples5.controller;

import com.zhuozhengsoft.pageoffice.OpenModeType;
import com.zhuozhengsoft.pageoffice.PageOfficeCtrl;
import com.zhuozhengsoft.pageoffice.wordwriter.WordDocument;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.servlet.ModelAndView;

import javax.servlet.http.HttpServletRequest;
import java.util.Map;

@Controller
@RequestMapping("RunMacroForDocm")
public class RunMacroForDocmController {
    @RequestMapping("/Word")
    public ModelAndView openWord(HttpServletRequest request, Map<String, Object> map) throws Exception{
        //******************************׿��PageOffice�����ʹ��*******************************
        //����PageOffice���������
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage(request.getContextPath() + "/poserver.zz"); //���б���
        //���ز˵���
        poCtrl.setMenubar(false);
        //�����Զ��幤����
        poCtrl.setCustomToolbar(false);
        //���ļ�m
        poCtrl.webOpen("/doc/RunMacroForDocm/test.docm", OpenModeType.docNormalEdit, "����");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        ModelAndView mv = new ModelAndView("RunMacroForDocm/Word");
        return mv;
    }
}
