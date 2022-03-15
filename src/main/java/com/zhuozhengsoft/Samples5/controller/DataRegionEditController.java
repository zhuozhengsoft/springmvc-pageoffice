package com.zhuozhengsoft.Samples5.controller;

import com.zhuozhengsoft.pageoffice.*;
import com.zhuozhengsoft.pageoffice.wordwriter.DataRegion;
import com.zhuozhengsoft.pageoffice.wordwriter.WordDocument;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.servlet.ModelAndView;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.util.Map;

@Controller
@RequestMapping("DataRegionEdit")
public class DataRegionEditController {

    @RequestMapping("/Default")
    public ModelAndView openWord(HttpServletRequest request, Map<String, Object> map) throws Exception{
        WordDocument doc = new WordDocument();
        doc.getTemplate().defineDataRegion("Name", "[ 姓名 ]");
        doc.getTemplate().defineDataRegion("Address", "[ 地址 ]");
        doc.getTemplate().defineDataRegion("Tel", "[ 电话 ]");
        doc.getTemplate().defineDataRegion("Phone", "[ 手机 ]");
        doc.getTemplate().defineDataRegion("Sex", "[ 性别 ]");
        doc.getTemplate().defineDataRegion("Age", "[ 年龄 ]");
        doc.getTemplate().defineDataRegion("Email", "[ 邮箱 ]");
        doc.getTemplate().defineDataRegion("QQNo", "[ QQ号 ]");
        doc.getTemplate().defineDataRegion("MSNNo", "[ MSN号 ]");

        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.addCustomToolButton("保存", "Save()", 1);
        poCtrl.addCustomToolButton("定义数据区域", "ShowDefineDataRegions()", 3);
        poCtrl.setServerPage(request.getContextPath() + "/poserver.zz");
        poCtrl.setSaveFilePage("save");
        poCtrl.setTheme(ThemeType.Office2007);
        poCtrl.setBorderStyle(BorderStyleType.BorderThin);
        poCtrl.setWriter(doc);
        //打开word
        poCtrl.webOpen("/doc/DataRegionEdit/test.doc", OpenModeType.docNormalEdit, "李斯");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        ModelAndView mv = new ModelAndView("DataRegionEdit/Default");
        return mv;
    }

    @RequestMapping("/DataRegionDlg")
    public ModelAndView DataRegionDlg(HttpServletRequest request, Map<String, Object> map) throws Exception{
        ModelAndView mv = new ModelAndView("DataRegionEdit/DataRegionDlg");
        return mv;
    }

    @RequestMapping("/save")
    public void save(HttpServletRequest request, HttpServletResponse response) throws Exception{
        FileSaver fs = new FileSaver(request, response);
        fs.saveToFile(request.getSession().getServletContext().getRealPath("doc/DataRegionEdit/")+fs.getFileName());
        fs.close();
    }
}