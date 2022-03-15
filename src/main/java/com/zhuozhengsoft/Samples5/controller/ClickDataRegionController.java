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
import java.awt.*;
import java.util.Map;

@Controller
@RequestMapping("ClickDataRegion")
public class ClickDataRegionController {

    @RequestMapping("/HTMLPage")
    public ModelAndView HTMLPage(HttpServletRequest request, Map<String, Object> map) throws Exception {
        ModelAndView mv = new ModelAndView("ClickDataRegion/HTMLPage");
        return mv;
    }


        @RequestMapping("/Default")
        public ModelAndView openWord (HttpServletRequest request, Map < String, Object > map) throws Exception {
            PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);

            WordDocument doc = new WordDocument();
            DataRegion dataReg = doc.openDataRegion("PO_deptName");
            dataReg.getShading().setBackgroundPatternColor(Color.pink);
            //dataReg.setEditing(true);
            poCtrl.setWriter(doc);

            //设置服务器页面
            poCtrl.setServerPage(request.getContextPath() + "/poserver.zz"); //此行必须
            poCtrl.addCustomToolButton("保存", "Save", 1);
            poCtrl.setJsFunction_OnWordDataRegionClick("OnWordDataRegionClick()");
            poCtrl.setOfficeToolbars(false);
            poCtrl.setCaption("为方便用户知道哪些地方可以编辑，所以设置了数据区域的背景色");
            poCtrl.setSaveFilePage("save");
            //打开word
            poCtrl.webOpen("/doc/ClickDataRegion/test.doc", OpenModeType.docSubmitForm, "李斯");
            map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
            ModelAndView mv = new ModelAndView("ClickDataRegion/Default");
            return mv;
        }

        @RequestMapping("/save")
        public void SaveData (HttpServletRequest request, HttpServletResponse response) throws Exception {
            FileSaver fs = new FileSaver(request, response);
            fs.saveToFile(request.getSession().getServletContext().getRealPath("doc/ClickDataRegion/") + "/" + fs.getFileName());
            fs.close();
        }
    }