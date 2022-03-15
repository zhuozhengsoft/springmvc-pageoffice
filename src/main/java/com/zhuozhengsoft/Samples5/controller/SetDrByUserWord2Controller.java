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
import java.io.FileOutputStream;
import java.util.Map;

@Controller
@RequestMapping("SetDrByUserWord2")
public class SetDrByUserWord2Controller {
    @RequestMapping("/Default")
    public ModelAndView openWord(HttpServletRequest request) throws Exception{
        ModelAndView mv = new ModelAndView("SetDrByUserWord/Default");
        return mv;

    }

    @RequestMapping("/SetDataRegionByUserName")
    public ModelAndView openWord(HttpServletRequest request, Map<String, Object> map) throws Exception{
        String userName = request.getParameter("userName");
        //***************************卓正PageOffice组件的使用********************************
        WordDocument doc = new WordDocument();
        //打开数据区域
        DataRegion d1 = doc.openDataRegion("PO_com1");
        DataRegion d2 = doc.openDataRegion("PO_com2");

        //给数据区域赋值
        d1.setValue("[word]/doc/SetDrByUserWord2/content1.doc[/word]");
        d2.setValue("[word]/doc/SetDrByUserWord2/content2.doc[/word]");

        //若要将数据区域内容存入文件中，则必须设置属性“setSubmitAsFile”值为true
        d1.setSubmitAsFile(true);
        d2.setSubmitAsFile(true);

        //根据登录用户名设置数据区域可编辑性
        //甲客户：zhangsan登录后
        if (userName.equals("zhangsan")) {
            d1.setEditing(true);
            d2.setEditing(false);
        }
        //乙客户：lisi登录后
        else {
            d2.setEditing(true);
            d1.setEditing(false);
        }

        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        //设置服务器页面
        poCtrl.setServerPage(request.getContextPath() + "/poserver.zz"); //此行必须
        poCtrl.setWriter(doc);
        //添加自定义按钮
        poCtrl.addCustomToolButton("保存", "Save", 1);
        poCtrl.addCustomToolButton("全屏/还原", "IsFullScreen", 4);
        //设置保存页
        poCtrl.setSaveDataPage("SaveData?userName=" + userName);
        poCtrl.setMenubar(false);
        //打开word
        poCtrl.webOpen("/doc/SetDrByUserWord2/test.doc", OpenModeType.docSubmitForm, userName);

        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        map.put("userName", userName);
        ModelAndView mv = new ModelAndView("SetDrByUserWord2/SetDataRegionByUserName");
        return mv;
    }

    @RequestMapping("/SaveData")
    public void SaveData(HttpServletRequest request, HttpServletResponse response) throws Exception{
        //-----------  PageOffice 服务器端编程开始  -------------------//
        com.zhuozhengsoft.pageoffice.wordreader.WordDocument doc = new com.zhuozhengsoft.pageoffice.wordreader.WordDocument(request, response);
        byte[] bytes = null;
        String filePath = "";
        if (request.getParameter("userName") != null && request.getParameter("userName").trim().equalsIgnoreCase("zhangsan")) {
            bytes = doc.openDataRegion("PO_com1").getFileBytes();
            filePath = "content1.doc";
        } else {
            bytes = doc.openDataRegion("PO_com2").getFileBytes();
            filePath = "content2.doc";
        }
        doc.close();

        filePath = request.getSession().getServletContext().getRealPath("doc/SetDrByUserWord2/") + "/" + filePath;
        FileOutputStream outputStream = new FileOutputStream(filePath);
        outputStream.write(bytes);
        outputStream.flush();
        outputStream.close();
    }
}