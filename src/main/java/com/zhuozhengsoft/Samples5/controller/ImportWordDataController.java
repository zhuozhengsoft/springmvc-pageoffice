package com.zhuozhengsoft.Samples5.controller;

import com.zhuozhengsoft.pageoffice.OpenModeType;
import com.zhuozhengsoft.pageoffice.PageOfficeCtrl;
import com.zhuozhengsoft.pageoffice.wordwriter.WordDocument;
import jdk.nashorn.internal.ir.JumpToInlinedFinally;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.servlet.ModelAndView;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.util.Map;

@Controller
@RequestMapping("ImportWordData")
public class ImportWordDataController {

    @RequestMapping("/Word")
    public ModelAndView openWord(HttpServletRequest request, Map<String, Object> map) throws Exception{

        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage(request.getContextPath() + "/poserver.zz");

        poCtrl.addCustomToolButton("导入文件", "importData()", 15);
        poCtrl.addCustomToolButton("提交数据", "submitData()", 1);

        WordDocument doc = new WordDocument();
        poCtrl.setWriter(doc);

        poCtrl.setSaveDataPage("SaveData");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        ModelAndView mv = new ModelAndView("ImportWordData/Word");
        return mv;
    }
    @RequestMapping("/SaveData")
    public void SaveData(HttpServletRequest request, HttpServletResponse response) throws Exception{
        String ErrorMsg = "";
        //StringBuilder ErrorMsg = null;

        String BaseUrl = "";

        com.zhuozhengsoft.pageoffice.wordreader.WordDocument doc = new com.zhuozhengsoft.pageoffice.wordreader.WordDocument(request, response);
        String sName = doc.openDataRegion("PO_name").getValue();
        String sDept = doc.openDataRegion("PO_dept").getValue();
        String sCause = doc.openDataRegion("PO_cause").getValue();
        String sNum = doc.openDataRegion("PO_num").getValue();
        String sDate = doc.openDataRegion("PO_date").getValue();

        if (sName.equals("")) {
            ErrorMsg = ErrorMsg + "<li>申请人</li>";
        }
        if (sDept.equals("")) {
            ErrorMsg = ErrorMsg + "<li>部门名称</li>";
        }
        if (sCause.equals("")) {
            ErrorMsg = ErrorMsg + "<li>请假原因</li>";
        }
        if (sDate.equals("")) {
            ErrorMsg = ErrorMsg + "<li>日期</li>";
        }
        try {
            if (sNum != "") {
                if (Integer.parseInt(sNum) < 0) {
                    ErrorMsg = ErrorMsg + "<li>请假天数不能是负数</li>";
                }
            } else {
                ErrorMsg = ErrorMsg + "<li>请假天数</li>";
            }
        } catch (Exception Ex) {
            ErrorMsg = ErrorMsg + "<li><font color=red>注意：</font>请假天数必须是数字</li>";
        }finally {
            if (ErrorMsg == "") {
                // 您可以在此编程，保存这些数据到数据库中。
                response.setContentType("text/plain; charset=utf-8");
                ErrorMsg += "提交的数据为：<br/>";
                ErrorMsg += "姓名：" + sName + "<br/>";
                ErrorMsg += "部门：" + sDept + "<br/>";
                ErrorMsg += "原因：" + sCause + "<br/>";
                ErrorMsg += "天数：" + sNum + "<br/>";
                ErrorMsg += "日期：" + sDate + "<br/>";
                response.getWriter().write(ErrorMsg);
//            response.getWriter().println("提交的数据为：<br/>");
//            response.getWriter().println("姓名：" + sName + "<br/>");
//            response.getWriter().println("部门：" + sDept + "<br/>");
//            response.getWriter().println("原因：" + sCause + "<br/>");
//            response.getWriter().println("天数：" + sNum + "<br/>");
//            response.getWriter().println("日期：" + sDate + "<br/>");
                doc.showPage(578, 380);
            } else {
                ErrorMsg = "<div style='color:#FF0000;'>请修改以下信息：</div> "
                        + ErrorMsg;

                response.setContentType("text/plain; charset=utf-8");
                response.getWriter().write(ErrorMsg);
                doc.showPage(578, 380);
            }
            doc.close();
        }
    }
}
