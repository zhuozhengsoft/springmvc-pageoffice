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
import java.util.Map;

@Controller
@RequestMapping("SaveDataAndFile")
public class SaveDataAndFileController {

    @RequestMapping("/Word")
    public ModelAndView openWord(HttpServletRequest request, Map<String, Object> map) throws Exception{
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        //设置服务器页面
        poCtrl.setServerPage(request.getContextPath() + "/poserver.zz");
        WordDocument wordDoc = new WordDocument();
        //打开数据区域，openDataRegion方法的参数代表Word文档中的书签名称
        DataRegion dataRegion1 = wordDoc.openDataRegion("PO_userName");
        //设置DataRegion的可编辑性
        dataRegion1.setEditing(true);
        DataRegion dataRegion2 = wordDoc.openDataRegion("PO_deptName");
        dataRegion2.setEditing(true);
        poCtrl.setWriter(wordDoc);

        //添加自定义按钮
        poCtrl.addCustomToolButton("保存", "Save", 1);
        //设置保存数据的页面
        poCtrl.setSaveDataPage("SaveData");
        //设置保存文档的页面
        poCtrl.setSaveFilePage("SaveFile");
        //打开Word文档,当需要同时保存数据和保存文档时,OpenModeType必须是docSubmitForm模式。
        poCtrl.webOpen("/doc/SaveDataAndFile/test.doc", OpenModeType.docSubmitForm, "李斯");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        ModelAndView mv = new ModelAndView("SaveDataAndFile/Word");
        return mv;

    }

    @RequestMapping("/SaveFile")
    public void save(HttpServletRequest request, HttpServletResponse response) throws Exception{
        FileSaver fs = new FileSaver(request, response);
        fs.saveToFile(request.getSession().getServletContext().getRealPath("doc/SaveDataAndFile/")+fs.getFileName());
        fs.close();
    }
    @RequestMapping("/SaveData")
    public void SaveData(HttpServletRequest request, HttpServletResponse response) throws Exception{
        com.zhuozhengsoft.pageoffice.wordreader.WordDocument doc = new com.zhuozhengsoft.pageoffice.wordreader.WordDocument(request, response);
        //获取提交的数值
        String dataUserName = doc.openDataRegion("PO_userName").getValue();
        String dataDeptName = doc.openDataRegion("PO_deptName").getValue();
        String companyName = doc.getFormField("txtCompany");

        System.out.println(dataUserName+dataDeptName+companyName);
        /**获取到的公司名称,员工姓名,部门名称等内容可以保存到数据库,这里可以连接开发人员自己的数据库,例如连接mysql数据库。
         *Class.forName("com.mysql.jdbc.Driver");
         *Connection conn = (Connection) DriverManager.getConnection("jdbc:mysql://localhost/myDataBase", "root", "111111");
         *String sql="update user set userName='"+dataUserName+"',deptName='"+dataDeptName+"',companyName='"+companyName+"' where userId=1";
         *PreparedStatement ps = conn.prepareStatement(sql);
         *int rs = ps.executeUpdate(upsql);
         * if (rs>0) {
         *    out.println("更新成功");
         *     }
         *     else{
         *   out.println("更新失败");
         *    }
         *
         *rs.close();
         *conn.close();
         */
        doc.close();
    }
}