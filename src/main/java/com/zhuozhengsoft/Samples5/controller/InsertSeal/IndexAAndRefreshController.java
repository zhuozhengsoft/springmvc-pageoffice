package com.zhuozhengsoft.Samples5.controller.InsertSeal;

import com.zhuozhengsoft.pageoffice.FileSaver;
import com.zhuozhengsoft.pageoffice.OpenModeType;
import com.zhuozhengsoft.pageoffice.PageOfficeCtrl;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.servlet.ModelAndView;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.util.Map;

@Controller
@RequestMapping("InsertSeal")
public class IndexAAndRefreshController {
    @RequestMapping("/index")
    public ModelAndView index(HttpServletRequest request, Map<String, Object> map) throws Exception{
        ModelAndView mv = new ModelAndView("InsertSeal/index");
        return mv;
    }
    @RequestMapping("/refresh")
    public ModelAndView refresh(HttpServletRequest request, Map<String, Object> map) throws Exception{
        String rootPath = request.getSession().getServletContext().getRealPath("doc/InsertSeal") + "/";
        //System.out.println(rootPath);
        copyFile(rootPath + "Word/AddSeal1/word1_bak.doc", rootPath + "Word/AddSeal1/word1.doc");
        copyFile(rootPath + "Word/AddSeal2/word2_bak.doc", rootPath + "Word/AddSeal2/word2.doc");
        copyFile(rootPath + "Word/AddSeal3/word3_bak.doc", rootPath + "Word/AddSeal3/word3.doc");
        copyFile(rootPath + "Word/AddSeal4/word4_bak.doc", rootPath + "Word/AddSeal4/word4.doc");
        copyFile(rootPath + "Word/AddSeal5/test_bak.doc", rootPath + "Word/AddSeal5/test.doc");
        copyFile(rootPath + "Word/AddSeal6/test_bak.doc", rootPath + "Word/AddSeal6/test.doc");
        copyFile(rootPath + "Word/AddSeal7/test_bak.doc", rootPath + "Word/AddSeal7/test.doc");
        copyFile(rootPath + "Word/AddSeal8/word8_bak.doc", rootPath + "Word/AddSeal8/word8.doc");
        copyFile(rootPath + "Word/AddSeal9/word9_bak.doc", rootPath + "Word/AddSeal9/word9.doc");

        copyFile(rootPath + "Word/AddSign1/word1_bak.doc", rootPath + "Word/AddSign1/word1.doc");
        copyFile(rootPath + "Word/AddSign2/test_bak.doc", rootPath + "Word/AddSign2/test.doc");
        copyFile(rootPath + "Word/AddSign3/test_bak.doc", rootPath + "Word/AddSign3/test.doc");
        copyFile(rootPath + "Word/AddSign4/test_bak.doc", rootPath + "Word/AddSign4/test.doc");
        copyFile(rootPath + "Word/AddSign5/word2_bak.doc", rootPath + "Word/AddSign5/word2.doc");

        copyFile(rootPath + "PDF/AddSeal1/test_bak.pdf", rootPath + "PDF/AddSeal1/test.pdf");
        copyFile(rootPath + "PDF/AddSign1/test_bak.pdf", rootPath + "PDF/AddSign1/test.pdf");


        copyFile(rootPath + "Word/BatchAddSeal/bak/test1.doc", rootPath + "Word/BatchAddSeal/test1.doc");
        copyFile(rootPath + "Word/BatchAddSeal/bak/test2.doc", rootPath + "Word/BatchAddSeal/test2.doc");
        copyFile(rootPath + "Word/BatchAddSeal/bak/test3.doc", rootPath + "Word/BatchAddSeal/test3.doc");
        copyFile(rootPath + "Word/BatchAddSeal/bak/test4.doc", rootPath + "Word/BatchAddSeal/test4.doc");

        copyFile(rootPath + "Excel/AddSeal1/test_bak.xls", rootPath + "Excel/AddSeal1/test.xls");
        copyFile(rootPath + "Excel/AddSeal2/test_bak.xls", rootPath + "Excel/AddSeal2/test.xls");
        copyFile(rootPath + "Excel/AddSeal3/test_bak.xls", rootPath + "Excel/AddSeal3/test.xls");
        copyFile(rootPath + "Excel/AddSeal4/test_bak.xls", rootPath + "Excel/AddSeal4/test.xls");
        copyFile(rootPath + "Excel/AddSeal5/test_bak.xls", rootPath + "Excel/AddSeal5/test.xls");

        copyFile(rootPath + "Excel/AddSign1/test_bak.xls", rootPath + "Excel/AddSign1/test.xls");
        copyFile(rootPath + "Excel/AddSign2/test_bak.xls", rootPath + "Excel/AddSign2/test.xls");
        copyFile(rootPath + "Excel/AddSign3/test_bak.xls", rootPath + "Excel/AddSign3/test.xls");
        ModelAndView mv = new ModelAndView("InsertSeal/refresh");
        return mv;
    }
    //拷贝文件
    public void copyFile(String oldPath, String newPath) {
        try {
            int bytesum = 0;
            int byteread = 0;
            File oldfile = new File(oldPath);
            if (oldfile.exists()) { //文件存在时
                InputStream inStream = new FileInputStream(oldPath); //读入原文件
                FileOutputStream fs = new FileOutputStream(newPath);
                byte[] buffer = new byte[1444];
                int length;
                while ((byteread = inStream.read(buffer)) != -1) {
                    bytesum += byteread; //字节数 文件大小
                    //System.out.println(bytesum);
                    fs.write(buffer, 0, byteread);
                }
                fs.close();
                inStream.close();
            }
        } catch (Exception e) {
            System.out.println("复制单个文件操作出错");
            e.printStackTrace();
        }
    }
}
