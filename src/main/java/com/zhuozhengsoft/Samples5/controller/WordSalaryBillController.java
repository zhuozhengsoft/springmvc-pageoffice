package com.zhuozhengsoft.Samples5.controller;

import com.zhuozhengsoft.pageoffice.FileSaver;
import com.zhuozhengsoft.pageoffice.OpenModeType;
import com.zhuozhengsoft.pageoffice.PageOfficeCtrl;
import com.zhuozhengsoft.pageoffice.wordwriter.DataRegion;
import com.zhuozhengsoft.pageoffice.wordwriter.DataRegionInsertType;
import com.zhuozhengsoft.pageoffice.wordwriter.Table;
import com.zhuozhengsoft.pageoffice.wordwriter.WordDocument;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.servlet.ModelAndView;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.net.FileNameMap;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.Statement;
import java.text.NumberFormat;
import java.text.SimpleDateFormat;
import java.util.Locale;
import java.util.Map;

@Controller
@RequestMapping("WordSalaryBill")
public class WordSalaryBillController {

    @RequestMapping("/index")
    public ModelAndView index(HttpServletRequest request, Map<String, Object> map) throws Exception{
        Class.forName("org.sqlite.JDBC");
    String strUrl = "jdbc:sqlite:"
            + request.getServletContext().getRealPath("demodata/") + "/WordSalaryBill.db";
        Connection conn = DriverManager.getConnection(strUrl);
        Statement stmt = conn.createStatement();
        ResultSet rs = stmt.executeQuery("select * from Salary  order by ID");
        boolean flg = false;//标识是否有数据
        StringBuilder strHtmls = new StringBuilder();
        //SimpleDateFormat  formatDate = new SimpleDateFormat("yyyy-MM-dd");
        //DateFormat format = new SimpleDateFormat("yyyy-MM-dd");
        //NumberFormat formater = NumberFormat.getCurrencyInstance(Locale.CHINA);
        strHtmls.append("<tr  style='height:40px;padding:0; border-right:1px solid #a2c5d9; border-bottom:1px solid #a2c5d9; background:#edf8fe; font-weight:bold; color:#666;text-align:center; text-indent:0px;'>");
        strHtmls.append("<td width='5%' >选择</td>");
        strHtmls.append("<td width='10%' >员工编号</td>");
        strHtmls.append("<td width='10%' >员工姓名</td>");
        strHtmls.append("<td width='15%' >所在部门</td>");
        strHtmls.append("<td width='10%' >应发工资</td>");
        strHtmls.append("<td width='10%' >扣除金额</td>");
        strHtmls.append("<td width='10%' >实发工资</td>");
        strHtmls.append("<td width='10%' >发放日期</td>");
        strHtmls.append("<td width='20%' >操作</td>");
        strHtmls.append("</tr>");
        while (rs.next()) {
            flg = true;
            String pID = rs.getString("ID");
            strHtmls.append("<tr  style='height:40px; text-indent:10px; padding:0; border-right:1px solid #a2c5d9; border-bottom:1px solid #a2c5d9; color:#666;'>");
            strHtmls.append("<td style=' text-align:center;'><input id='check" + pID + "'  type='checkbox' /></td>");
            strHtmls.append("<td style=' text-align:left;'>" + pID + "</td>");
            strHtmls.append("<td style=' text-align:left;'>" + rs.getString("UserName").toString() + "</td>");
            strHtmls.append("<td style=' text-align:left;'>" + rs.getString("DeptName").toString() + "</td>");

            strHtmls.append("<td style=' text-align:left;'>" + rs.getString("SalTotal").toString() + "</td>");
            strHtmls.append("<td style=' text-align:left;'>" + rs.getString("SalDeduct").toString() + "</td>");
            strHtmls.append("<td style=' text-align:left;'>" + rs.getString("SalCount").toString() + "</td>");
            strHtmls.append("<td style=' text-align:center;'>" + rs.getString("DataTime") + "</td>");
            strHtmls.append("<td style=' text-align:center;'><a href='javascript:POBrowser.openWindowModeless(\"View?ID=" + pID + "\" ,\"width=1200px;height=800px;\");'>查看</a>&nbsp;&nbsp;&nbsp;&nbsp;<a href='javascript:POBrowser.openWindowModeless(\"Openfile?ID=" + pID + "\" ,\"width=1200px;height=800px;\");'>编辑</a></td>");
            strHtmls.append("</tr>");
        }

        if (!flg) {
            strHtmls.append("<tr>\r\n");
            strHtmls.append("<td width='100%' height='100' align='center'>对不起，暂时没有可以操作的数据。\r\n");
            strHtmls.append("</td></tr>\r\n");
        }
        request.setAttribute("strHtmls",strHtmls);
        ModelAndView mv = new ModelAndView("WordSalaryBill/index");
        return mv;
    }

    @RequestMapping("/View")
    public ModelAndView view(HttpServletRequest request,HttpServletResponse response, Map<String, Object> map) throws Exception{
        String err = "";
        String id = request.getParameter("ID").trim();
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        //创建WordDocment对象
        WordDocument doc = new WordDocument();
        if (id != null && id.length() > 0) {
            String strSql = "select * from Salary where id =" + id
                    + " order by ID";
            Class.forName("org.sqlite.JDBC");
            String strUrl = "jdbc:sqlite:"
                    + request.getServletContext().getRealPath("demodata/") + "\\WordSalaryBill.db";
            Connection conn = DriverManager.getConnection(strUrl);
            Statement stmt = conn.createStatement();
            ResultSet rs = stmt.executeQuery(strSql);

            //打开数据区域
            DataRegion datareg = doc.openDataRegion("PO_table");

            SimpleDateFormat formatDate = new SimpleDateFormat("yyyy-MM-dd");
            NumberFormat formater = NumberFormat
                    .getCurrencyInstance(Locale.CHINA);
            if (rs.next()) {
                //给数据区域赋值
                doc.openDataRegion("PO_ID").setValue(id);
                doc.openDataRegion("PO_UserName").setValue(rs.getString("UserName"));
                doc.openDataRegion("PO_DeptName").setValue(rs.getString("DeptName"));

                String saltotal = rs.getString("SalTotal");
                if (saltotal != null && saltotal != "") {
                    doc.openDataRegion("SalTotal").setValue(saltotal);
                } else {
                    doc.openDataRegion("SalTotal").setValue("￥0.00");
                }

                String saldeduct = rs.getString("SalDeduct");
                if (saldeduct != null && saldeduct != "") {
                    doc.openDataRegion("SalDeduct").setValue(saldeduct);
                } else {
                    doc.openDataRegion("SalDeduct").setValue("￥0.00");
                }
                String salcount = rs.getString("SalCount");
                if (salcount != null && salcount != "") {
                    doc.openDataRegion("SalCount").setValue(salcount);
                } else {
                    doc.openDataRegion("SalCount").setValue("￥0.00");
                }
                String datatime = rs.getString("DataTime");
                if (datatime != null && datatime != "") {
                    doc.openDataRegion("DataTime").setValue(datatime);
                } else {
                    doc.openDataRegion("DataTime").setValue("");
                }

            } else {
                err = "<script>alert('未获得该员工的工资信息！');location.href='index'</script>";
                response.setContentType("text/plain; charset=utf-8");
                response.getWriter().write(err);
            }
            rs.close();
            conn.close();
        } else {
            err = "<script>alert('未获得该工资信息的ID！');location.href='index'</script>";
            response.setContentType("text/plain; charset=utf-8");
            response.getWriter().write(err);
        }

        poCtrl.setWriter(doc);
        poCtrl.setServerPage(request.getContextPath() + "/poserver.zz");
        poCtrl.webOpen("/doc/WordSalaryBill/template.doc", OpenModeType.docSubmitForm, "张三");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        ModelAndView mv = new ModelAndView("WordSalaryBill/View");
        return mv;
    }
    @RequestMapping("/Openfile")
    public ModelAndView Openfile(HttpServletRequest request,HttpServletResponse response, Map<String, Object> map) throws Exception{
        String err = "";
        String id = request.getParameter("ID").trim();
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        if (id != null && id.length() > 0) {
            String strSql = "select * from Salary where id =" + id
                    + " order by ID";
            Class.forName("org.sqlite.JDBC");
            String strUrl = "jdbc:sqlite:"
                    + request.getServletContext().getRealPath("demodata/") + "\\WordSalaryBill.db";
            Connection conn = DriverManager.getConnection(strUrl);
            Statement stmt = conn.createStatement();
            ResultSet rs = stmt.executeQuery(strSql);

            //创建WordDocment对象
            WordDocument doc = new WordDocument();
            //打开数据区域
            DataRegion datareg = doc.openDataRegion("PO_table");

            SimpleDateFormat formatDate = new SimpleDateFormat("yyyy-MM-dd");
            NumberFormat formater = NumberFormat
                    .getCurrencyInstance(Locale.CHINA);

            if (rs.next()) {
                //给数据区域赋值
                doc.openDataRegion("PO_ID").setValue(id);

                //设置数据区域的可编辑性
                doc.openDataRegion("PO_UserName").setEditing(true);
                doc.openDataRegion("PO_DeptName").setEditing(true);
                doc.openDataRegion("PO_SalTotal").setEditing(true);
                doc.openDataRegion("PO_SalDeduct").setEditing(true);
                doc.openDataRegion("PO_SalCount").setEditing(true);
                doc.openDataRegion("PO_DataTime").setEditing(true);
                doc.openDataRegion("PO_UserName").setValue(rs.getString("UserName"));
                doc.openDataRegion("PO_DeptName").setValue(rs.getString("DeptName"));

                String saltotal = rs.getString("SalTotal");
                if (saltotal != null && saltotal != "") {
                    doc.openDataRegion("SalTotal").setValue(saltotal);
                    //out.print(rs.getString("SalTotal"));
                } else {
                    doc.openDataRegion("SalTotal").setValue("￥0.00");
                }

                String saldeduct = rs.getString("SalDeduct");
                if (saldeduct != null && saldeduct != "") {
                    doc.openDataRegion("SalDeduct").setValue(saldeduct);
                } else {
                    doc.openDataRegion("SalDeduct").setValue("￥0.00");
                }
                String salcount = rs.getString("SalCount");
                if (salcount != null && salcount != "") {
                    doc.openDataRegion("SalCount").setValue(salcount);
                } else {
                    doc.openDataRegion("SalCount").setValue("￥0.00");
                }
                String datatime = rs.getString("DataTime");
                if (datatime != null && datatime != "") {
                    doc.openDataRegion("DataTime").setValue(datatime);
                } else {
                    doc.openDataRegion("DataTime").setValue("");
                }

            } else {
                err = "<script>alert('未获得该员工的工资信息！');location.href='index'</script>";
                response.setContentType("text/plain; charset=utf-8");
                response.getWriter().write(err);
            }
            rs.close();
            conn.close();

            poCtrl.addCustomToolButton("保存", "Save()", 1);
            poCtrl.setSaveDataPage("SaveData?ID=" + id);
            poCtrl.setWriter(doc);
        } else {
            err = "<script>alert('未获得该工资信息的ID！');location.href='index'</script>";
            response.setContentType("text/plain; charset=utf-8");
            response.getWriter().write(err);
        }
        poCtrl.setServerPage(request.getContextPath() + "/poserver.zz");
        poCtrl.webOpen("/doc/WordSalaryBill/template.doc", OpenModeType.docSubmitForm, "someBody");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        ModelAndView mv = new ModelAndView("WordSalaryBill/Openfile");
        return mv;
    }

    @RequestMapping("/Compose")
    public ModelAndView taoHong(HttpServletRequest request, Map<String, Object> map) throws Exception {
        if (request.getParameter("ids").equals(null)
                || request.getParameter("ids").equals("")) {
            ModelAndView mv = new ModelAndView("WordSalaryBill/index");
            return mv;
        }
        String idlist = request.getParameter("ids").trim();
        //从数据库中读取数据
        String strSql = "select * from Salary where ID in(" + idlist
                + ") order by ID";

        Class.forName("org.sqlite.JDBC");
        String strUrl = "jdbc:sqlite:"
                + request.getServletContext().getRealPath("demodata/") + "\\WordSalaryBill.db";
        Connection conn = DriverManager.getConnection(strUrl);
        Statement stmt = conn.createStatement();
        ResultSet rs = stmt.executeQuery(strSql);

        SimpleDateFormat formatDate = new SimpleDateFormat("yyyy-MM-dd");
        NumberFormat formater = NumberFormat.getCurrencyInstance(Locale.CHINA);

        WordDocument doc = new WordDocument();
        DataRegion data = null;
        Table table = null;
        int i = 0;
        while (rs.next()) {
            data = doc.createDataRegion("reg" + i,
                    DataRegionInsertType.Before, "[End]");
            data.setValue("[word]/doc/WordSalaryBill/template.doc[/word]");

            table = data.openTable(1);
            table.openCellRC(2, 1).setValue(rs.getString("ID"));

            //给单元格赋值
            table.openCellRC(2, 2).setValue(rs.getString("UserName"));
            table.openCellRC(2, 3).setValue(rs.getString("DeptName"));

            String saltotal = rs.getString("SalTotal");
            if (saltotal != null && saltotal != "") {
                table.openCellRC(2, 4).setValue(saltotal);
                //out.print(rs.getString("SalTotal"));
            } else {
                table.openCellRC(2, 4).setValue("￥0.00");
            }

            String saldeduct = rs.getString("SalDeduct");
            if (saldeduct != null && saldeduct != "") {
                table.openCellRC(2, 5).setValue(saldeduct);
            } else {
                table.openCellRC(2, 5).setValue("￥0.00");
            }
            String salcount = rs.getString("SalCount");
            if (salcount != null && salcount != "") {
                table.openCellRC(2, 6).setValue(salcount);
            } else {
                table.openCellRC(2, 6).setValue("￥0.00");
            }
            String datatime = rs.getString("DataTime");
            if (datatime != null && datatime != "") {
                table.openCellRC(2, 7).setValue(datatime);
            } else {
                table.openCellRC(2, 7).setValue("");
            }
            i++;
        }
        conn.close();
        // 设置PageOffice组件服务页面
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setWriter(doc);
        poCtrl.setServerPage(request.getContextPath() + "/poserver.zz");
        poCtrl.setCaption("生成工资条");
        poCtrl.setCustomToolbar(false);
        //打开word
        poCtrl.webOpen("/doc/WordSalaryBill/test.doc", OpenModeType.docNormalEdit, "李斯");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        ModelAndView mv = new ModelAndView("WordSalaryBill/Compose");
        return mv;
    }


    @RequestMapping("/SaveData")
    public void SaveData(HttpServletRequest request, HttpServletResponse response) throws Exception {
        String err = "";
        String id = request.getParameter("ID");
        if (id != null && id.length() > 0) {
            String strSql = "select * from Salary where id =" + id
                    + " order by ID";
            Class.forName("org.sqlite.JDBC");
            String strUrl = "jdbc:sqlite:"
                    + request.getServletContext().getRealPath("demodata/") + "\\WordSalaryBill.db";
            Connection conn = DriverManager.getConnection(strUrl);
            Statement stmt = conn.createStatement();

            String userName = "", deptName = "", salTotoal = "0", salDeduct = "0", salCount = "0", dateTime = "";
            //-----------  PageOffice 服务器端编程开始  -------------------//
            com.zhuozhengsoft.pageoffice.wordreader.WordDocument doc = new com.zhuozhengsoft.pageoffice.wordreader.WordDocument(request, response);
            userName = doc.openDataRegion("PO_UserName").getValue();
            deptName = doc.openDataRegion("PO_DeptName").getValue();
            //将格式化的数据转化为String存到数据库
            salTotoal = doc.openDataRegion("PO_SalTotal").getValue();
            salDeduct = doc.openDataRegion("PO_SalDeduct").getValue();
            salCount = doc.openDataRegion("PO_SalCount").getValue();
            dateTime = doc.openDataRegion("PO_DataTime").getValue();

            String sql = "UPDATE Salary SET UserName='" + userName
                    + "',DeptName='" + deptName + "',SalTotal='" + salTotoal
                    + "',SalDeduct='" + salDeduct + "',SalCount='" + salCount
                    + "',DataTime='" + dateTime + "' WHERE ID=" + id;

            int count = stmt.executeUpdate(sql);
            if (count > 0) {
                //向客户端控件返回以上代码执行成功的消息。
                doc.setCustomSaveResult("ok");
            }
            doc.close();
            conn.close();
        } else {
            err = "<script>alert('未获得文件的ID，保存失败！');location.href='index'</script>";
            response.setContentType("text/plain; charset=utf-8");
            response.getWriter().write(err);
        }
    }
}