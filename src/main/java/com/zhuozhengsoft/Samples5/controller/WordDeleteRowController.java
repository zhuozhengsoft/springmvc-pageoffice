package com.zhuozhengsoft.Samples5.controller;

import com.zhuozhengsoft.pageoffice.OpenModeType;
import com.zhuozhengsoft.pageoffice.PageOfficeCtrl;
import com.zhuozhengsoft.pageoffice.wordwriter.Cell;
import com.zhuozhengsoft.pageoffice.wordwriter.Table;
import com.zhuozhengsoft.pageoffice.wordwriter.WordDocument;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.servlet.ModelAndView;

import javax.servlet.http.HttpServletRequest;
import java.util.Map;

@Controller
@RequestMapping("WordDeleteRow")
public class WordDeleteRowController {

    @RequestMapping("/WordTable")
    public ModelAndView openWord(HttpServletRequest request, Map<String, Object> map) throws Exception{
        String FilePath = request.getContextPath() + "/WordDeleteRow/doc";
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage(request.getContextPath() + "/poserver.zz");
        WordDocument doc = new WordDocument();
        Table table1 = doc.openDataRegion("PO_table").openTable(1);
        Cell cell = table1.openCellRC(2, 1);
        //删除坐标为(2,1)的单元格所在行
        table1.removeRowAt(cell);
        poCtrl.setCustomToolbar(false);
        poCtrl.setWriter(doc);
        //打开Word文档
        poCtrl.webOpen("/doc/WordDeleteRow/test_table.doc", OpenModeType.docNormalEdit, "张三");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        map.put("FilePath", FilePath);
        ModelAndView mv = new ModelAndView("WordDeleteRow/WordTable");
        return mv;
    }
}
