package com.file.upordownfile.controller;

import cn.hutool.core.io.IoUtil;
import cn.hutool.poi.excel.ExcelUtil;
import cn.hutool.poi.excel.ExcelWriter;
import cn.hutool.poi.excel.StyleSet;
import cn.hutool.poi.excel.style.StyleUtil;
import com.file.upordownfile.User;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.web.bind.annotation.*;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.OutputStream;
import java.net.URLEncoder;
import java.util.*;

import static javafx.scene.input.KeyCode.*;

@RestController
@RequestMapping("/file")
public class DownloadFile {


    @PostMapping(value = "/downfile")
    public void downFile(
            @RequestParam(name = "names") String names,
            HttpServletRequest request, HttpServletResponse response) throws Exception {

        List<Map<String, Object>> list = new ArrayList<>();
        for (int i = 1; i < 4; i++) {
            Map<String, Object> map = new HashMap<>();
            map.put("index",i);
            map.put("name","yf"+i);
            map.put("firstPay1",100+i);
            map.put("firstRate1",i/100);
            map.put("firstDate1","2010-01-0"+i);

            map.put("firstPay"+2,100+i);
            map.put("firstRate"+2,i/100);
            map.put("firstDate"+2,"2010-01-0"+i);

            map.put("firstPay"+3,100+i);
            map.put("firstRate"+3,i/100);
            map.put("firstDate"+3,"2010-01-0"+i);

            list.add(map);
        }

        ExcelWriter writer = ExcelUtil.getWriter();

        writer.addHeaderAlias("index","序号");
        writer.addHeaderAlias("name","姓名");

        writer.addHeaderAlias("firstPay1","支付1");
        writer.addHeaderAlias("firstRate1","支付汇率1");
        writer.addHeaderAlias("firstDate1","支付日期1");

        writer.addHeaderAlias("firstPay2","支付2");
        writer.addHeaderAlias("firstRate2","支付汇率2");
        writer.addHeaderAlias("firstDate2","支付日期2");

        writer.addHeaderAlias("firstPay3","支付3");
        writer.addHeaderAlias("firstRate3","支付汇率3");
        writer.addHeaderAlias("firstDate3","支付日期3");

        /**
         * 合并单元格 多级
         * 行和 列都是从 0 开始 切记
         * passRows  0开始算起 避免数据占用表头
         */

        writer.merge(10,"表头");
        writer.merge(1,2,0,0,"序号",true);
        writer.merge(1,2,1,1,"姓名",true);

        writer.merge(1,1,2,4,"第一次",true);
        writer.merge(1,1,5,7,"第二次",true);
        writer.merge(1,1,8,10,"第三次",true);

        CellStyle writerCellStyle = writer.getCellStyle();

        Cell cell = writer.getCell(0, 0);


        writerCellStyle.setFillForegroundColor((short) 011);

        cell.setCellStyle(writerCellStyle);

//        Workbook workbook = writer.getWorkbook();

//        writer.writeCellValue(3,4,"nidayede");

//        StyleSet styleSet = new StyleSet(workbook);

//        styleSet.setBackgroundColor(IndexedColors.RED,false);

//        writer.setStyleSet(styleSet);


        writer.passRows(1);
        String fileName = "qqqqq";
        response.setContentType("application/vnd.ms-excel;charset=utf-8");
        response.setHeader("Content-Disposition", "attachment;filename=" + fileName + ".xls");

        OutputStream os = response.getOutputStream();
        writer.write(list, true);
        writer.flush(os);
        writer.close();
    }


    @PostMapping(value = "/export")
    public void export(
            @RequestParam(name = "names") String names,
            HttpServletRequest request, HttpServletResponse response) throws Exception {

        List<Map<String, Object>> list = new ArrayList<>();

        List<Map<String, Object>> list1 = new ArrayList<>();
        for (int i = 0; i <10 ; i++) {
            Map<String,Object> map = new HashMap<>();
            map.put("index",i);
            map.put("pc","A"+i);
            map.put("sj","200"+1);
            map.put("code",(int)Math.ceil(Math.random()*2+1));
            list.add(map);
        }
        ExcelWriter writer = ExcelUtil.getWriter();

        writer.addHeaderAlias("index","序号");
        writer.addHeaderAlias("name","姓名");

        writer.addHeaderAlias("firstPay1","支付1");
        writer.addHeaderAlias("firstRate1","支付汇率1");
        writer.addHeaderAlias("firstDate1","支付日期1");

        writer.addHeaderAlias("firstPay2","支付2");
        writer.addHeaderAlias("firstRate2","支付汇率2");
        writer.addHeaderAlias("firstDate2","支付日期2");

        writer.addHeaderAlias("firstPay3","支付3");
        writer.addHeaderAlias("firstRate3","支付汇率3");
        writer.addHeaderAlias("firstDate3","支付日期3");

        Workbook workbook = writer.getWorkbook();

        StyleSet styleSet = new StyleSet(workbook);
        styleSet.setBackgroundColor(IndexedColors.RED,false);

        writer.setStyleSet(styleSet);

        writer.passRows(1);
        String fileName = "qqqqq";
        response.setContentType("application/vnd.ms-excel;charset=utf-8");
        response.setHeader("Content-Disposition", "attachment;filename=" + fileName + ".xls");

        OutputStream os = response.getOutputStream();
        writer.write(list, true);
        writer.flush(os);
        writer.close();
    }





    @PostMapping(value = "/hello")
    public String hello() {
        return "aaaaaaaaaa";
    }

    public static void main(String[] args) {

        System.out.println((int)Math.ceil(Math.random()*2+2)  );
        System.out.println((int)Math.ceil(Math.random()*2+2)  );
        System.out.println((int)Math.ceil(Math.random()*2+2)  );
        System.out.println((int)Math.ceil(Math.random()*2+2)  );



    }



}
