package com.baizhi;


import org.apache.poi.hssf.usermodel.*;
import org.junit.Test;
import org.junit.runner.RunWith;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.test.context.junit4.SpringRunner;

import java.io.File;
import java.io.IOException;
import java.util.Date;

@SpringBootTest(classes = AoiApplication.class)
@RunWith(SpringRunner.class)
public class AoiApplicationTests {

    @Test
    public void contextLoads() {
        //创建Excel工作簿对象
        HSSFWorkbook workbook = new HSSFWorkbook();
        //处理日期格式
        HSSFCellStyle cellStyle = workbook.createCellStyle();//样式对象
        HSSFDataFormat dataFormat = workbook.createDataFormat();//日期格式
        cellStyle.setDataFormat(dataFormat.getFormat("yyyy年MM月dd日"));//设置日期格式


        //创建工作表
        HSSFSheet sheet = workbook.createSheet();
        //创建标题行
        HSSFRow row = sheet.createRow(0);
        String[] title = {"姓名", "性别", "年龄", "生日"};
        //创建单元格对象并在单元格中放入值
        for (int i = 0; i < title.length; i++) {
            HSSFCell cell = row.createCell(i);
            cell.setCellValue(title[i]);
        }


        //处理数据行
        for (int i = 1; i < 11; i++) {
            HSSFRow row1 = sheet.createRow(i);
            row1.createCell(0).setCellValue("张三" + i);
            row1.createCell(1).setCellValue("男");
            row1.createCell(2).setCellValue("2" + i);
            //设置出生年月格式
            HSSFCell cell = row1.createCell(3);
            cell.setCellValue(new Date());
            cell.setCellStyle(cellStyle);
        }


        try {
            workbook.write(new File("D:/用户.xls"));
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            try {
                workbook.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

}
