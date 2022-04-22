package com.zhao;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.Test;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

/**
 * @Auther: zhaomo
 * @Date: 2020/11/22 22:20
 * @Description:
 */
public class ExcelReadTest {
    String PATH  = "B:\\IDEA_WorkSpace3\\POI\\";

    @Test
    public void testRead03() throws Exception {
        //获取文件流
        FileInputStream inputStream = new FileInputStream(PATH + "赵默朋友统计表03.xls");
        //1. 创建一个工作簿   使用Excel能操作的这边都可以操作
        Workbook workbook = new HSSFWorkbook(inputStream);
        //2.得到表
        Sheet sheet = workbook.getSheetAt(0);
        //3.得到行
        Row row = sheet.getRow(0);
        //4.得到列
        Cell cell = row.getCell(0);
        //关闭流
        inputStream.close();
        //读取值的时候要注意值类型
        System.out.println(cell.getStringCellValue());
    }

    public void testCellType(){
        //获取文件流
        FileInputStream inputStream = null;
        try {
            inputStream = new FileInputStream(PATH + "赵默朋友统计表03.xls");
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }
        //1. 创建一个工作簿   使用Excel能操作的这边都可以操作
        try {
            Workbook workbook = new HSSFWorkbook(inputStream);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
