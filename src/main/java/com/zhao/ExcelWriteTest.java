package com.zhao;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.joda.time.DateTime;
import org.junit.Test;

import java.io.FileOutputStream;

/**
 * @Auther: zhaomo
 * @Date: 2020/11/20 22:02
 * @Description:
 */
public class ExcelWriteTest {

    String PATH  = "B:\\IDEA_WorkSpace3\\POI";

    @Test
    public void testWrite03() throws Exception {
        //1. 创建一个工作簿
        Workbook workbook = new HSSFWorkbook();
        //2. 创建一个工作表
        Sheet sheet = workbook.createSheet("赵默朋友统计表");
        //3. 创建一个行   (1,1)
        Row row1 = sheet.createRow(0);
        //4. 创建一个单元格
        Cell cell11 = row1.createCell(0);
        cell11.setCellValue("今日朋友");
        //  (1.2)
        Cell cell12 = row1.createCell(1);
        cell12.setCellValue(6666);

        //第二行  (2,1)
        Row row2 = sheet.createRow(1);
        Cell cell21 = row2.createCell(0);
        cell21.setCellValue("统计时间");
        // (2,2)
        Cell cell22 = row2.createCell(1);
        String time = new DateTime().toString("yyyy-MM-dd HH:mm:ss");
        cell22.setCellValue(time);

        //生成一张表 (IO流)   03版本就是使用xls结尾
        FileOutputStream outputStream = new FileOutputStream(PATH = "赵默朋友统计表03.xls");
        //输出
        workbook.write(outputStream);
        //关闭流
        outputStream.close();

        System.out.println(" 03生成完毕!");
    }


    @Test
    public void testWrite07() throws Exception {
        //1. 创建一个工作簿  07
        Workbook workbook = new XSSFWorkbook();
        //2. 创建一个工作表
        Sheet sheet = workbook.createSheet("赵默朋友统计表");
        //3. 创建一个行   (1,1)
        Row row1 = sheet.createRow(0);
        //4. 创建一个单元格
        Cell cell11 = row1.createCell(0);
        cell11.setCellValue("今日朋友");
        //  (1.2)
        Cell cell12 = row1.createCell(1);
        cell12.setCellValue(6666);

        //第二行  (2,1)
        Row row2 = sheet.createRow(1);
        Cell cell21 = row2.createCell(0);
        cell21.setCellValue("统计时间");
        // (2,2)
        Cell cell22 = row2.createCell(1);
        String time = new DateTime().toString("yyyy-MM-dd HH:mm:ss");
        cell22.setCellValue(time);

        //生成一张表 (IO流)   03版本就是使用xls结尾
        FileOutputStream outputStream = new FileOutputStream(PATH = "赵默朋友统计表07.xlsx");
        //输出
        workbook.write(outputStream);
        //关闭流
        outputStream.close();

        System.out.println(" 07生成完毕!");
    }

    @Test
    public void testWrite03BigData() throws Exception{
        //时间差
        long begin = System.currentTimeMillis();
        //创建一个工作簿
        HSSFWorkbook workbook = new HSSFWorkbook();
        //创建表
        Sheet sheet = workbook.createSheet();
        //写入数据
        for (int rowNum = 0; rowNum < 65536 ; rowNum++) {
            Row row = sheet.createRow(rowNum);
            for (int cellNum = 0; cellNum < 10 ; cellNum++) {
                Cell cell = row.createCell(cellNum);
                cell.setCellValue(cellNum);
            }
        }
        System.out.println("over");
        FileOutputStream outputStream = new FileOutputStream("PATH" + "testWrite03BigData.xls");
        workbook.write(outputStream);
        outputStream.close();
        long end = System.currentTimeMillis();
        System.out.println((double) (end - begin)/1000);

    }

    //特别慢
    @Test
    public void testWrite07BigData() throws Exception{
        //时间差
        long begin = System.currentTimeMillis();
        //创建一个工作簿
        XSSFWorkbook workbook = new XSSFWorkbook();
        //创建表
        Sheet sheet = workbook.createSheet();
        //写入数据
        for (int rowNum = 0; rowNum < 65537 ; rowNum++) {
            Row row = sheet.createRow(rowNum);
            for (int cellNum = 0; cellNum < 10 ; cellNum++) {
                Cell cell = row.createCell(cellNum);
                cell.setCellValue(cellNum);
            }
        }
        System.out.println("over");
        FileOutputStream outputStream = new FileOutputStream("PATH" + "testWrite07BigData.xlsx");
        workbook.write(outputStream);
        outputStream.close();
        long end = System.currentTimeMillis();
        System.out.println((double) (end - begin)/1000);

    }


    @Test
    public void testWrite07BigDataS() throws Exception{
        //时间差
        long begin = System.currentTimeMillis();
        //创建一个工作簿
        SXSSFWorkbook workbook = new SXSSFWorkbook();
        //创建表
        Sheet sheet = workbook.createSheet();
        //写入数据
        for (int rowNum = 0; rowNum < 65537 ; rowNum++) {
            Row row = sheet.createRow(rowNum);
            for (int cellNum = 0; cellNum < 10 ; cellNum++) {
                Cell cell = row.createCell(cellNum);
                cell.setCellValue(cellNum);
            }
        }
        System.out.println("over");
        FileOutputStream outputStream = new FileOutputStream("PATH" + "testWrite07BigDataS.xlsx");
        workbook.write(outputStream);
        outputStream.close();
        //清除临时文件!
        workbook.dispose();
        long end = System.currentTimeMillis();
        System.out.println((double) (end - begin)/1000);

    }
}
