package com.itheima.poi;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;

public class Poi {

    public static void main(String[] args) throws IOException {
        //创建工作簿对象
        Workbook wk=new XSSFWorkbook("E:/hello.xlsx");
        //获得工作表对象
        Sheet sht = wk.getSheetAt(0);//从0开始
        //获取一共有多少行
        //sht.getPhysicalNumberOfRows();
        //最后一行的下标
        //sht.getLastRowNum();

        //遍历工作表,获得行对象
        for (Row row : sht) {
            //遍历行对象,获取单元格对象
            for (Cell cell : row) {
                //获得单元格里面的内容
                if (cell.getCellType()==Cell.CELL_TYPE_STRING){//单元格的格式
                    System.out.print(cell.getStringCellValue()+",");
                }else if (cell.getCellType()==Cell.CELL_TYPE_NUMERIC){
                    System.out.print(cell.getNumericCellValue()+",");
                }
            }
            System.out.println();


        }
        wk.close();
    }


}
