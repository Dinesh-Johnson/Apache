package com.xowrkz.excel;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;

public class IceCreamRunner {
    public static void main(String[] args) throws IOException {


        String excel = "C:\\Users\\Personal\\Documents\\icecream.xlsx";

        try(FileInputStream fileInputStream = new FileInputStream(excel);
        XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream)){
        XSSFSheet sheet = workbook.cloneSheet(0);

        Iterator iterator = sheet.iterator();

        while (iterator.hasNext()){

            XSSFRow row =(XSSFRow) iterator.next();
            Iterator cellIterator = row.cellIterator();

            while (cellIterator.hasNext()){

                XSSFCell cell =(XSSFCell) cellIterator.next();
                switch (cell.getCellType()){
                    case STRING:
                        System.out.print(cell.getStringCellValue()+" "); break;
                    case NUMERIC:
                        System.out.print(cell.getNumericCellValue()+" ");break;
                    case BOOLEAN:
                        System.out.print(cell.getBooleanCellValue()+" ");break;

                }
                System.out.print("  |  ");
            }
            System.out.println();
        }
        }catch (FileNotFoundException f){
            System.out.println("File not found: " + excel);
        } catch (IOException e) {
            System.out.println("Error reading the Excel file: " + e.getMessage());
        }
    }
}
