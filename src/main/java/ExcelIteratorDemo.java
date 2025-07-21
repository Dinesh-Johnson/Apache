import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;

public class ExcelIteratorDemo {
    public static void main(String[] args) throws IOException {

        String excel = "C:\\Users\\Personal\\Documents\\SportsPerson.xlsx"; //path

        FileInputStream inputStream = new FileInputStream(excel);  //open the file
        XSSFWorkbook workbook = new XSSFWorkbook(inputStream);    //get the workbook
        XSSFSheet sheet = workbook.getSheetAt(0);           //get the Sheet

        Iterator iterator = sheet.iterator();

        while (iterator.hasNext()){

           XSSFRow row=(XSSFRow) iterator.next();
           Iterator cellIterator=row.cellIterator();

           while (cellIterator.hasNext()){
               XSSFCell cell = (XSSFCell) cellIterator.next();
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

    }
}
