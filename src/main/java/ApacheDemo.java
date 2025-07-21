import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

public class ApacheDemo {
    public static void main(String[] args) throws IOException {

        String excelfile = "C:\\Users\\Personal\\Documents\\SportsPerson.xlsx";

        FileInputStream inputStream = new FileInputStream(excelfile);

        XSSFWorkbook workbook = new XSSFWorkbook(inputStream);

        //XSSFSheet sheet=workbook.getSheet("Sheet1");
        XSSFSheet sheet=workbook.getSheetAt(0);

        //for LOOP
        int rows = sheet.getLastRowNum();
        int col= sheet.getRow(1).getLastCellNum();

        for (int r = 0; r <= rows; r++) {       //Rows

            XSSFRow row = sheet.getRow(r);
            if (row == null)
                continue;
            for (int c = 0; c < col; c++) {     //Columns

                XSSFCell cell=row.getCell(c);
                if (cell == null) {
                    continue;
                }

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
        workbook.close();
        inputStream.close();
    }

}
