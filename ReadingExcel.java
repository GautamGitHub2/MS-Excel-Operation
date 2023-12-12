package MSExcelOperation;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;

public class ReadingExcel {
    public static void main(String[] args) throws IOException {
        String excelFilePath="/Users/gautamkumar/Documents/Gautam_MacBookAir/My Documents/Study_Trainings_Interviews/Quality Assurance Study/Automation Testing/Selenium with Java Automation Testing/Selenium_with_Java_Projects_IntelliJ/Selenium/DataFiles/Test Data_MSExcel_Read_Data.xlsx";
        FileInputStream inputStream=new FileInputStream(excelFilePath);

        XSSFWorkbook workbook=new XSSFWorkbook(inputStream);
        XSSFSheet sheet=workbook.getSheet("Sheet 1");
        //XSSFSheet sheet=workbook.getSheetAt(0); // this can also be used in place of sheet name

        // Using for loop how to read the data from sheet
        /*
        int rows=sheet.getLastRowNum();
        int cols=sheet.getRow(1).getLastCellNum();

        for (int r=0;r<=rows;r++)
        {
            XSSFRow row=sheet.getRow(r);

            for (int c=0;c<cols;c++)
            {
                XSSFCell cell=row.getCell(c);

                switch (cell.getCellType())
                {
                    case STRING -> System.out.print(cell.getStringCellValue()+"   |   ");
                    case NUMERIC -> System.out.print(cell.getNumericCellValue()+"   |   ");
                    case BOOLEAN -> System.out.print(cell.getBooleanCellValue()+"   |   ");
                }
            }
            System.out.println();
        }*/

        // Using Iterator Method

       Iterator iterator= sheet.iterator();
       while (iterator.hasNext())
       {
           XSSFRow row= (XSSFRow) iterator.next();

           Iterator cellIterator=row.cellIterator();

           while (cellIterator.hasNext())
           {
               XSSFCell cell= (XSSFCell) cellIterator.next();

               switch (cell.getCellType())
               {
                   case STRING -> System.out.print(cell.getStringCellValue()+"   |   ");
                   case NUMERIC -> System.out.print(cell.getNumericCellValue()+"   |   ");
                   case BOOLEAN -> System.out.print(cell.getBooleanCellValue()+"   |   ");
               }
           }
           System.out.println();
       }
    }
}

/*

Excel Sheet Data: Excel Sheet is not available in MacBook (MacOS) I have to do whatever I want to add /enter datas in ‘Numbers’ (Spreadsheet available in MacOS) and save
then export as ‘Excel’ sheet and after exporting don’t do any changes/update in excel sheet otherwise it will again save as number (spreadsheet). Save it at the folder location in IntelliJ.

String excelFilePath="/Users/gautamkumar/Documents/Gautam_MacBookAir/My Documents/Study_Trainings_Interviews/Quality Assurance Study/Automation Testing/Selenium with Java Automation Testing/Selenium_with_Java_Projects_IntelliJ/Selenium/DataFiles/Test Data_MSExcel copy.xlsx” —> Change ‘File Name’ with new file always with the same location.

Output: --> Working as expected

Table 1   |
Country   |   Capital   |   Population   |
India   |   Delhi   |   500000.0   |
France    |   Paris   |   800000.0   |
Germany   |   Berlin   |   3200000.0   |
England   |   London   |   2300000.0   |
Belarus   |   Minsk   |   7500000.0   |
Belgium   |   Brussels   |   2200000.0   |
Denmark   |   Copenhagen   |   4300000.0   |
Ireland   |   Dublin   |   6400000.0   |

*/
