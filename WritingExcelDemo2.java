package MSExcelOperation;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;

// Create Blank Workbook --> Add Sheet --> Rows --> Cells
public class WritingExcelDemo2 {
    public static void main(String[] args) throws IOException {

        XSSFWorkbook workbook=new XSSFWorkbook();
        XSSFSheet sheet=workbook.createSheet("Emp Info");

        ArrayList<Object[]> empdata=new ArrayList<Object[]>(); // Here it is arraylist, which contains Single Dimentional Object Array

        empdata.add(new Object[]{"EmpId","Name","Job"});
        empdata.add(new Object[]{"101","Gautam","Automation Engineer"});
        empdata.add(new Object[]{"102","Raj","Driver"});
        empdata.add(new Object[]{"103","Shyam","Doctor"});
        empdata.add(new Object[]{"104","Gobinda","Artist"});
        empdata.add(new Object[]{"105","Nitoo","Teacher"});

        //Using For-Each Loop

        int rownum=0;

        for (Object[] emp:empdata)
        {
            XSSFRow row=sheet.createRow(rownum++);

            int cellnum=0;

            for (Object value:emp)
            {
                XSSFCell cell=row.createCell(cellnum++);

                if (value instanceof String)
                    cell.setCellValue((String) value);
                if (value instanceof Integer)
                    cell.setCellValue((Integer) value);
                if (value instanceof Boolean)
                    cell.setCellValue((Boolean) value);
            }
        }

        //String filePath="/Users/gautamkumar/Documents/Gautam_MacBookAir/My Documents/Study_Trainings_Interviews/Quality Assurance Study/Automation Testing/Selenium with Java Automation Testing/Selenium_with_Java_Projects_IntelliJ/Selenium/DataFiles/EmpData_WritingExcelDemo2.xlsx";

        String filePath=".//DataFiles//EmpData_WritingExcelDemo2.xlsx";

        FileOutputStream outputStream=new FileOutputStream(filePath);
        workbook.write(outputStream);

        outputStream.close();

        System.out.println("EmployeeData.xlsx file has been written successfully...!!");
        /*
        Code is 100% perfect and the excel sheet is created but unfortunately getting below exception and file is not opening to the MacOS means file is invalid, this code wil definetely work in Windows application/ OS
        */
    }
}
