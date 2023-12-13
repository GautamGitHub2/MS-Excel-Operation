package MSExcelOperation;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;

// Create Blank Workbook --> Add Sheet --> Rows --> Cells
public class WritingExcelDemo1 {
    public static void main(String[] args) throws IOException {

        XSSFWorkbook workbook=new XSSFWorkbook();
        XSSFSheet sheet=workbook.createSheet("Emp Info");

        Object empdata[][]={
                            {"EmpID","Name","Job","Contact No.","Address"},
                            {101,"Gautam","Automation Engineer",565456,"Bengaluru"},
                            {102,"Nitoo","House Wife",767876,"Delhi"},
                            {103,"Purushottam Raj","Software Developer",7656789,"Mumbai"},
                            {104,"Manu","Civil Engineer",9989878,"Kolkata"},
                            {105,"Appu","Advocate",6565456,"Patna"},
                            {106,"Raj","Actor",989878,"Ranchi"},
                            {107,"Ram","Criketer",9876545,"Goa"},

                            };//Here object data can hold (heterogenous) multiple types of data

        //Using For Loop
        /*
        int rows=empdata.length;
        int cols=empdata[0].length;

        System.out.println(rows);//8
        System.out.println(cols);//5

        for (int r=0;r<rows;r++)
        {
            //Create Row in Excel sheet (Just before going to the cells/columns of the row)
            XSSFRow row=sheet.createRow(r);

            for (int c=0;c<cols;c++)
            {
                // Now before writing data into the cell, i have to create cells/cloumns for that rows
                XSSFCell cell=row.createCell(c);

                //now Cell is created, now i have create data into that cells by taking data from the 2D Array that i have created above "Object empdata[][]"

                //Read those data and update in the cell
                Object value=empdata[r][c];

                //now take the data from 2D array and exactly copy to the excel sheet
                // before setting values to the excel sheet we have to check that the values are String, Integer or Boolean and accordingly i have to types cast

                if(value instanceof String)
                    cell.setCellValue((String) value);

                if (value instanceof Integer)
                    cell.setCellValue((Integer) value);

                if(value instanceof Boolean)
                    cell.setCellValue((Boolean) value);
            }
        }*/

        //Using For-Each Loop

        int rowCount=0;

        for (Object emp[]:empdata)
        {
            XSSFRow row= sheet.createRow(rowCount++);
            int columnCount=0;
            for (Object value:emp)
            {
                XSSFCell cell=row.createCell(columnCount++);
                        if(value instanceof String)
                            cell.setCellValue((String) value);
                        if(value instanceof Integer)
                            cell.setCellValue((Integer) value);
                        if (value instanceof Boolean)
                            cell.setCellValue((Boolean) value);
            }
        }
        String filePath="/Users/gautamkumar/Documents/Gautam_MacBookAir/My Documents/Study_Trainings_Interviews/Quality Assurance Study/Automation Testing/Selenium with Java Automation Testing/Selenium_with_Java_Projects_IntelliJ/Selenium/DataFiles/EmployeeData.xlsx";
        FileOutputStream outputStream=new FileOutputStream(filePath);
        workbook.write(outputStream);

        outputStream.close();

        System.out.println("EmployeeData.xlsx file has been written successfully...!!");
        /*
        Code is 100% perfect and the excel sheet is created but unfortunately getting below exception and file is not opening to the MacOS means file is invalid, this code wil definetely work in Windows application/ OS
        */
    }
}
