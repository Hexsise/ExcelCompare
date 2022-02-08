import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.Iterator;

public class ExcelCompare {

    /*
        Compares two given Excel files. If there are differences between cells, writes to a new Excel file
        the cell information from the first given file.
     */
    public static void compareFile(File file1, File file2) {
        try
        {
            //Blank workbook
            XSSFWorkbook workbook = new XSSFWorkbook();

            //Create a blank sheet
            XSSFSheet sheet = workbook.createSheet("Difference output");

            // Counter for row in output file
            int rowNum = 0;

            // Read in from two Excel files
            FileInputStream fileInput1 = new FileInputStream(file1);
            FileInputStream fileInput2 = new FileInputStream(file2);

            //Create Workbook instance holding reference to .xlsx file
            XSSFWorkbook workbook1 = new XSSFWorkbook(fileInput1);
            XSSFWorkbook workbook2 = new XSSFWorkbook(fileInput2);

            //Get first/desired sheet from the workbook
            XSSFSheet sheet1 = workbook1.getSheetAt(0);
            XSSFSheet sheet2 = workbook2.getSheetAt(0);

            Iterator iterator1 = sheet1.iterator();
            Iterator iterator2 = sheet2.iterator();

            while(iterator1.hasNext() || iterator2.hasNext()) {

                XSSFRow row1=(XSSFRow) iterator1.next();
                XSSFRow row2=(XSSFRow) iterator2.next();

                Iterator cellIterator1 = row1.cellIterator();
                Iterator cellIterator2 = row2.cellIterator();

                while(cellIterator1.hasNext() || cellIterator2.hasNext()) {

                    XSSFCell cell1 = (XSSFCell) cellIterator1.next();
                    XSSFCell cell2 = (XSSFCell) cellIterator2.next();
                    String formattedCell1 = "";
                    String formattedCell2 = "";

                    // determine types for each cell
                    switch(cell1.getCellType()) {

                        case STRING: formattedCell1 = ""+cell1.getStringCellValue(); break;
                        case NUMERIC: formattedCell1 = ""+cell1.getNumericCellValue(); break;
                        case BOOLEAN: formattedCell1 = ""+cell1.getBooleanCellValue(); break;
                    }
                    switch(cell2.getCellType()) {

                        case STRING: formattedCell2 = ""+cell2.getStringCellValue(); break;
                        case NUMERIC: formattedCell2 = ""+cell2.getNumericCellValue(); break;
                        case BOOLEAN: formattedCell2 = ""+cell2.getBooleanCellValue(); break;
                    }

                    // if there is a difference, write the cell value of cell1 to differences file
                    if(!formattedCell1.equals(formattedCell2)) {
                        Row row = sheet.createRow(rowNum++);
                        Cell cell = row.createCell(0);
                        cell.setCellValue(formattedCell1);
                    }
                }
            }

            try
            {
                //Write the workbook in file system
                FileOutputStream out = new FileOutputStream(new File("differences.xlsx"));
                workbook.write(out);
                out.close();
                System.out.println("differences written successfully on disk.");
            }
            catch (Exception e)
            {
                e.printStackTrace();
            }

            fileInput1.close();
            fileInput2.close();
        }
        catch (Exception e)
        {
            e.printStackTrace();
        }
    }

    public static void main(String[] args) {
        compareFile(new File("exampleData1.xlsx"), new File("exampleData2.xlsx"));
    }

}

