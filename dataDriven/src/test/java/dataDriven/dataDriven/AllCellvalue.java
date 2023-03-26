package dataDriven.dataDriven;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class AllCellvalue {

    public static void main(String[] args) throws IOException {

        String filePath = "C://Users//Dell//Documents//TestData.xlsx";
        FileInputStream inputStream = new FileInputStream(new File(filePath));

        Workbook workbook = WorkbookFactory.create(inputStream);
        int noofSheets=workbook.getNumberOfSheets();
        
        for(int i=0;i<noofSheets;i++) {
        Sheet sheet = workbook.getSheetAt(i);
        
        // Iterate through each row
        for (Row row : sheet) {

            // Iterate through each cell in the row
            for (Cell cell : row) {

                // Check if cell has data
                if (cell.getCellType() != CellType.BLANK) {

                    // Print the cell value
                    System.out.print(cell.toString() + "\t");
                }
            }

            // Move to the next row
            System.out.println();
        }

        // Close the input stream
        inputStream.close();
    }
    }
}