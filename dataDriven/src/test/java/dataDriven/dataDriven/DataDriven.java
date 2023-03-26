package dataDriven.dataDriven;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
public class DataDriven {

	public static void main(String[] args) throws IOException {
		FileInputStream fis=new FileInputStream("C://Users//Dell//Documents//TestData.xlsx");//Excel sheet path 
		XSSFWorkbook workbook=new XSSFWorkbook(fis);// to get control over Excel sheet 
		
	int sheets=	workbook.getNumberOfSheets();//no.of sheets in the excel
	System.out.println("No.of sheets are  : "+sheets);
	for(int i=0;i<sheets;i++) {
		if(workbook.getSheetName(i).equalsIgnoreCase("sheet1")) {//to land on particular page ex:sheet1, sheet2 etc
		XSSFSheet sheet=workbook.getSheetAt(i);
		System.out.println(sheet);
		Iterator<Row> rows=sheet.rowIterator();
		Row firstrow=rows.next();
		Iterator<Cell> ce=	firstrow.cellIterator();
		int k=1;
		while(ce.hasNext()) {
			Cell value=ce.next();
			System.out.println(value);
			if(value.getStringCellValue().equalsIgnoreCase("Phase1")) {
				System.out.println(k);
			}
			k++;
		}
		System.out.println(k);
		}
	}
		

	}

}
