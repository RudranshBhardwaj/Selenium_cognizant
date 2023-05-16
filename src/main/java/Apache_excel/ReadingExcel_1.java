package Apache_excel;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadingExcel_1 {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
		//File-->Workbook-->Sheets-->Rows-->Cells
		//
		//FileInputStream
		//FileOutputStream
		//
		//XSSFWorkbook
		//XSSFSheet
		//XSSFRow
		//XSSFCell
		FileInputStream file=new FileInputStream("C:\\SDET\\WorkSpace\\WebDriver_Selinium\\testdata\\data.xlsx");
		XSSFWorkbook workBook=new XSSFWorkbook(file);
		XSSFSheet sheet= workBook.getSheet("Sheet2");
		int rowNo=sheet.getLastRowNum();
		System.out.println(rowNo);
		int colNo=sheet.getRow(1).getLastCellNum();
		System.out.println(colNo);
		
		for(int i=0;i<=rowNo;i++) {
			XSSFRow r=sheet.getRow(i);
			for(int c=0;c<colNo;c++) {
				String col_detail=r.getCell(c).toString();
				System.out.print(col_detail+"      ");
			}
			System.out.println();
		}

	}

}
