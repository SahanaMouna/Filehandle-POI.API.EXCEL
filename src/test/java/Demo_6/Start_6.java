package Demo_6;  //*XL_FILE_API.poi handling*//

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.ss.formula.functions.Column;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Start_6 {
	public static void main(String[] args) throws IOException {
		FileInputStream fis=new FileInputStream("data.xlsx");
		XSSFWorkbook workbook=new XSSFWorkbook(fis);
		int sheetcount = workbook.getNumberOfSheets();
		for(int i=0;i<sheetcount;i++) {
			if(workbook.getSheetName(i).equalsIgnoreCase("Sheet2")) {
				XSSFSheet sheet = workbook.getSheetAt(i);
				Iterator<Row> rows = sheet.iterator();
//				Row firstrow = rows.next();
//				Row secondrow = rows.next();                       //* when iterating only one row*//
//				Row thirdrow = rows.next();
//				Iterator<Cell> column = firstrow.iterator();
//				Iterator<Cell> column1 = secondrow.iterator();
//				 Iterator<Cell> column2 = thirdrow.iterator();
//				System.out.println(	column.next().getStringCellValue()); //* when "next()" outside loop.it print only one data from row.
//                
//				while(column.hasNext()) {
//					System.out.println(column.next().getStringCellValue()); //* when "hasNext()" used inside loop.it print entire data from row1*//
//				}
					
//				ArrayList<String> textValue = new ArrayList<String>();
//				while(column1.hasNext()) {
//				Cell currentcell = column1.next();
//					
//				if(currentcell.getCellType()==CellType.STRING) {
//					
//					textValue.add(currentcell.getStringCellValue());
//				}
//				else {
//					textValue.add(NumberToTextConverter.toText(currentcell.getNumericCellValue()));
//				}
//				
//		}
/    		
//                workbook.close();
//				for(String value:textValue) {
//					System.out.println(value);
//				}
//}
//} 
				//********while iterating data set**********//
						while(rows.hasNext()) {
							Row currentrow = rows.next();
							ArrayList<String> textValue = new ArrayList<String>();
							Iterator<Cell> columns=currentrow.iterator();
							while(columns.hasNext()) {
								Cell column=columns.next();
								if(column.getCellType()==CellType.STRING) {
									
									textValue.add(column.getStringCellValue());
								}
								else {
									textValue.add(NumberToTextConverter.toText(column.getNumericCellValue()));
								}
							}
							 workbook.close();
								for(String value:textValue) {
									System.out.println(value);
								}
		}	
}
}
	}}