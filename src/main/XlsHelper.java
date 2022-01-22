package main;



import java.io.File;
import java.io.FileInputStream;
import java.util.Iterator;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

public class XlsHelper {
	HSSFWorkbook wb;
	HSSFSheet sheet;
	int totalRows;
	int totalColumns;
	XlsHelper(String xlsLoc) {
		try {
			FileInputStream fis=new FileInputStream(new File(xlsLoc));  
			wb=new HSSFWorkbook(fis);   
			sheet=wb.getSheetAt(0);
			totalRows=sheet.getLastRowNum();
			totalColumns=sheet.getRow(0).getLastCellNum();
		}catch(Exception e) {
			e.printStackTrace();
			System.out.println("Excel not found!!!!");
		}
	}

	// Function prints column values by index provided
	public void getColValuesByColNum(int colIndex) {
		System.out.println("Printing column values for column number "+colIndex);
		for (int row = 1; row < totalRows; row++) {
			HSSFCell cellValue=sheet.getRow(row).getCell(colIndex);
			if(cellValue!=null) {
				System.out.println(cellValue.toString());
			}else {
				break;
			}
		}
	}



	//	this function prints all cell values column wise
	public void getAllValuesColWise() {
		System.out.println("Printing all values column wise.");
		int totalRows=sheet.getLastRowNum();
		int totalColumns=sheet.getRow(0).getLastCellNum();

		for (int col = 0; col < totalColumns; col++) {
			for (int row = 1; row < totalRows; row++) {
				HSSFCell cellValue=sheet.getRow(row).getCell(col);
				if(cellValue!=null) {
					System.out.println(cellValue.toString());
				}else {
					break;
				}
			}
			System.out.println();
		}
	}
}
