package com.FrameworkFunctions;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelUtils {

	public static XSSFSheet ExcelWSheet;

	private static XSSFWorkbook ExcelWBook;

	private static XSSFCell Cell;

	private static XSSFRow Row;
	static DataFormatter formatter=new DataFormatter(); 

	/**
	 * This method is to set the File path and to open the Excel file, Pass Excel Path and Sheetname as Arguments to this method
	 * @param Path
	 * @param SheetName
	 * @throws Exception
	 */
	public static void setExcelFile(String Path,String SheetName) throws Exception {


		try {

			// Open the Excel file

			FileInputStream ExcelFile = new FileInputStream(Path);

			// Access the required test data sheet

			ExcelWBook = new XSSFWorkbook(ExcelFile);

			ExcelWSheet = ExcelWBook.getSheet(SheetName);

		} catch (Exception e){

			throw (e);

		}

	}

	public static Object[][] getTableArray(String FilePath, String SheetName, int iTestCaseRow)    throws Exception


	{   

		String[][] tabArray = null;


		try{

			FileInputStream ExcelFile = new FileInputStream(FilePath);

			// Access the required test data sheet

//			ExcelWBook = new XSSFWorkbook(ExcelFile);

			ExcelWSheet = ExcelWBook.getSheet(SheetName);

			int startCol = 1;

			int ci=0,cj=0;

			int usedCells = 0;

			//int totalRows = 1;

			int totalCols = ExcelWSheet.getRow(iTestCaseRow).getLastCellNum()-1;
			//Loop to get only used cells from Excel - getPhysicalNumberofCells not working	   
			for (int j=startCol;j<=totalCols;j++){
				String cellValue = getCellData(iTestCaseRow,j);
				if(cellValue != ""){
					usedCells++;
				}
				else{
					//do nothing
				}

			}

			//Create array to store values	  			   
			tabArray=new String[1][usedCells];

			for (int k=startCol;k<=usedCells;k++, cj++)

			{
				tabArray[ci][cj]=getCellData(iTestCaseRow,k).trim();
			}
			System.out.println("Value fetched successfully");

		}

		catch (FileNotFoundException e)

		{

			System.out.println("Could not read the Excel sheet");

			e.printStackTrace();

		}

		catch (IOException e)

		{

			System.out.println("Could not read the Excel sheet");

			e.printStackTrace();

		}

		return(tabArray);

	}

	/**
	 * This method is to read the test data from the Excel cell, in this we are passing parameters as Row num and Col num
	 * @param RowNum
	 * @param ColNum
	 * @return
	 * @throws Exception
	 */
	/*public static String getCellData(int RowNum, int ColNum) throws Exception{


		try{

			Cell = ExcelWSheet.getRow(RowNum).getCell(ColNum);

			String CellData = Cell.getStringCellValue();

			return CellData;

		}catch (Exception e){

			return"";

		}

	}*/
	
	public static String getCellData(int RowNum, int ColNum) throws Exception{


        try{

              Cell = ExcelWSheet.getRow(RowNum).getCell(ColNum);
              String  CellData=formatter.formatCellValue(ExcelWSheet.getRow(RowNum).getCell(ColNum));
//            String CellData = Cell.getStringCellValue();
//            String CellData = formatter.formatCellValue(Cell);
              
              if(Cell == null || Cell.getCellType() == XSSFCell.CELL_TYPE_BLANK) {
                    Cell = ExcelWSheet.getRow(RowNum).createCell(ColNum);
                    Cell = ExcelWSheet.getRow(RowNum).getCell(ColNum);
                    Cell.setCellType(Cell.CELL_TYPE_STRING);
//                  Cell.setCellValue("'");
                    CellData = Cell.getStringCellValue().trim();
              }
//            
//            if(CellData.equals(""))  {
//                  Cell.setCellType(Cell.CELL_TYPE_STRING); 
////                Cell.setCellValue("'");
//                  CellData = Cell.getStringCellValue();
//                  
//            }
//             CellData = formatter.formatCellValue(Cell);//cell.getStringCellValue();

              return CellData;

             }catch (Exception e){

                  return"";

                  }

            }


	public static int getRowContains(String sTestCaseName, int colNum,int iterate) throws Exception{
		int i;

		try {

			int rowCount = ExcelUtils.getRowUsed();

			for ( i=0 ; i<rowCount; i++){

				if  (ExcelUtils.getCellData(i,colNum).equalsIgnoreCase(sTestCaseName)){

					break;

				}

			}

			return i;

		}catch (Exception e){

			throw(e);

		}

	}

	public static int getRowUsed() throws Exception {


		try{

			int RowCount = ExcelWSheet.getLastRowNum();

			return RowCount;

		}catch (Exception e){

			System.out.println(e.getMessage());

			throw (e);

		}

	}


}

