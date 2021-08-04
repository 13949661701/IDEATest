package com.sqp.utils;

import java.text.DecimalFormat;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Scanner;

import org.apache.commons.logging.Log;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;


/**
 * <p>Title: Excel导出工具类</p>
 * @author 上海信而富企业管理有限公司
 * @version 1.0
 */
public class ExcelReader {

	private static Log log;
	private HSSFSheet aSheet;
	private Map columnCommentMap;
	private Map columnTypeMap;
	private Map columnNameMap;
	private List rowList;

	public ExcelReader() {
/*		try {
			log = LogFactory.getLog(getClass().getName());
			workbook = new HSSFWorkbook(new FileInputStream(
					"E:\\doc\\KE_YUN_XIAN_LU20091127.xls"));
			log.info("the Numbers Of Sheets : " + workbook.getNumberOfSheets());
			aSheet = workbook.getSheetAt(0);
			columnCommentMap = new HashMap();
			columnNameMap = new HashMap();
			columnTypeMap = new HashMap();
			rowList = new ArrayList();
		} catch (FileNotFoundException e) {
			log.info("文件没找到 ！");
		} catch (IOException e) {
			log.info("文件读入错 ！");
		}*/
	}

	public void readColumnComment() {
		DecimalFormat df = new DecimalFormat("#");
		HSSFRow aRow = aSheet.getRow(0);
		
		for (short cellNumOfRow = 0; cellNumOfRow <= aRow.getLastCellNum(); cellNumOfRow++) {

			if (null != aRow.getCell(cellNumOfRow)) {
				HSSFCell aCell = aRow.getCell(cellNumOfRow);

				int cellType = aCell.getCellType();
				
				switch (cellType) {
				case 0:// Numeric
					String strCell = df.format(aCell.getNumericCellValue());

					log.info(strCell);

					break;
				case 1:// String
					strCell = aCell.getStringCellValue();
					columnCommentMap.put(cellNumOfRow, strCell);
					
					log.info("String : " + strCell);

					break;
				default:
					// System.out.println("格式不对不读");//其它格式的数据
				}

			}

		}
	}

	public void readRowData() {
		DecimalFormat df = new DecimalFormat("#");
		
		log.info("Last Row Number : " + aSheet.getLastRowNum());
		
		for(int rowNumOfSheet = 1; rowNumOfSheet <= aSheet.getLastRowNum(); rowNumOfSheet++){
			HSSFRow aRow = aSheet.getRow(rowNumOfSheet);
			
			if(aRow==null) continue;
			Map cellDataMap = new HashMap();
			rowList.add(cellDataMap);
			//log.info("aRow size : " + aRow.getLastCellNum());
			
			for (short cellNumOfRow = 0; cellNumOfRow <= aRow.getLastCellNum(); cellNumOfRow++) {

				if (null != aRow.getCell(cellNumOfRow)) {
					HSSFCell aCell = aRow.getCell(cellNumOfRow);

					int cellType = aCell.getCellType();
					
					switch (cellType) {
					case 0:// Numeric
						String strCell = df.format(aCell.getNumericCellValue());
						columnTypeMap.put(cellNumOfRow, "Numeric");
						cellDataMap.put(cellNumOfRow, strCell);
						break;
					case 1:// String
						strCell = aCell.getStringCellValue();
						columnTypeMap.put(cellNumOfRow, "String");
						cellDataMap.put(cellNumOfRow, strCell);
						break;
					default:
						System.out.println("格式不对不读");//其它格式的数据
					}

				}
				//else 
					//log.info("cell is null");

			}
			
		}
		log.info("row size : " + rowList.size());
		
	}
	
	public void editColumnName() {
		Scanner in = new Scanner(System.in);
		for(Object  col:columnCommentMap.keySet()){
			System.out.print("name this column of " + columnCommentMap.get(col) + " :  ");
			String columnName = in.nextLine();
			if(columnName.equalsIgnoreCase("quit")) break;
			else if (columnName.equalsIgnoreCase("next")) {
				continue;
			}
			else{
				columnNameMap.put(col, columnName);
			}
		}
	}
	
	public void printColumnName() {
		for(Object  col:columnCommentMap.keySet()){
			String columnName = (String)columnNameMap.get(col);
			String columnUpperCaseName = "undefined" ;
			if(columnName !=null)  columnUpperCaseName = columnName.toUpperCase();
			System.out.print("public static final String " + columnUpperCaseName );
			System.out.print(" =\"" + columnName + "\"; ");
			System.out.print(" \t// " + columnCommentMap.get(col) + " \t" + col);
			System.out.println(" \t" + columnTypeMap.get(col));
			
		}
	}

	public static void main(String[] args) {
		ExcelReader poi = new ExcelReader();
		// poi.CreateExcel();
		poi.readColumnComment();
		poi.readRowData();
		
		Scanner in = new Scanner(System.in);
		while(true){
			System.out.println("input command :");
			String command = in.nextLine();
			if(command.equalsIgnoreCase("quit")) break;
			else if(command.equalsIgnoreCase("edit")) poi.editColumnName();
			else if(command.equalsIgnoreCase("print")) poi.printColumnName();
		}
		System.out.println("the End");
	}
	
}
