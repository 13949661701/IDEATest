package com.sqp.utils;

/**
 * <p>Title: Excel导出工具类</p>
 * @author 上海信而富企业管理有限公司
 * @version 1.0
 */
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.Calendar;
import java.util.Date;

import org.apache.commons.lang.ObjectUtils;
import org.apache.commons.lang.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.util.CellRangeAddress;


/**
 * <p>Title: 写Excel文件</p>
 * <p>Description:生成Excel文件，文件只包含一个工作表
 *    只处理String类型的值，没有计算功能
 * </p>
 * <p>Copyright: Copyright (c) 2003-12-31</p>
 * <p>Company: china317</p>
 * @author mjw 
 * @modify shenji 2012-5-14
 * @version 1.0
 */

public class ExcelWriter {

    //field
    private HSSFWorkbook wb ;
    private HSSFSheet sheet;
    private ExcelStyle style;
    private ExcelFont font;
    private int defaultRowHeight = 20;
      
    //get set method
    public HSSFWorkbook getWorkBook(){ return wb; }
    public HSSFSheet getSheet(){ return sheet;}
    public ExcelStyle getExcelStyle(){ return this.style ;}
    public ExcelFont getExcelFont(){ return this.font ; }
    public int getDefaultRowHeight(){ return this.defaultRowHeight ;}
    public void setDefaultRowHeight(int height){ this.defaultRowHeight = height; }

    public int getDefaultColumnWidth(){
        return this.getSheet() .getDefaultColumnWidth()  ;
    }
    public void setDefaultColumnWidth(int width){
        this.getSheet().setDefaultColumnWidth(width) ;
    }

    /**
     * 构造器
     * */
    public ExcelWriter() {
        wb = new HSSFWorkbook();
        sheet = wb.createSheet() ;
        style = new ExcelStyle(this);
        font = new ExcelFont(this);
    }

    public ExcelWriter(File sourceFile) throws Exception{
    	wb = new HSSFWorkbook(this.getClass().getResourceAsStream(sourceFile.getAbsolutePath()));
        sheet = wb.getSheetAt(0);
        style = new ExcelStyle(this);
        font = new ExcelFont(this);
    }
    
    public ExcelWriter(File sourceFile,int sheetIndex) throws Exception{
    	wb = new HSSFWorkbook(this.getClass().getResourceAsStream(sourceFile.getAbsolutePath()));
        sheet = wb.getSheetAt(sheetIndex);
        style = new ExcelStyle(this);
        font = new ExcelFont(this);
    }

    public ExcelWriter(String sourceFile) throws Exception{
    	wb = new HSSFWorkbook(this.getClass().getResourceAsStream(sourceFile));
        sheet = wb.getSheetAt(0);
        style = new ExcelStyle(this);
        font = new ExcelFont(this);
    }
    
    public ExcelWriter(String sourceFile,int sheetIndex) throws Exception{
    	wb = new HSSFWorkbook(this.getClass().getResourceAsStream(sourceFile));
        sheet = wb.getSheetAt(sheetIndex);
        style = new ExcelStyle(this);
        font = new ExcelFont(this);
    }

    /**
     * 增加标题，占用一行
     *
     * @param rowIndex 行号
     * @param title 标题值
     * @param columnFrom 起始列号（包含此列）
     * @param columnTo 结束列号 （包含此列）
     * @return HSSFRow
     * */
    public HSSFRow addTitle(int rowIndex,String title,int columnFrom ,int columnTo){
        HSSFRow row = sheet.createRow(rowIndex) ;
        HSSFCell cell = row.createCell(columnFrom) ;
        cell.setCellValue(title) ;
        cell.setCellStyle(this.getExcelStyle().getHeaderStyle());
        CellRangeAddress region = new CellRangeAddress(rowIndex,rowIndex,columnFrom,columnTo);
        sheet.addMergedRegion(region) ;
        return row ;
    }//addTitle//


   
    
    /**
     * 增加一行,values中的每个元素占用此行的一个单元格
     * @param rowIndex 行号
     * @param columnFrom 起始列号（包含此列）
     * @param values 单元格的值
     * @param style 使用的风格
     * @return HSSFRow
     * */
    public HSSFRow addRow(int rowIndex ,int columnFrom,Object[] values,HSSFCellStyle style){
        HSSFRow row = sheet.createRow(rowIndex ) ;
        row.setHeightInPoints(this.getDefaultRowHeight() ) ;
        for(int i=0;i<values.length ;i++){
            HSSFCell cell = row.createCell(columnFrom) ;
            HSSFCellStyle cloneStyle = style;//this.getExcelStyle() .cloneStyle(style) ;
            cell.setCellStyle(cloneStyle) ;
            setCellValue(values[i], cell);
            columnFrom ++;
        }
        return row ;
    }//addRow//
    
    public HSSFRow addRow(int rowIndex ,int columnFrom,Object[] values,HSSFCellStyle style,HSSFCellStyle dateStyle,HSSFCellStyle txtStyle){
        HSSFRow row = sheet.createRow(rowIndex ) ;
        row.setHeightInPoints(this.getDefaultRowHeight() ) ;
        for(int i=0;i<values.length ;i++){
            HSSFCell cell = row.createCell(columnFrom) ;
            //HSSFCellStyle cloneStyle = style;//this.getExcelStyle() .cloneStyle(style) ;
            //cell.setCellStyle(cloneStyle) ;
            
            if(values[i] instanceof Date){
            	HSSFCellStyle cloneStyle = dateStyle;//this.getExcelStyle() .cloneStyle(style) ;
                cell.setCellStyle(cloneStyle) ;
            //}else if(i==values.length-1){
           // 	HSSFCellStyle cloneStyle = txtStyle;//this.getExcelStyle() .cloneStyle(style) ;
           //     cell.setCellStyle(cloneStyle) ;
            }else{
            	 HSSFCellStyle cloneStyle = style;//this.getExcelStyle() .cloneStyle(style) ;
                 cell.setCellStyle(cloneStyle) ;
            }
            
            setCellValue(values[i], cell);
            columnFrom ++;
        }
        return row ;
    }
    /**
     * 增加一行
     * values中的每个元素占用此行的单元格的数量由ownerColumns指定，values的长度必须和ownerColumns的长度一至
     *
     * @param rowIndex 行号
     * @param columnFrom 起始列号（包含此列
     * @param values 单元格的值
     * @param ownerColumns values中每个元素占用的单元格数
     * @param style 使用的风格
     * */
    public HSSFRow addRow(int rowIndex,int columnFrom,Object[] values,int[] ownerColumns,HSSFCellStyle style){
    	HSSFRow row = sheet.createRow((short)rowIndex) ;
        row.setHeightInPoints(this.getDefaultRowHeight() ) ;
        for(int i=0 ;i<values.length ;i++){
            int tempColumnFrom = columnFrom;
            int tempColumnTo = tempColumnFrom+ownerColumns[i]-1;
            CellRangeAddress region = new CellRangeAddress(rowIndex,rowIndex,tempColumnFrom,tempColumnTo);
            sheet.addMergedRegion(region) ;
            for(int j=tempColumnFrom;j<=tempColumnTo;j++){
                HSSFCell cell = row.createCell(j) ;
                cell = row.createCell(j);
                HSSFCellStyle cloneStyle = style;// this.getExcelStyle() .cloneStyle(style) ;
                cell.setCellStyle(cloneStyle);
                
                setCellValue(values[i], cell);
            }
            columnFrom = columnFrom+ownerColumns[i];
        }
        return row;
    }//addRow//

   

  
    /**
     * 增加一个单元,此单元可占用多个单元格
     *
     * @param rowFrom 起始行号（包括此行）
     * @param rowTo 结束行号（包括此行）
     * @param columnFrom 起始列号（包括列行）
     * @param columnTo 结束列号（包括列行）
     * @param value 单元的值
     * @param style 使用的风格
     * */
    public HSSFCell addCell(int rowFrom,int rowTo,int columnFrom ,int columnTo,Object value,HSSFCellStyle style){
//    	设置第一个单元格的值
    	HSSFRow row = getRow(rowFrom);
    	HSSFCell cell = row.createCell(columnFrom); 
    	setCellValue(value, cell);
        
        //合并单元格
    	CellRangeAddress region = new CellRangeAddress(rowFrom,rowTo,columnFrom,columnTo);
    
        sheet.addMergedRegion(region);
        
        //设置样式
        HSSFCellStyle cloneStyle = style;
        for(int i=rowFrom;i<=rowTo;i++){
        	HSSFRow row_temp = getRow(i);
        	for(int j=columnFrom;j<=columnTo;j++){
            	HSSFCell cell_temp = row_temp.getCell(j);
            	if(cell_temp ==null ){
            		cell_temp = row_temp.createCell(j);
            	}
            	cell_temp.setCellStyle(cloneStyle);
        	}
        }
        
        return cell;
                
    }//addCell//
    /**
     * 得到一行，如果没有，则新建一行
     * @param rowIndex
     * @return
     */
    public HSSFRow getRow(int rowIndex) {
    	HSSFRow row = sheet.getRow(rowIndex);
    	if(row==null){
    		row = sheet.createRow(rowIndex);
    		row.setHeightInPoints(this.getDefaultRowHeight()) ;
    	}
    	return row;
    }
   
    
    /**
     * 增加一个单元,此单元可占用一个单元格
     *
     * @param rowIndex 行号
     * @param columnIndex 列号
     * @param value 单元的值
     * @param style 使用的风格
     * @retrun HSSFCell
     * */
    
    public HSSFCell addCell(int rowIndex,int columnIndex,Object value,HSSFCellStyle style){
        HSSFRow row = sheet.getRow((short)rowIndex) ;
        row.setHeightInPoints(this.getDefaultRowHeight()  ) ;
        HSSFCell cell = row.createCell(columnIndex) ;
        cell.setCellStyle(style) ;
        setCellValue(value, cell);
        return cell;
    }//addCell//
    
    
	

    /**
     * 修改某行上边框的风格
     *
     * @param rowIndex 行号
     * @param columnFrom 起始列号（包含此列）
     * @param columnTo 结束列号（包含此列）
     * @param border 边框的风格（参照HSSFCellStyle）
     * */
    public void modifyRowTopBorder(int rowIndex,int columnFrom,int columnTo,short border){
        HSSFRow row = sheet.getRow(rowIndex) ;
        for(int i=columnFrom;i<=columnTo;i++){
            row.getCell(i).getCellStyle() .setBorderTop(border)  ;
        }//for//
    }

    /**
     * 修改某行下边框的风格
     *
     * @param rowIndex 行号
     * @param columnFrom 起始列号（包含此列）
     * @param columnTo 结束列号（包含此列）
     * @param border 边框的风格（参照HSSFCellStyle）
     * */
    public void modifyRowBottomBorder(int rowIndex,int columnFrom,int columnTo,short border){
        HSSFRow row = sheet.getRow(rowIndex) ;
        for(int i=columnFrom;i<=columnTo;i++){
            row.getCell(i).getCellStyle() .setBorderBottom(border)  ;
        }//for//
    }

    /**
     * 修改某列左边框的风格
     *
     * @param columnIndex 列号
     * @param rowFrom 起始行号（包含此行）
     * @param rowTo 结束行号（包含此行）
     * @param border 边框的风格（参照HSSFCellStyle）
     * */
    public void modifyColumnLeftBorder(int columnIndex,int rowFrom,int rowTo,short border){
        for(int i=rowFrom;i<=rowTo;i++){
            HSSFRow row = sheet.getRow(i) ;
            row.getCell(columnIndex).getCellStyle() .setBorderLeft(border)  ;
        }//for//
    }


    /**
     * 修改某列右边框的风格
     *
     * @param columnIndex 列号
     * @param rowFrom 起始行号（包含此行）
     * @param rowTo 结束行号（包含此行）
     * @param border 边框的风格（参照HSSFCellStyle）
     * */
    public void modifyColumnRightBorder(int columnIndex,int rowFrom,int rowTo,short border){
        for(int i=rowFrom;i<=rowTo;i++){
            HSSFRow row = sheet.getRow(i) ;
            row.getCell(columnIndex).getCellStyle() .setBorderRight(border)  ;
        }//for//
    }

    /**
     * 设置某列的宽度
     *
     * @param columnIndex 列号
     * @param columnWidth 列宽 每单位为一个字符宽
     * */
    public void setColumnWidth(int columnIndex,int columnWidth){
        sheet.setColumnWidth(columnIndex,(short)(columnWidth*256)) ;
    }

    /**
     * 设置某行的高度
     *
     * @param rowIndex 行号
     * @param rowHeight 行高
     * */
    public void setRowHeight(int rowIndex,float rowHeight){
        sheet.getRow(rowIndex).setHeightInPoints(rowHeight)  ;
    }

    /**
     * 获取某个单元格的风格
     *
     * @param rowIndex 行号
     * @param columnIndex 列号
     * @return HSSFCellStyle
     * */
    public HSSFCellStyle getCellStyle(int rowIndex,int columnIndex){
        return sheet.getRow(rowIndex).getCell(columnIndex).getCellStyle()   ;
    }


    /**
     * 保存成文件
     *
     * @Param fileName 文件名
     * @throws IOException
     * */
    public void saveAs(String fileName)throws IOException{
        FileOutputStream file = new FileOutputStream(fileName) ;
        wb.write(file );
        file.flush() ;
        file.close() ;
    }
    
    public void output(OutputStream os)throws IOException{
    	wb.write(os);
    	os.flush();
    }
    
    
    private void setCellValue(Object value, HSSFCell cell) {
    	if(value == null){
    		cell.setCellValue("");
    		return;
    	}
		if(value instanceof Integer ){
			//整型
        	cell.setCellValue(((Integer)value).doubleValue());
		} else if(value instanceof Long){
			//整型
			cell.setCellValue(((Long)value).doubleValue());
		} else if(value instanceof Short){
			//整型
			cell.setCellValue(((Short)value).doubleValue());
		} else if(value instanceof Float){
			//浮点
			cell.setCellValue(((Float)value).floatValue());
        } else if(value instanceof Double ){
        	//浮点
        	cell.setCellValue(((Double)value).doubleValue());
        } else if(value instanceof String) {
        	//字符
        	cell.setCellValue((String)value);
        } else if(value instanceof Calendar) {
        	//日历
        	cell.setCellValue((Calendar)value);
        } else if(value instanceof Date){
        	//日期
        	cell.setCellValue((Date)value);
        } else {
        	cell.setCellValue(value.toString());
        }
	}
    
    
    public void createNewSheet(){
    	sheet = wb.createSheet();
    }
    
    /**
     * 增加一行,values中的每个元素占用此行的一个单元格
     * @param rowIndex 行号
     * @param columnFrom 起始列号（包含此列）
     * @param values 单元格的值
     * @param style 使用的风格
     * @return HSSFRow
     * */
    public HSSFRow addRow(int  rowIndex ,int columnFrom,Object[] values,HSSFCellStyle style,ExcelWriter ew,HSSFDataFormat format){
        HSSFRow row = sheet.createRow(rowIndex ) ;
        row.setHeightInPoints(this.getDefaultRowHeight()) ;
        for(int i=0;i<values.length ;i++){
            HSSFCell cell = row.createCell(columnFrom) ;
            HSSFCellStyle cloneStyle = style;
            if(rowIndex != 1){
            	cloneStyle.setAlignment(HSSFCellStyle.ALIGN_RIGHT); //水平布局：居右
            	cloneStyle.setDataFormat(format.getFormat("#,##0")); 
            }
            if(columnFrom == 0 && rowIndex != 1){
            	HSSFCellStyle createCellStyle = ew.getExcelStyle().getNormalStyle();
            	createCellStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER); //水平布局：居右
            	cell.setCellStyle(createCellStyle);
            }else{
            	cell.setCellStyle(cloneStyle);
            }
            setCellValue(values[i], cell);
            columnFrom ++;
        }
        return row ;
    }
    
    /**
     * 增加一行,values中的每个元素占用此行的一个单元格
     * @param rowIndex 行号
     * @param columnFrom 起始列号（包含此列）
     * @param values 单元格的值
     * @param style 使用的风格
     * @return HSSFRow
     * */
    public HSSFRow addRow(int  rowIndex ,int columnFrom,Object[] values,HSSFCellStyle style,HSSFCellStyle style2, ExcelWriter ew,HSSFDataFormat format){
        HSSFRow row = sheet.createRow(rowIndex ) ;
        row.setHeightInPoints(this.getDefaultRowHeight()) ;
        for(int i=0;i<values.length ;i++){
            HSSFCell cell = row.createCell(columnFrom) ;
            HSSFCellStyle cloneStyle = style;
            if(rowIndex != 1){
            	cloneStyle.setAlignment(HSSFCellStyle.ALIGN_RIGHT); //水平布局：居右
            	cloneStyle.setDataFormat(format.getFormat("#,##0")); 
            }
            if(columnFrom == 0 && rowIndex != 1){
            	HSSFCellStyle createCellStyle = ew.getExcelStyle().getNormalStyle();
            	createCellStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER); //水平布局：居右
            	cell.setCellStyle(createCellStyle);
            }else if(columnFrom == 1 && rowIndex != 1){
            	cell.setCellStyle(style2);
            }else{
            	cell.setCellStyle(cloneStyle);
            }
            setCellValue(values[i], cell);
            columnFrom ++;
        }
        return row ;
    }
    
    
    /**
     * 增加一行,values中的每个元素占用此行的一个单元格
     * @param rowIndex 行号
     * @param columnFrom 起始列号（包含此列）
     * @param values 单元格的值
     * @param style 使用的风格
     * @return HSSFRow
     * */
    public HSSFRow addRow(int rowIndex ,int columnFrom,Object[] values,HSSFCellStyle style,HSSFCellStyle style2,HSSFCellStyle style3,ExcelWriter ew,HSSFDataFormat format){
        HSSFRow row = sheet.createRow(rowIndex ) ;
        row.setHeightInPoints(this.getDefaultRowHeight()) ;
        for(int i=0;i<values.length ;i++){
            HSSFCell cell = row.createCell(columnFrom) ;
            HSSFCellStyle cloneStyle = style;
            if(rowIndex != 1){
            	cloneStyle.setAlignment(HSSFCellStyle.ALIGN_RIGHT); //水平布局：居右
            }
            if(columnFrom == 0 && rowIndex != 1){
            	cell.setCellStyle(style2);
            }else if((columnFrom == 1 || columnFrom == 2 || columnFrom == 5 || columnFrom == 6) && rowIndex != 1){
            	cell.setCellStyle(style3);
            } else{
            	cell.setCellStyle(style);
            }
            setCellValue(values[i], cell);
            columnFrom ++;
        }
        return row ;
    }
    
    /**
     * 增加一行,values中的每个元素占用此行的一个单元格 元素左对齐
     * @param rowIndex 行号
     * @param columnFrom 起始列号（包含此列）
     * @param values 单元格的值
     * @param style 使用的风格
     * @return HSSFRow
     * */
    public HSSFRow addRowLeftColumn(int  rowIndex ,int columnFrom,Object[] values,HSSFCellStyle style,ExcelWriter ew,HSSFDataFormat format){
        HSSFRow row = sheet.createRow(rowIndex ) ;
        row.setHeightInPoints(this.getDefaultRowHeight()) ;
        for(int i=0;i<values.length ;i++){
            HSSFCell cell = row.createCell(columnFrom) ;
            HSSFCellStyle cloneStyle = style;
            if(rowIndex != 1){
            	cloneStyle.setAlignment(HSSFCellStyle.ALIGN_LEFT); //水平布局：居右
            	cloneStyle.setDataFormat(format.getFormat("#,##0")); 
            }
            if(columnFrom == 0 && rowIndex != 1){
            	HSSFCellStyle createCellStyle = ew.getExcelStyle().getNormalStyle();
            	createCellStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER); //水平布局：居右
            	cell.setCellStyle(createCellStyle);
            }else{
            	cell.setCellStyle(cloneStyle);
            }
            setCellValue(values[i], cell);
            columnFrom ++;
        }
        return row ;
    }
    public static void main(String[] a)throws Exception{
    
    }
}
