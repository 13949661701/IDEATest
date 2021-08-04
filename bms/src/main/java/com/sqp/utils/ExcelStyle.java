package com.sqp.utils;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.util.HSSFColor;

/**
 * <p>Title: Excel导出工具类</p>
 * @author 上海信而富企业管理有限公司
 * @version 1.0
 */
public class ExcelStyle {

    private ExcelWriter excelWriter;
    public ExcelStyle( ExcelWriter excelWriter){
        this.excelWriter = excelWriter;
    }

    public HSSFCellStyle getDefaultStyle(){
       return excelWriter.getWorkBook() .createCellStyle() ;
    }

    /**
     * 标题列样式
     * @return
     */
    public  HSSFCellStyle getHeaderStyle(){
        HSSFCellStyle style = excelWriter.getWorkBook() .createCellStyle() ;
        style.setAlignment(HSSFCellStyle.ALIGN_CENTER ) ;
        style.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER ) ;
        style.setFont( this.excelWriter .getExcelFont().getTitleFont()) ;
        return style;
    }
    
    /**
     * th内容样式
     * @return
     */
    public  HSSFCellStyle getTitleStyle(){
        HSSFCellStyle style = excelWriter.getWorkBook() .createCellStyle() ;
        style.setAlignment(HSSFCellStyle.ALIGN_CENTER ) ;
        style.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER ) ;
        style.setFont( this.excelWriter .getExcelFont().getDefaultFont()  ) ;
        style.setBorderTop(HSSFCellStyle.BORDER_THIN ) ;
        style.setBorderBottom(HSSFCellStyle.BORDER_THIN ) ;
        style.setBorderLeft(HSSFCellStyle.BORDER_THIN ) ;
        style.setBorderRight(HSSFCellStyle.BORDER_THIN ) ;
        style.setFillForegroundColor(HSSFColor.LIME.index);
        style.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND) ;
        return style;
    }

    /**
     * 元素内容列样式
     * @return
     */
    public HSSFCellStyle getNormalStyle(){
        HSSFCellStyle style = excelWriter.getWorkBook() .createCellStyle() ;
        style.setAlignment(HSSFCellStyle.ALIGN_CENTER ) ;
        style.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER ) ;
        style.setFont( this.excelWriter .getExcelFont().getDefaultFont()  ) ;
        style.setBorderTop(HSSFCellStyle.BORDER_THIN ) ;
        style.setBorderBottom(HSSFCellStyle.BORDER_THIN ) ;
        style.setBorderLeft(HSSFCellStyle.BORDER_THIN ) ;
        style.setBorderRight(HSSFCellStyle.BORDER_THIN ) ;
        return style;
    }
    
    /**
     * 设置无边框单元格样式
     * @return
     */
    public HSSFCellStyle getNoBorderStyle(){
        HSSFCellStyle style = excelWriter.getWorkBook() .createCellStyle() ;
        style.setAlignment(HSSFCellStyle.ALIGN_CENTER ) ;
        style.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER ) ;
        style.setFont( this.excelWriter .getExcelFont().getDefaultFont()  ) ;
        style.setBorderTop(HSSFCellStyle.BORDER_NONE ) ;
        style.setBorderBottom(HSSFCellStyle.BORDER_NONE ) ;
        style.setBorderLeft(HSSFCellStyle.BORDER_NONE ) ;
        style.setBorderRight(HSSFCellStyle.BORDER_NONE ) ;
        return style;
    }

    public  HSSFCellStyle getNormalStyle(short fillColor){
        HSSFCellStyle style = getNormalStyle();
        style.setFillForegroundColor(fillColor) ;
        style.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND) ;
        return style;
    }

    private HSSFCellStyle cloneStyle(HSSFCellStyle source){
        HSSFCellStyle obj = this.excelWriter .getWorkBook() .createCellStyle() ;
        obj.setAlignment(source.getAlignment() ) ;
        obj.setBorderBottom(source.getBorderBottom() );
        obj.setBorderLeft(source.getBorderLeft() ) ;
        obj.setBorderRight(source.getBorderRight() ) ;
        obj.setBorderTop(source.getBorderTop() ) ;
        obj.setBottomBorderColor(source.getBottomBorderColor() ) ;
        obj.setDataFormat(source.getDataFormat() ) ;
        obj.setFillBackgroundColor(source.getFillBackgroundColor() ) ;
        obj.setFillForegroundColor(source.getFillForegroundColor() ) ;
        obj.setFillPattern(source.getFillPattern() ) ;
        obj.setFont(this.excelWriter .getExcelFont() .cloneFont( this.excelWriter .getWorkBook().getFontAt(source.getFontIndex() ) )) ;
        obj.setHidden(source.getHidden() ) ;
        obj.setIndention(source.getIndention() ) ;
        obj.setLeftBorderColor(source.getLeftBorderColor() ) ;
        obj.setLocked(source.getLocked() ) ;
        obj.setRightBorderColor(source.getRightBorderColor() ) ;
        obj.setRotation(source.getRotation() ) ;
        obj.setTopBorderColor(source.getTopBorderColor() ) ;
        obj.setVerticalAlignment(source.getVerticalAlignment() ) ;
        obj.setWrapText(source.getWrapText() ) ;
        return obj;
    }

}