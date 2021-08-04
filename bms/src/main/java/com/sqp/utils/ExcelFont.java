package com.sqp.utils;

import org.apache.poi.hssf.usermodel.HSSFFont;

/**
 * <p>Title: Excel导出工具类</p>
 * <p>Description: Excel样式</p>
 * <p>Company: </p>
 * @author 上海信而富企业管理有限公司
 * @version 1.0
 */
public class ExcelFont {

    private ExcelWriter excelWriter;
    public ExcelFont( ExcelWriter excelWriter){
        this.excelWriter = excelWriter;
    }

    public  HSSFFont getTitleFont() {
         HSSFFont font = excelWriter.getWorkBook() .createFont() ;
         font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD ) ;
         font.setFontHeightInPoints((short)20) ;
         return font;
    }

    public HSSFFont  getNormalFont() {
        HSSFFont font = excelWriter.getWorkBook() .createFont() ;
        font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD );
        return font;
    }

    public HSSFFont getDefaultFont(){
        return  excelWriter.getWorkBook() .createFont() ;
    }

    public HSSFFont cloneFont(HSSFFont source){
        HSSFFont font = this.excelWriter .getWorkBook() .createFont() ;
        font.setBoldweight(source.getBoldweight() ) ;
        font.setColor(source.getColor() ) ;
        font.setFontHeight(source.getFontHeight() ) ;
        font.setFontHeightInPoints(source.getFontHeightInPoints() ) ;
        font.setItalic(source.getItalic() ) ;
        font.setStrikeout(source.getStrikeout() ) ;
        font.setTypeOffset(source.getTypeOffset() ) ;
        font.setUnderline(source.getUnderline() ) ;

        return font;
    }


}