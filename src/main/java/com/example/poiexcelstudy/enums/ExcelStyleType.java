package com.example.poiexcelstudy.enums;

import lombok.Getter;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.BuiltinFormats;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;

import java.util.concurrent.atomic.AtomicReference;
import java.util.stream.Stream;

/**
 * Excel Style
 *
 * Id 전달. ( XSSFCellStyle에 Style 적용 시 1부터 시작한다. (ordinal은 0부터 시작) )
 * Id를 Enum 항목으로 넣지 않은 이유
 * : Style 적용 시 Number값이 저장되지 않고 등록하는 순번대로 적용
 * : 사용의 편의를 위해 Id 항목을 넣게 되면 개발자의 오류로 인하여 Style 적용시 오류가 발생할 수 있을 듯 함.
 */
@Getter
public enum ExcelStyleType {

    /**
     * <b>Title</b>
     * <br/>
     * <br/>Align : Center
     * <br/>Vertical-Align : Center
     * <br/>Font : 15 Size, Bold, Black
     * <br/>Background Color : White
     * <br/>Border Style : All None
     * <br/>Wrap Text : false
     */
    CELLSTYLE_TITLE
            (
                    "@",
                    HorizontalAlignment.CENTER,
                    VerticalAlignment.CENTER,
                    ExcelStyleFontType.FONT_15_BOLD_BLACK,
                    HSSFColor.HSSFColorPredefined.WHITE,
                    FillPatternType.SOLID_FOREGROUND,
                    ExcelStyleBorderType.BORDER_NONE_ALL,
                    false
            ),

    /**
     * <b>Header</b>
     * <br/>
     * <br/>Align : Center
     * <br/>Vertical-Align : Center
     * <br/>Font : 9 Size, Bold, Black
     * <br/>Background Color : Gray
     * <br/>Border Style : All Thin
     * <br/>Wrap Text : false
     */
    CELLSTYLE_HEADER
            (
                    "@",
                    HorizontalAlignment.CENTER,
                    VerticalAlignment.CENTER,
                    ExcelStyleFontType.FONT_9_BOLD_BLACK,
                    HSSFColor.HSSFColorPredefined.GREY_25_PERCENT,
                    FillPatternType.SOLID_FOREGROUND,
                    ExcelStyleBorderType.BORDER_THIN_ALL,
                    false
            ),

    /**
     * <b>Total Cnt</b>
     * <br/>
     * <br/>Align : Left
     * <br/>Vertical-Align : Center
     * <br/>Font : 10 Size, Bold, Black
     * <br/>Background Color : White
     * <br/>Border Style : All None
     * <br/>Wrap Text : false
     */
    CELLSTYLE_TOTAL_CNT
            (
                    "@",
                    HorizontalAlignment.LEFT,
                    VerticalAlignment.CENTER,
                    ExcelStyleFontType.FONT_10_BOLD_BLACK,
                    HSSFColor.HSSFColorPredefined.WHITE,
                    FillPatternType.SOLID_FOREGROUND,
                    ExcelStyleBorderType.BORDER_NONE_ALL,
                    false

            ),

    /**
     * <b>Total Sum</b>
     * <br/>
     * <br/>Align : Center
     * <br/>Vertical-Align : Center
     * <br/>Font : 10 Size, Bold, Black
     * <br/>Background Color : Gray
     * <br/>Border Style : All None
     * <br/>Wrap Text : false
     */
    CELLSTYLE_TOTAL_SUM
            (
                    "@",
                    HorizontalAlignment.CENTER,
                    VerticalAlignment.CENTER,
                    ExcelStyleFontType.FONT_10_BOLD_BLACK,
                    HSSFColor.HSSFColorPredefined.GREY_50_PERCENT,
                    FillPatternType.SOLID_FOREGROUND,
                    ExcelStyleBorderType.BORDER_THIN_ALL,
                    false
            ),

    /**
     * <b>Left</b>
     * <br/>
     * <br/>Align : Left
     * <br/>Vertical-Align : Center
     * <br/>Font : 9 Size, None, Black
     * <br/>Background Color : Auto (없음)
     * <br/>Border Style : All Thin
     */
    CELLSTYLE_LEFT
            (
                    "@",
                    HorizontalAlignment.LEFT,
                    VerticalAlignment.CENTER,
                    ExcelStyleFontType.FONT_9_NONE_BLACK,
                    HSSFColor.HSSFColorPredefined.AUTOMATIC,
                    FillPatternType.SOLID_FOREGROUND,
                    ExcelStyleBorderType.BORDER_THIN_ALL,
                    false
            ),

    /**
     * <b>Center</b>
     * <br/>
     * <br/>Align : Center
     * <br/>Vertical-Align : Center
     * <br/>Font : 9 Size, None, Black
     * <br/>Background Color : Auto (없음)
     * <br/>Border Style : All Thin
     */
    CELLSTYLE_CENTER
            (
                    "@",
                    HorizontalAlignment.CENTER,
                    VerticalAlignment.CENTER,
                    ExcelStyleFontType.FONT_9_NONE_BLACK,
                    HSSFColor.HSSFColorPredefined.AUTOMATIC,
                    FillPatternType.SOLID_FOREGROUND,
                    ExcelStyleBorderType.BORDER_THIN_ALL,
                    false
            ),

    /**
     * <b>Number</b>
     * <br/>
     * <br/>Type : 0 (숫자)
     * <br/>Align : Right
     * <br/>Vertical-Align : Center
     * <br/>Font : 9 Size, None, Black
     * <br/>Background Color : Auto (없음)
     * <br/>Border Style : All Thin
     */
    CELLSTYLE_NUMBER
            (
                    "0",
                    HorizontalAlignment.RIGHT,
                    VerticalAlignment.CENTER,
                    ExcelStyleFontType.FONT_9_NONE_BLACK,
                    HSSFColor.HSSFColorPredefined.AUTOMATIC,
                    FillPatternType.SOLID_FOREGROUND,
                    ExcelStyleBorderType.BORDER_THIN_ALL,
                    false
            ),

    /**
     * <b>Number Double</b>
     * <br/>
     * <br/>Type : 0.00 (숫자(Double))
     * <br/>Align : Right
     * <br/>Vertical-Align : Center
     * <br/>Font : 9 Size, None, Black
     * <br/>Background Color : Auto (없음)
     * <br/>Border Style : All Thin
     */
    CELLSTYLE_NUMBER_DOUBLE
            (
                    "0.00",
                    HorizontalAlignment.RIGHT,
                    VerticalAlignment.CENTER,
                    ExcelStyleFontType.FONT_9_NONE_BLACK,
                    HSSFColor.HSSFColorPredefined.AUTOMATIC,
                    FillPatternType.SOLID_FOREGROUND,
                    ExcelStyleBorderType.BORDER_THIN_ALL,
                    false
            ),

    /**
     * <b>Currency</b>
     * <br/>
     * <br/>Type : #,##0 (통화(정수))
     * <br/>Align : Right
     * <br/>Vertical-Align : Center
     * <br/>Font : 9 Size, None, Black
     * <br/>Background Color : Auto (없음)
     * <br/>Border Style : All Thin
     */
    CELLSTYLE_CURRENCY
            (
                    "#,##0",
                    HorizontalAlignment.RIGHT,
                    VerticalAlignment.CENTER,
                    ExcelStyleFontType.FONT_9_NONE_BLACK,
                    HSSFColor.HSSFColorPredefined.AUTOMATIC,
                    FillPatternType.SOLID_FOREGROUND,
                    ExcelStyleBorderType.BORDER_THIN_ALL,
                    false
            ),

    /**
     * <b>Currency Double</b>
     * <br/>
     * <br/>Type : #,##0.00 (통화(Double))
     * <br/>Align : Right
     * <br/>Vertical-Align : Center
     * <br/>Font : 9 Size, None, Black
     * <br/>Background Color : Auto (없음)
     * <br/>Border Style : All Thin
     */
    CELLSTYLE_CURRENCY_DOUBLE
            (
                    "#,##0.00",
                    HorizontalAlignment.RIGHT,
                    VerticalAlignment.CENTER,
                    ExcelStyleFontType.FONT_9_NONE_BLACK,
                    HSSFColor.HSSFColorPredefined.AUTOMATIC,
                    FillPatternType.SOLID_FOREGROUND,
                    ExcelStyleBorderType.BORDER_THIN_ALL,
                    false
            ),

    /**
     * <b>Currency Total Sum</b>
     * <br/>
     * <br/>Type : #,##0 (숫자)
     * <br/>Align : Right
     * <br/>Vertical-Align : Center
     * <br/>Font : 9 Size, None, Black
     * <br/>Background Color : Auto (없음)
     * <br/>Border Style : All Thin
     */
    CELLSTYLE_CURRENCY_TOTAL_SUM
            (
                    "#,##0",
                    HorizontalAlignment.RIGHT,
                    VerticalAlignment.CENTER,
                    ExcelStyleFontType.FONT_9_NONE_BLACK,
                    HSSFColor.HSSFColorPredefined.GREY_25_PERCENT,
                    FillPatternType.SOLID_FOREGROUND,
                    ExcelStyleBorderType.BORDER_THIN_ALL,
                    false
            ),

    ;


    /** Cell Format */
    private final String format;
    /** 가로정렬 */
    private final HorizontalAlignment align;
    /** 세로정렬 */
    private final VerticalAlignment verticalAlign;
    /** Font Style */
    private final ExcelStyleFontType font;
    /** Background Color */
    private final HSSFColor.HSSFColorPredefined backgroundColor;
    /** Bacground Color 채우기  */
    private final FillPatternType pattern;
    /** Border Style */
    private final ExcelStyleBorderType excelStyleBorderType;
    /** 줄바꿈가능여부 */
    private final boolean wrapText;



    ExcelStyleType(String format, HorizontalAlignment align, VerticalAlignment verticalAlign, ExcelStyleFontType font,
                   HSSFColor.HSSFColorPredefined backgroundColor, FillPatternType pattern, ExcelStyleBorderType excelStyleBorderType, boolean wrapText) {
        this.format = format;
        this.align = align;
        this.verticalAlign = verticalAlign;
        this.font = font;
        this.backgroundColor = backgroundColor;
        this.pattern = pattern;
        this.excelStyleBorderType = excelStyleBorderType;
        this.wrapText = wrapText;
    }


    /**
     * 전달된 workbook에 Style적용.
     *
     * @param workbook
     * @return XSSFFont
     */
    public static XSSFCellStyle getWorkBookStyle(SXSSFWorkbook workbook) {
        AtomicReference<XSSFCellStyle> cellStyle = null;
        Stream.of(ExcelStyleType.values()).forEach(v ->
                {
                    cellStyle.set((XSSFCellStyle) workbook.createCellStyle());
                    
                    cellStyle.get().setDataFormat((short) BuiltinFormats.getBuiltinFormat(v.format));
                    cellStyle.get().setAlignment(v.align);
                    cellStyle.get().setVerticalAlignment(v.verticalAlign);
                    cellStyle.get().setFont(workbook.getFontAt(v.font.ordinal() + 1));
                    cellStyle.get().setFillForegroundColor(v.backgroundColor.getIndex());
                    cellStyle.get().setFillPattern(v.pattern);
                    cellStyle.get().setBorderLeft(v.excelStyleBorderType.getBorderLeft());
                    cellStyle.get().setBorderRight(v.excelStyleBorderType.getBorderRight());
                    cellStyle.get().setBorderBottom(v.excelStyleBorderType.getBorderBottom());
                    cellStyle.get().setBorderTop(v.excelStyleBorderType.getBorderTop());
                    cellStyle.get().setWrapText(v.wrapText);
                }
        );
        return cellStyle.get();
    }



}
