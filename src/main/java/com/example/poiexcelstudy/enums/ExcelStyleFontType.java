package com.example.poiexcelstudy.enums;

import lombok.Getter;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFFont;

import java.util.concurrent.atomic.AtomicReference;
import java.util.stream.Stream;

/**
 * Excel Style Font
 *
 * Id를 Enum 항목으로 넣지 않은 이유
 * : Style 적용 시 Number값이 저장되지 않고 등록하는 순번대로 적용
 * : 사용의 편의를 위해 Id 항목을 넣게 되면 개발자의 오류로 인하여 Style 적용시 오류가 발생할 수 있을 듯 함.
 */
@Getter
public enum ExcelStyleFontType {

    /**
     * Default
     */
    FONT_9_NONE_BLACK("맑은 고딕", (short)9, false, HSSFColor.HSSFColorPredefined.BLACK),
    /**
     * Header
     */
    FONT_9_BOLD_BLACK("맑은 고딕", (short)9, true, HSSFColor.HSSFColorPredefined.BLACK),
    /**
     * Total Count, Sum
     */
    FONT_10_BOLD_BLACK("맑은 고딕", (short)9, true, HSSFColor.HSSFColorPredefined.BLACK),
    /**
     * Title
     */
    FONT_15_BOLD_BLACK("맑은 고딕", (short)15, true, HSSFColor.HSSFColorPredefined.BLACK),
    ;

    private final String fontName;
    private final Short fontSize;
    private final Boolean fontBold;
    private final HSSFColor.HSSFColorPredefined fontColor;


    ExcelStyleFontType(String fontName, Short fontSize, Boolean fontBold, HSSFColor.HSSFColorPredefined fontColor) {
        this.fontName = fontName;
        this.fontSize = fontSize;
        this.fontBold = fontBold;
        this.fontColor = fontColor;
    }


    /**
     * 전달된 workbook에 Font Style적용.
     *
     * @param workbook
     * @return XSSFFont
     */
    public static XSSFFont getWorkBookFontStyle(SXSSFWorkbook workbook) {
        AtomicReference<XSSFFont> font = null;
        Stream.of(ExcelStyleFontType.values()).forEach(v ->
                {
                    font.set((XSSFFont) workbook.createFont());

                    font.get().setFontName(v.fontName);
                    font.get().setFontHeightInPoints(v.fontSize);
                    font.get().setBold(v.fontBold);
                    font.get().setColor(v.fontColor.getIndex());
                }
        );
        return font.get();
    }
}