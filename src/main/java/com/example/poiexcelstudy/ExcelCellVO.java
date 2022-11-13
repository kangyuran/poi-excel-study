package com.example.poiexcelstudy;

import lombok.Getter;
import lombok.Setter;
import lombok.ToString;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;

@Getter @Setter @ToString
public class ExcelCellVO {

    /** 셀 값 */
    private String cellValue;

    /** 셀 스타일 */
    private XSSFCellStyle cellStyle;

    /** 줄 높이 */
    private Integer rowHeight;

    /** 셀 길이 */
    private Integer cellWidth;


    /**
     * 생성자
     *
     * @param	cellValue : 셀 값
     * @param	cellStyle : 셀 스타일
     */
    public ExcelCellVO(String cellValue,  XSSFCellStyle cellStyle) {
        this(cellValue,  cellStyle, null, null);
    }


    /**
     * 생성자
     *
     * @param	cellValue : 셀 값
     * @param	cellStyle : 셀 스타일
     * @param	rowHeight : 줄 높이
     */
    public ExcelCellVO(String cellValue,  XSSFCellStyle cellStyle, Integer rowHeight) {
        this(cellValue,  cellStyle, rowHeight, null);
    }


    /**
     * 생성자
     *
     * @param	cellValue : 셀 값
     * @param	cellStyle : 셀 스타일
     * @param	rowHeight : 줄 높이
     * @param	cellWidth : 셀 길이
     */
    public ExcelCellVO(String cellValue,  XSSFCellStyle cellStyle, Integer rowHeight, Integer cellWidth) {
        this.cellValue	= cellValue;
        this.cellStyle	= cellStyle;
        this.rowHeight	= rowHeight;
        this.cellWidth	= cellWidth;
    }
}
