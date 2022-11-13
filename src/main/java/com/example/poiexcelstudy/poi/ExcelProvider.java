package com.example.poiexcelstudy.poi;

import com.example.poiexcelstudy.ExcelCellVO;
import com.example.poiexcelstudy.consts.ExcelStyleConst;
import com.example.poiexcelstudy.enums.ExcelStyleType;
import lombok.RequiredArgsConstructor;
import org.apache.commons.lang3.math.NumberUtils;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.springframework.stereotype.Component;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.UnsupportedEncodingException;
import java.math.BigDecimal;
import java.net.URLEncoder;
import java.text.NumberFormat;
import java.util.List;

@Component
@RequiredArgsConstructor
public class ExcelProvider {


//    private final MessageUtil messageUtil;



    /**
     * 엑셀 파일 명 및 ResponseHeader 세팅
     *
     * @param    excelFileName 엑셀 파일명
     * @param    request
     * @param    response
     * @throws UnsupportedEncodingException
     */
    public void setExcelFile(String excelFileName, HttpServletRequest request, HttpServletResponse response)
            throws UnsupportedEncodingException {
        String userAgent = request.getHeader("User-Agent");
        String encSpace = URLEncoder.encode(" ", "UTF-8");

        if (userAgent.indexOf("MSIE 5.5") > -1) {
            excelFileName = (URLEncoder.encode(excelFileName, "UTF-8")).replace(encSpace, " ");
        } else if (userAgent.indexOf("MSIE") > -1 || userAgent.indexOf("Trident") > -1
                || userAgent.indexOf("Edge") > -1) {
            excelFileName = (URLEncoder.encode(excelFileName, "UTF-8")).replace(encSpace, " ");
        } else {
            excelFileName = new String(excelFileName.getBytes("UTF-8"), "latin1");
        }

        response.setHeader("Content-Disposition", "attachment; filename=" + excelFileName);
    }


    /**
     * 타이틀, 헤더, No데이터 등 데이터 영역을 제외한 시트 내용 생성(기본 레이아웃)
     *
     * @param    rowIdx 줄 인덱스
     * @param    workbook 워크북
     * @param    titleCell 타이틀 셀 정보
     * @param    headerCellList 헤더 셀 정보
     * @param    dataCnt 데이터 수
     * @throws Exception
     */
    public int createSheetDefault(int rowIdx, SXSSFWorkbook workbook, ExcelCellVO titleCell,
                                  List<ExcelCellVO> headerCellList, int dataCnt) throws Exception {

        SXSSFSheet sheet = createSheet(workbook, titleCell.getCellValue());
        rowIdx = createTitleRow(rowIdx, workbook, titleCell, headerCellList, sheet);
        rowIdx = createTotalCntRow(rowIdx, workbook, headerCellList, dataCnt, sheet);
        rowIdx = createHeaderRow(rowIdx, workbook, headerCellList);
        if (dataCnt == 0) {
            rowIdx = createNoDataRow(rowIdx, workbook, headerCellList, dataCnt, sheet);
        }

        return rowIdx;
    }


    /**
     * 시트생성(이미 생성된 시트명이 있을경우 "시트명 + 넘버링")
     *
     * @param workbook
     * @param sheetName
     * @return
     */
    public SXSSFSheet createSheet(SXSSFWorkbook workbook, String sheetName) {
        SXSSFSheet sheet;
        try {
            sheet = workbook.createSheet(sheetName);
        } catch (IllegalArgumentException ie) {
            sheet = workbook.createSheet(sheetName + (workbook.getNumberOfSheets() + 1));
        }
        return sheet;
    }


    /**
     * Title Row 생성 ( 행 전체 )
     *
     * @param rowIdx
     * @param workbook
     * @param titleCell
     * @param headerCellList
     * @param sheet
     * @return
     * @throws Exception
     */
    public int createTitleRow(int rowIdx, SXSSFWorkbook workbook, ExcelCellVO titleCell, List<ExcelCellVO> headerCellList, SXSSFSheet sheet) throws Exception {
        /** 타이틀 영역 생성 */
        for (int i = 0; i < headerCellList.size(); i++) {
            if (i == 0) {
                rowIdx = this.createCell(rowIdx, i, workbook, titleCell.getCellValue(),
                        titleCell.getCellStyle(),
                        (titleCell.getRowHeight() == null ? ExcelStyleConst.ROW_HEIGHT_TITLE_DEFAULT : titleCell.getRowHeight()));
            } else {
                rowIdx = this.createCell(rowIdx, i, workbook, null, titleCell.getCellStyle());
            }
        }

        /** 셀 머지 */
        sheet.addMergedRegion(new CellRangeAddress(rowIdx, rowIdx, 0, headerCellList.size() - 1));

        return rowIdx;
    }


    /**
     * Total Cnt Row 생성 ( 행 전체 )
     *
     * @param rowIdx
     * @param workbook
     * @param headerCellList
     * @param dataCnt
     * @param sheet
     * @return
     * @throws Exception
     */
    private int createTotalCntRow(int rowIdx, SXSSFWorkbook workbook, List<ExcelCellVO> headerCellList, int dataCnt, SXSSFSheet sheet) throws Exception {
        /** 총 데이터 수 생성 */
        int cellStyleNum = ExcelStyleType.CELLSTYLE_TOTAL_CNT.ordinal() + 1;
        for (int i = 0; i < headerCellList.size(); i++) {
            if (i == 0) {
                rowIdx = this.createCell(rowIdx, i, workbook,
                        "Total : " + NumberFormat.getInstance().format(dataCnt),
                        (XSSFCellStyle) workbook.getCellStyleAt(cellStyleNum));
            } else {
                rowIdx = this.createCell(rowIdx, i, workbook, null,
                        (XSSFCellStyle) workbook.getCellStyleAt(cellStyleNum));
            }
        }

        /** 셀 머지 */
        sheet.addMergedRegion(new CellRangeAddress(rowIdx, rowIdx, 0, headerCellList.size() - 1));
        return rowIdx;
    }


    /**
     * Header Row 생성
     *
     * @param rowIdx
     * @param workbook
     * @param headerCellList
     * @return
     * @throws Exception
     */
    private int createHeaderRow(int rowIdx, SXSSFWorkbook workbook, List<ExcelCellVO> headerCellList) throws Exception {
        /** 헤더영역 생성 */
        int cellIdx = 0;

        for (ExcelCellVO headerCell : headerCellList) {
            rowIdx = this.createCell(rowIdx, cellIdx++, workbook, headerCell.getCellValue(),
                    headerCell.getCellStyle(), headerCell.getRowHeight(), headerCell.getCellWidth());
        }
        return rowIdx;
    }


    /**
     * No데이터 ( 데이터가 없을 경우 )
     *
     * @param rowIdx
     * @param workbook
     * @param headerCellList
     * @param dataCnt
     * @param sheet
     * @return
     * @throws Exception
     */
    public int createNoDataRow(int rowIdx, SXSSFWorkbook workbook, List<ExcelCellVO> headerCellList, int dataCnt, SXSSFSheet sheet) throws Exception {
        for (int i = 0; i < headerCellList.size(); i++) {
            if (i == 0) {
                rowIdx = this.createCellDefault(rowIdx, i, workbook,
                        "검색된 결과가 없습니다.",
                        (XSSFCellStyle) workbook.getCellStyleAt(ExcelStyleType.CELLSTYLE_CENTER.ordinal()));
            } else {
                rowIdx = this.createCellDefault(rowIdx, i, workbook, null,
                        (XSSFCellStyle) workbook.getCellStyleAt(ExcelStyleType.CELLSTYLE_CENTER.ordinal()));
            }
        }

        /** 셀 머지 */
        sheet.addMergedRegion(new CellRangeAddress(3, 3, 0, headerCellList.size() - 1));

        return rowIdx;
    }

    /**
     * 셀 생성(기본 레이아웃)
     *
     * @param    rowIdx 줄 인덱스
     * @param    cellIdx 셀 인덱스
     * @param    workbook 워크북
     * @param    cellValue 셀 값
     * @param    cellStyle 셀 스타일
     */
    public int createCell(int rowIdx, int cellIdx, SXSSFWorkbook workbook, Object cellValue,
                                    XSSFCellStyle cellStyle) throws Exception {
        return this.createCellDefault(rowIdx, cellIdx, workbook, cellValue, cellStyle, null);
    }


    /**
     * 셀 생성(기본 레이아웃)
     *
     * @param rowHeight 줄 높이
     * @param    rowIdx 줄 인덱스
     * @param    cellIdx 셀 인덱스
     * @param    workbook 워크북
     * @param    cellValue 셀 값
     * @param    cellStyle 셀 스타일
     */
    public int createCell(int rowIdx, int cellIdx, SXSSFWorkbook workbook, Object cellValue,
                                    XSSFCellStyle cellStyle, Integer rowHeight) throws Exception {
        return this.createCellDefault(rowIdx, cellIdx, workbook, cellValue, cellStyle, rowHeight, null);
    }


    /**
     * 셀 생성(기본 레이아웃)
     *
     * @param rowHeight 줄 높이
     * @param    rowIdx 줄 인덱스
     * @param    cellIdx 셀 인덱스
     * @param    workbook 워크북
     * @param    cellValue 셀 값
     * @param    cellStyle 셀 스타일
     * @param    cellWidth 셀 길이
     */
    public int createCell(int rowIdx, int cellIdx, SXSSFWorkbook workbook, Object cellValue,
                                    XSSFCellStyle cellStyle, Integer rowHeight, Integer cellWidth) throws Exception {
        return this.createCellDefault(rowIdx, cellIdx, workbook, cellValue, cellStyle, rowHeight, cellWidth, null, null);
    }


    /**
     * 셀 생성(기본 레이아웃)
     *
     * @param    rowHeight 줄 높이
     * @param    rowIdx 줄 인덱스
     * @param    cellIdx 셀 인덱스
     * @param    workbook 워크북
     * @param    cellValue 셀 값
     * @param    cellStyle 셀 스타일
     * @param    cellWidth 셀 길이
     * @param    mergeRowCnt 줄합치기 갯수
     * @param    mergeCellCnt 셀합치기 갯수
     */
    public int createCell(int rowIdx, int cellIdx, SXSSFWorkbook workbook, Object cellValue,
                                    XSSFCellStyle cellStyle, Integer rowHeight, Integer cellWidth, Integer mergeRowCnt, Integer mergeCellCnt) {
        SXSSFSheet sheet = workbook.getSheetAt(workbook.getNumberOfSheets() - 1);
        SXSSFRow row = null;

        if (cellIdx == 0) {        /** 첫번째 셀일 경우 줄(row) 생성 및 줄 높이 세팅 */
            /** 줄(row) 생성 */
            row = sheet.createRow(rowIdx++);

            /** 줄 높이 세팅 */
            if (rowHeight == null) {
                row.setHeight(ExcelStyleConst.ROW_HEIGHT_DEFAULT.shortValue());
            } else {
                row.setHeight(rowHeight.shortValue());
            }
        } else {
            row = sheet.getRow(sheet.getLastRowNum());
        }

        /** 셀 생성 */
        SXSSFCell cell = row.createCell(cellIdx);

        /** 셀 값 세팅 */
        if (cellValue != null) {
            if (cellValue instanceof Integer) {
                cell.setCellValue((Integer) cellValue);
            } else if (cellValue instanceof Long) {
                cell.setCellValue((Long) cellValue);
            } else if (cellValue instanceof Double) {
                cell.setCellValue((Double) cellValue);
            } else if (cellValue instanceof Float) {
                cell.setCellValue((Float) cellValue);
            } else if (cellValue instanceof BigDecimal) {
                cell.setCellValue(String.valueOf(cellValue));
            } else if (
                    cellValue instanceof String &&
                            NumberUtils.isCreatable((String) cellValue) &&
                            (
                                    workbook.getCellStyleAt(ExcelStyleType.CELLSTYLE_NUMBER.ordinal() + 1).equals(cellStyle) ||
                                            workbook.getCellStyleAt(ExcelStyleType.CELLSTYLE_CURRENCY.ordinal() + 1).equals(cellStyle)
                            )
            ) {
                cell.setCellValue(Long.parseLong((String) cellValue));
            } else if (
                    cellValue instanceof String &&
                            NumberUtils.isCreatable((String) cellValue) &&
                            (
                                    workbook.getCellStyleAt(ExcelStyleType.CELLSTYLE_NUMBER_DOUBLE.ordinal() + 1).equals(cellStyle) ||
                                            workbook.getCellStyleAt(ExcelStyleType.CELLSTYLE_CURRENCY_DOUBLE.ordinal() + 1)
                                                    .equals(cellStyle)
                            )
            ) {
                cell.setCellValue(Double.parseDouble((String) cellValue));
            } else {
                cell.setCellValue((String) cellValue);
            }
        }

        /** 셀 길이 세팅 */
        if (cellWidth != null) {
            sheet.setColumnWidth(cellIdx, cellWidth);
        }

        /** 셀 스타일 세팅 */
        if (cellStyle == null) {
            cell.setCellStyle(workbook.getCellStyleAt(ExcelStyleType.CELLSTYLE_LEFT.ordinal() + 1));
        } else {
            cell.setCellStyle(cellStyle);
        }

        /** 줄 합치기(행합치기) 세팅 */
        if (mergeRowCnt != null) {
            sheet.addMergedRegion(
                    new CellRangeAddress(row.getRowNum(), row.getRowNum() + mergeRowCnt - 1, cellIdx, cellIdx));
        }

        /** 셀 합치기(열합치기) 세팅 */
        if (mergeCellCnt != null) {
            sheet.addMergedRegion(
                    new CellRangeAddress(row.getRowNum(), row.getRowNum(), cellIdx, cellIdx + mergeCellCnt - 1));
        }

        return rowIdx;
    }


    /**
     * 셀 길이 자동세팅(이 메소드를 호출할 경우 헤더 데이터에 셀 길이를 세팅할 필요 없음)
     *
     * @param    workbook 워크북
     * @param    cellCnt 필드 갯수
     */
    public void setAutoSizingDefault(SXSSFWorkbook workbook, int cellCnt) {
        for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
            workbook.getSheetAt(i).trackAllColumnsForAutoSizing();

            for (int j = 0; j < cellCnt; j++) {
                workbook.getSheetAt(i).autoSizeColumn(j);
            }

            workbook.getSheetAt(i).untrackAllColumnsForAutoSizing();
        }
    }
}
