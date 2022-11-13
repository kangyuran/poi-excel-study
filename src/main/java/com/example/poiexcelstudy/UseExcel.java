package com.example.poiexcelstudy;

import java.text.NumberFormat;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import com.example.poiexcelstudy.enums.ExcelStyleType;
import com.example.poiexcelstudy.poi.ExcelProvider;
import com.example.poiexcelstudy.poi.ExcelXlsxHandler;
import kr.co.zlgoon.branch.admin.system.base.BaseExcel;
import kr.co.zlgoon.branch.admin.system.base.ExcelCellVO;
import kr.co.zlgoon.branch.core.master.adjust.monthly.vo.AdjustMonthlyListResVO;
import kr.co.zlgoon.branch.core.master.adjust.monthly.vo.AdjustMonthlyReqVO;
import kr.co.zlgoon.branch.core.master.adjust.monthly.vo.AdjustMonthlySumResVO;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.springframework.stereotype.Component;

@Component
public class UseExcel extends ExcelXlsxHandler {

    @Override
    protected void buildExcelDocument(Map<String, Object> model, Workbook workbook, HttpServletRequest request, HttpServletResponse response) {

        ExcelProvider excelProvider = new ExcelProvider();


        /** 타이틀 */
        String title = "부가세신고용합계표";
        excelProvider.createSheet((SXSSFWorkbook) workbook, title);


        /** 헤더 데이터 세팅{셀 값, 셀 스타일, 줄 높이, 셀 길이} --> 셀 길이 자동세팅을 할 경우 셀 길이를 전달할 필요 없음 */
        List<ExcelCellVO> headerCellList = new ArrayList<>();

        headerCellList.add(new ExcelCellVO("정산기간", (XSSFCellStyle) workbook.getCellStyleAt(ExcelStyleType.CELLSTYLE_HEADER), 300, 256 * 40));
        headerCellList.add(new ExcelCellVO("쿠폰금액", (XSSFCellStyle) workbook.getCellStyleAt(BaseExcel.CELLSTYLE_HEADER), 300, 256 * 15));
        headerCellList.add(new ExcelCellVO("교환수량", (XSSFCellStyle) workbook.getCellStyleAt(BaseExcel.CELLSTYLE_HEADER), 300, 256 * 15));
        headerCellList.add(new ExcelCellVO("사용금액(A)", (XSSFCellStyle) workbook.getCellStyleAt(BaseExcel.CELLSTYLE_HEADER), 300, 256 * 15));
        headerCellList.add(new ExcelCellVO("수수료(공급가액)", (XSSFCellStyle) workbook.getCellStyleAt(BaseExcel.CELLSTYLE_HEADER), 300, 256 * 15));
        headerCellList.add(new ExcelCellVO("수수료(부가세)", (XSSFCellStyle) workbook.getCellStyleAt(BaseExcel.CELLSTYLE_HEADER), 300, 256 * 15));
        headerCellList.add(new ExcelCellVO("수수료(합계)(B)", (XSSFCellStyle) workbook.getCellStyleAt(BaseExcel.CELLSTYLE_HEADER), 300, 256 * 15));
        headerCellList.add(new ExcelCellVO("송금액(A)-(B)", (XSSFCellStyle) workbook.getCellStyleAt(BaseExcel.CELLSTYLE_HEADER), 300, 256 * 15));
        headerCellList.add(new ExcelCellVO("비고", (XSSFCellStyle) workbook.getCellStyleAt(BaseExcel.CELLSTYLE_HEADER), 300, 256 * 40));





        /** 타이틀 데이터 세팅{셀 값, 셀 스타일, 줄 높이, 셀 길이} */
        rowIdx = createTitleRow(0, workbook, titleCell, headerCellList, sheet);
        rowIdx = createTotalCntRow(rowIdx, workbook, headerCellList, dataCnt, sheet);
        ExcelCellVO titleCell = new ExcelCellVO(title, (XSSFCellStyle) workbook.getCellStyleAt(BaseExcel.CELLSTYLE_TITLE));


        rowIdx = createHeaderRow(rowIdx, workbook, headerCellList);

        /** DB 데이터 리스트 */

        /** 줄 인덱스 */
        int rowIdx = 0;

        /** 타이틀, 헤더, 데이터가 없을 경우 등 데이터 영역을 제외한 시트 내용 생성(기본 레이아웃) */
        rowIdx = excelProvider.createSheetDefault(rowIdx, (SXSSFWorkbook) workbook, titleCell, headerCellList, list.size());
        if (dataCnt == 0) {
            rowIdx = createNoDataRow(rowIdx, workbook, headerCellList, dataCnt, sheet);
        }


        int cellIdx = 0;
        rowIdx = super.createCellDefault(rowIdx, cellIdx++, (SXSSFWorkbook) workbook, "합 계", (XSSFCellStyle) workbook.getCellStyleAt(BaseExcel.CELLSTYLE_TOTAL_SUM));
        rowIdx = super.createCellDefault(rowIdx, cellIdx++, (SXSSFWorkbook) workbook, sum.getSumGoodsPrice(), (XSSFCellStyle) workbook.getCellStyleAt(BaseExcel.CELLSTYLE_CURRENCY));
        rowIdx = super.createCellDefault(rowIdx, cellIdx++, (SXSSFWorkbook) workbook, sum.getSumUseCnt(), (XSSFCellStyle) workbook.getCellStyleAt(BaseExcel.CELLSTYLE_CURRENCY));
        rowIdx = super.createCellDefault(rowIdx, cellIdx++, (SXSSFWorkbook) workbook, sum.getSumAdjPrice(), (XSSFCellStyle) workbook.getCellStyleAt(BaseExcel.CELLSTYLE_CURRENCY));
        rowIdx = super.createCellDefault(rowIdx, cellIdx++, (SXSSFWorkbook) workbook, sum.getSumAgencyFeeSupply(), (XSSFCellStyle) workbook.getCellStyleAt(BaseExcel.CELLSTYLE_CURRENCY));
        rowIdx = super.createCellDefault(rowIdx, cellIdx++, (SXSSFWorkbook) workbook, sum.getSumAgencyFeeAdd(), (XSSFCellStyle) workbook.getCellStyleAt(BaseExcel.CELLSTYLE_CURRENCY));
        rowIdx = super.createCellDefault(rowIdx, cellIdx++, (SXSSFWorkbook) workbook, sum.getSumAgencyFee(), (XSSFCellStyle) workbook.getCellStyleAt(BaseExcel.CELLSTYLE_CURRENCY));
        rowIdx = super.createCellDefault(rowIdx, cellIdx++, (SXSSFWorkbook) workbook, sum.getSumAgencySendPrice(), (XSSFCellStyle) workbook.getCellStyleAt(BaseExcel.CELLSTYLE_CURRENCY));
        rowIdx = super.createCellDefault(rowIdx, cellIdx++, (SXSSFWorkbook) workbook, "", (XSSFCellStyle) workbook.getCellStyleAt(BaseExcel.CELLSTYLE_CURRENCY));

        /** 데이터 영역 생성(셀 인덱스, 워크북, 셀 값, 셀 스타일, 줄 높이 전달) --> 셀 길이는 헤더(또는 셀길이 자동세팅)에서 세팅하므로 전달할 필요 없음 */
        for (AdjustMonthlyListResVO data : list) {
            cellIdx = 0;

            String textPeriod =
                    data.getTargetYymm().substring(0, 4) + "년 " + data.getTargetYymm().substring(4, 6) + "월 ("
                            + data.getAdjustPeriodStrDatetime().format(DateTimeFormatter.ISO_LOCAL_DATE) + "~"
                            + data.getAdjustPeriodEndDatetime().format(DateTimeFormatter.ISO_LOCAL_DATE) + ")";

            String textNote = "";
            if (data.getSupportPrice() > 0) {
                textNote = NumberFormat.getInstance().format(data.getSupportPrice()) + "(전체지원금)," +
                        NumberFormat.getInstance().format(data.getHeadquartersSupportPrice()) + "(본사지원금)," +
                        NumberFormat.getInstance().format(data.getZlgoonSupportPrice()) + "(즐거운지원금)";
            }

            rowIdx = super.createCellDefault(rowIdx, cellIdx++, (SXSSFWorkbook) workbook, textPeriod, (XSSFCellStyle) workbook.getCellStyleAt(BaseExcel.CELLSTYLE_CENTER));
            rowIdx = super.createCellDefault(rowIdx, cellIdx++, (SXSSFWorkbook) workbook, data.getGoodsPrice(), (XSSFCellStyle) workbook.getCellStyleAt(BaseExcel.CELLSTYLE_CURRENCY));
            rowIdx = super.createCellDefault(rowIdx, cellIdx++, (SXSSFWorkbook) workbook, data.getUseCnt(), (XSSFCellStyle) workbook.getCellStyleAt(BaseExcel.CELLSTYLE_CURRENCY));
            rowIdx = super.createCellDefault(rowIdx, cellIdx++, (SXSSFWorkbook) workbook, data.getAdjPrice(), (XSSFCellStyle) workbook.getCellStyleAt(BaseExcel.CELLSTYLE_CURRENCY));
            rowIdx = super.createCellDefault(rowIdx, cellIdx++, (SXSSFWorkbook) workbook, data.getAgencyFeeSupply(), (XSSFCellStyle) workbook.getCellStyleAt(BaseExcel.CELLSTYLE_CURRENCY));
            rowIdx = super.createCellDefault(rowIdx, cellIdx++, (SXSSFWorkbook) workbook, data.getAgencyFeeAdd(), (XSSFCellStyle) workbook.getCellStyleAt(BaseExcel.CELLSTYLE_CURRENCY));
            rowIdx = super.createCellDefault(rowIdx, cellIdx++, (SXSSFWorkbook) workbook, data.getAgencyFee(), (XSSFCellStyle) workbook.getCellStyleAt(BaseExcel.CELLSTYLE_CURRENCY));
            rowIdx = super.createCellDefault(rowIdx, cellIdx++, (SXSSFWorkbook) workbook, data.getAgencySendPrice(), (XSSFCellStyle) workbook.getCellStyleAt(BaseExcel.CELLSTYLE_CURRENCY));
            rowIdx = super.createCellDefault(rowIdx, cellIdx++, (SXSSFWorkbook) workbook, textNote, (XSSFCellStyle) workbook.getCellStyleAt(BaseExcel.CELLSTYLE_CENTER));
        }

        /** 엑셀 파일 명 및 ResponseHeader 세팅 */
        super.setExcelFile(title + "_" + LocalDateTime.now().format(DateTimeFormatter.BASIC_ISO_DATE) + ".xlsx", request, response);
    }
}