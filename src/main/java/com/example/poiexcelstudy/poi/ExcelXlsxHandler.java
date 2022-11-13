package com.example.poiexcelstudy.poi;

import com.example.poiexcelstudy.enums.ExcelStyleFontType;
import com.example.poiexcelstudy.enums.ExcelStyleType;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.springframework.web.servlet.view.document.AbstractXlsxStreamingView;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.util.Map;

public class ExcelXlsxHandler extends AbstractXlsxStreamingView {

    @Override
    protected void buildExcelDocument(Map<String, Object> model, Workbook workbook, HttpServletRequest request, HttpServletResponse response) throws Exception {
    }

    @Override
    protected SXSSFWorkbook createWorkbook(Map<String, Object> model, HttpServletRequest request) {
        SXSSFWorkbook workbook = new SXSSFWorkbook();

        ExcelStyleFontType.getWorkBookFontStyle(workbook);
        ExcelStyleType.getWorkBookStyle(workbook);

        return workbook;
    }
}
