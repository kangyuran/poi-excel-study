package com.example.poiexcelstudy.enums;

import lombok.Getter;
import org.apache.poi.ss.usermodel.BorderStyle;

@Getter
public enum ExcelStyleBorderType {

    BORDER_NONE_ALL(BorderStyle.NONE, BorderStyle.NONE, BorderStyle.NONE, BorderStyle.NONE),
    BORDER_THIN_ALL(BorderStyle.THIN, BorderStyle.THIN, BorderStyle.THIN, BorderStyle.THIN),
    BORDER_THIN_TOP_BOTTOM(BorderStyle.NONE, BorderStyle.NONE, BorderStyle.THIN, BorderStyle.THIN),
    ;


    private final BorderStyle borderLeft;
    private final BorderStyle borderRight;
    private final BorderStyle borderTop;
    private final BorderStyle borderBottom;


    ExcelStyleBorderType(BorderStyle borderLeft, BorderStyle borderRight, BorderStyle borderTop, BorderStyle borderBottom) {
        this.borderLeft = borderLeft;
        this.borderRight = borderRight;
        this.borderTop = borderTop;
        this.borderBottom = borderBottom;
    }
}
