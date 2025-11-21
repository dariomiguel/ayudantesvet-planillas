package com.ayudantesvet.planillas.infrastructure.excel;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.extensions.XSSFCellBorder;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.model.ThemesTable;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTXf;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTBorder;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTBorderPr;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STBorderStyle;

public class BordesDiagonalUtil {

    private static CTBorder getCTBorder(StylesTable stylesSource, CTXf cellXf) {
        CTBorder ct;
        if (cellXf.getApplyBorder()) {
            int idx = (int) cellXf.getBorderId();
            XSSFCellBorder cf = stylesSource.getBorderAt(idx);
            ct = (CTBorder) cf.getCTBorder().copy();
        } else {
            ct = CTBorder.Factory.newInstance();
        }
        return ct;
    }

    // 👉 OJO: ahora recibe también el Workbook
    public static void aplicarDiagonales(Workbook workbook,
                                         CellStyle style,
                                         BorderStyle borderStyle) {

        if (!(workbook instanceof XSSFWorkbook)) {
            // Solo soportamos XSSFWorkbook (.xlsx)
            return;
        }
        if (!(style instanceof XSSFCellStyle)) {
            // El estilo tiene que ser XSSFCellStyle
            return;
        }

        XSSFWorkbook xssfWb = (XSSFWorkbook) workbook;
        StylesTable stylesSource = xssfWb.getStylesSource();
        ThemesTable theme = stylesSource.getTheme();

        XSSFCellStyle xssfStyle = (XSSFCellStyle) style;
        CTXf cellXf = xssfStyle.getCoreXf();

        CTBorder ct = getCTBorder(stylesSource, cellXf);
        CTBorderPr pr = ct.isSetDiagonal() ? ct.getDiagonal() : ct.addNewDiagonal();

        if (borderStyle == BorderStyle.NONE) {
            ct.unsetDiagonal();
        } else {
            ct.setDiagonalDown(true);
            ct.setDiagonalUp(true);
            pr.setStyle(STBorderStyle.Enum.forInt(borderStyle.getCode() + 1));
        }

        int idx = stylesSource.putBorder(
                new XSSFCellBorder(ct, theme, stylesSource.getIndexedColors())
        );
        cellXf.setBorderId(idx);
        cellXf.setApplyBorder(true);
    }
}
