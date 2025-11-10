package com.ayudantesvet.planillas.domain.services;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;

@Service
public class ServicioPlanillaExcelImpl implements ServicioPlanillaExcel {

    @Override
    public void exportarPlanillaAExcel(HttpServletResponse response) throws IOException {

        try {
            Workbook planillaExcelSSOP = new XSSFWorkbook();
            Sheet hojaPoes = planillaExcelSSOP.createSheet("POES");
            Row filaEncabezado = hojaPoes.createRow(0);
            Cell celdaPrueba = filaEncabezado.createCell(0);

            celdaPrueba.setCellValue("Prueba de Excel SSOP");


            // Configuración de la respuesta HTTP
            response.setContentType(
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            );
            response.setHeader(
                    "Content-Disposition",
                    "attachment; filename=planilla-mensual-de-prueba.xlsx"
            );

            planillaExcelSSOP.write(response.getOutputStream());
            planillaExcelSSOP.close();

        }catch (IOException e ){
            throw new IOException("Error al exportar la planilla a Excel");
        }


    }
}
