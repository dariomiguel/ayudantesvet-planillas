package com.ayudantesvet.planillas.domain.services;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

@Service
public class ServicioPlanillaExcelImpl implements ServicioPlanillaExcel {

    @Override
    public void exportarPlanillaAExcel(HttpServletResponse response) throws IOException {

        try {
            Workbook planillaExcelSSOP = new XSSFWorkbook();

            Sheet hojaPOES = crearHoja(planillaExcelSSOP, "POES", 30 , 97);
            Sheet hojaBPM = crearHoja(planillaExcelSSOP, "BPM", 30 , 15);
            Sheet hojaHACCP = crearHoja(planillaExcelSSOP, "HACCP", 30, 1);
            Sheet hojaSPS = crearHoja(planillaExcelSSOP, "SPS", 30, 82);
            Sheet hojaBienestarAnimal = crearHoja(planillaExcelSSOP, "Bienestar Animal", 30 , 1);

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

    private Sheet crearHoja(Workbook planilla, String nombreHoja, int cantDiasDelMes, int cantTareas) {

        Sheet hoja = planilla.createSheet(nombreHoja);

        //Estilos
        CellStyle estiloLineaSuperiorTabla = planilla.createCellStyle();
        estiloLineaSuperiorTabla.setBorderBottom(BorderStyle.THIN);
        estiloLineaSuperiorTabla.setBottomBorderColor(IndexedColors.BLACK.getIndex());

        CellStyle estiloLineaIzquierdaDeTabla = planilla.createCellStyle();
        estiloLineaIzquierdaDeTabla.setBorderRight(BorderStyle.THIN);
        estiloLineaIzquierdaDeTabla.setRightBorderColor(IndexedColors.BLACK.getIndex());

        CellStyle estiloLineaInferiorTabla = planilla.createCellStyle();
        estiloLineaInferiorTabla.setBorderTop(BorderStyle.THIN);
        estiloLineaInferiorTabla.setTopBorderColor(IndexedColors.BLACK.getIndex());

        CellStyle estiloLineaDerechaDeTabla = planilla.createCellStyle();
        estiloLineaDerechaDeTabla.setBorderLeft(BorderStyle.THIN);
        estiloLineaDerechaDeTabla.setLeftBorderColor(IndexedColors.BLACK.getIndex());


        //Valores de configuración de tamaño para la tabla
        int inicioDeTablaFila = 1;
        int inicioDeTablaColumna = 0;

        int cantidadColumnasTotalesParaTabla = (2 + cantDiasDelMes);
        int cantidadFilasTotalesParaTabla = (18 + cantTareas);

        //Creando las filas necesarias para la tabla con sus celdas
        List<Row> filas = new ArrayList<>();
        for (int i = 0; i < cantidadFilasTotalesParaTabla; i++) {
            filas.add(hoja.createRow(i));
            for(int j = 0; j < cantidadColumnasTotalesParaTabla + 2; j++){
                filas.get(i).createCell(j);
            }
        }

        //Pintando las lineas de la tabla
        for(int i = inicioDeTablaColumna; i < cantidadColumnasTotalesParaTabla; i++){
            //linea superior de la tabla
            Cell celdaSuperiorDeTabla = filas.get(inicioDeTablaColumna).getCell(i+1);
            if(celdaSuperiorDeTabla == null){
                celdaSuperiorDeTabla = filas.get(inicioDeTablaColumna).createCell(i+1);
            }
            celdaSuperiorDeTabla.setCellStyle(estiloLineaSuperiorTabla);

            //linea inferior de la tabla
            Cell celdaInferiorDeTabla = filas.get(cantidadFilasTotalesParaTabla-1).getCell(i+1);
            if(celdaInferiorDeTabla == null){
                celdaInferiorDeTabla = filas.get(cantidadFilasTotalesParaTabla-1).createCell(i+1);
            }
            celdaInferiorDeTabla.setCellStyle(estiloLineaInferiorTabla);
        }

        for(int i = inicioDeTablaFila; i < cantidadFilasTotalesParaTabla-1; i++){
            //linea izquierda de la tabla
            Cell celdaLateralIzquierdaDeTabla = filas.get(i).getCell(inicioDeTablaColumna);
            if(celdaLateralIzquierdaDeTabla == null){
                celdaLateralIzquierdaDeTabla = filas.get(i).createCell(inicioDeTablaColumna);
            }
            celdaLateralIzquierdaDeTabla.setCellStyle(estiloLineaIzquierdaDeTabla);

            //linea derecha de la tabla
            Cell celdaLateralDerechaDeTabla = filas.get(i).getCell(cantidadColumnasTotalesParaTabla+1);
            if(celdaLateralDerechaDeTabla == null){
                celdaLateralDerechaDeTabla = filas.get(i).createCell(cantidadColumnasTotalesParaTabla+1);
            }
            celdaLateralDerechaDeTabla.setCellStyle(estiloLineaDerechaDeTabla);
        }

        return hoja;
    }
}
