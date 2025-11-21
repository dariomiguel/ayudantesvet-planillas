package com.ayudantesvet.planillas.domain.services;

import com.ayudantesvet.planillas.infrastructure.excel.BordesDiagonalUtil;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.time.LocalDate;
import java.time.Month;
import java.util.ArrayList;
import java.util.List;

@Service
public class ServicioPlanillaExcelImpl implements ServicioPlanillaExcel {

    @Override
    public void exportarPlanillaAExcel(HttpServletResponse response) throws IOException {

        try {
            Workbook libroExcelSSOP = new XSSFWorkbook();

            List<String> tareasPoes = new ArrayList<>(
                    List.of(
                            "PRE-OPERACIONAL: DESPOSTADA",
                            "DEPÓSITO - GUARDA INSTRUMENTAL DE TRABAJO",
                            "AUSENCIA DE CONDENSACIÓN",
                            "SUPERFICIES DE PALCOS",
                            "CINTA DE HUESOS L.N 1",
                            "CINTA DE CORTES L.N 1",
                            "MESAS DESPOSTE L.N 1",
                            "MESAS DE CHARQUEO L.N 1",
                            "CINTA DE HUESOS L.N 2",
                            "CINTA DE CORTES L.N 2",
                            "MESAS DE DESPOSTE L.N 2",
                            "MESAS DE CHARQUEO L.N 2",
                            "SIERRAS",
                            "MESAS DE ENVASADO PRIMARIO",
                            "BALANZA DE CORTES",
                            "CINTA DE CORTES ENVASADOS",
                            "MESA DE RE - INSPECCIÓN VETERINARIA",
                            "MAQUINA DE VACÍO",
                            "TÚNEL DE TERMOCONTRACCIÓN",
                            "TEMPERATURA DE ESTERILIZADORES",
                            "FUNCIONAMIENTO DE LAVAMANOS",
                            "CONTROL DE CLORO",
                            "ROLDANAS",
                            "HERRAMIENTAS",
                            "PRE-OPERACIONAL: PLAYA DE FAENA",
                            "DEPÓSITO - GUARDA INSTRUMENTAL DE TRABAJO",
                            "AUSENCIA DE CONDENSACIÓN",
                            "CORTADORA DE CUERNOS",
                            "CORTADORA DE PATAS",
                            "DESOLLADORAS",
                            "SELLOS",
                            "TIRABUZÓN - COCODRILOS",
                            "SIERRA DE PECHO",
                            "GANCHOS NORIA DE CABEZAS",
                            "SIERRA DE 1/2 RESES",
                            "BANDEJA DE VISCERAS",
                            "TUBOS VISCERAS",
                            "SUPERFICIES DE PALCOS",
                            "PALCO DE TIPIFICACIÓN - ETIQUETAS.",
                            " BATEAS",
                            "CARROS",
                            "CABINA DE SECADO DE MEDIAS RESES",
                            "CABINA DE ÁCIDO LÁCTICO",
                            "TEMPERATURA DE ESTERILIZADORES",
                            "FUNCIONAMIENTO DE LAVAMANOS",
                            "CONTROL DE CLORO",
                            "ROLDANAS",
                            "HERRAMIENTAS",
                            "PRE-OPERACIONAL: MENUDENCIAS - TRIPERÍA",
                            "DEPÓSITO - GUARDA INSTRUMENTAL DE TRABAJO",
                            "CENTRIFUGA",
                            "MESAS DE TRABAJO",
                            "CANASTOS DE PLASTICOS",
                            "BATEAS, TAMBORES",
                            "TUBOS MENUDENCIAS",
                            "CORRECTA DISPOSICIÓN DE HIELO",
                            "MONDONGUERA",
                            "DESARRADORA",
                            "DESORILLADORA",
                            "CONTROL DE CLORO",
                            "ROLDANAS",
                            "HACHA CORTADORA DE CABEZAS PARA EXTRACCIÓN DE SESOS",
                            "PRE-OPERACIONAL: CUARTEO",
                            "DEPÓSITO - GUARDA INSTRUMENTAL DE TRABAJO",
                            "CONTROL DE CLORO",
                            "CHAPON",
                            "SIERRA DE HUESOS",
                            "ROLDANAS",
                            "FUNCIONAMIENTO DE LAVAMANOS",
                            "TEMPERATURA DE ESTERILIZADORES",
                            "PRE - OPERACIONAL: SALADERO",
                            "HIGIENE PILETONES",
                            "ROLDANAS",
                            "CABINA LAVADO POR ASPERSIÓN",
                            "HERRAMIENTAS",
                            "FUNCIONAMIENTO DE LAVAMANOS",
                            "TEMPERATURA ESTERILIZACIÓN",
                            " ",
                            "OPERATIVIDAD: DESPOSTADA",
                            "HIGIENE DE PISOS - PAREDES Y TECHOS",
                            "HIGIENE DE EQUIPOS Y UTENCILLOS",
                            "TEMPERATURA DE ESTERILIZADORES",
                            "FUNCIONAMIENTO DE LAVAMOS",
                            "TEMPERATURA DE CARNE EN EL INGRESO",
                            "TEMPERATURA DE CORTE A DESPACHAR",
                            "CONTROL DE CLORO",
                            "OPERATIVIDAD: PLAYA DE FAENA",
                            "SINCRONIZACION CABEZA - RES - VISCERAS",
                            "HIGIENE DE EQUIPOS Y UTENCILLOS",
                            "HIGIENE DE PISOS - PAREDES Y TECHOS",
                            "TEMPERATURA DE ESTERILIZADORES",
                            "FUNCIONAMIENTO DE LAVAMOS",
                            "HIGIENE DE PISOS - PAREDES Y TECHOS",
                            "CONTROL DE CLORO",
                            "OPERATIVIDAD: MENUDENCIAS - TRIPERÍA",
                            "TEMPERATURA DE ESTERILIZADORES",
                            "FUNCIONAMIENTO DE LAVAMOS",
                            "TEMPERATURA DE COCINAS",
                            "OPERATIVIDAD: CUARTEO",
                            "TEMPERATURA DE INGRESO DE MATERIA PRIMA",
                            "TEMPERATURA DE ESTERILIZADORES",
                            "FUNCIONAMIENTO DE LAVAMOS",
                            "OPERATIVIDAD: SALADERO",
                            "TEMPERATUA ESTERILIZADORES",
                            "FUNCIONAMIENTO DE LAVAMOS",
                            "TEMPERATURA DE SALA ( >10° )",
                            "TEMPERATURA DE AGUA DE PILETONES"
                            )
            );

            int numeroMesPrueba = 11;

            Sheet hojaPOES = crearHoja(libroExcelSSOP, "POES", numeroMesPrueba , tareasPoes);

            /*
            Sheet hojaBPM = crearHoja(libroExcelSSOP, "BPM",numeroMesPrueba , tareasPoes);
            Sheet hojaHACCP = crearHoja(libroExcelSSOP, "HACCP",numeroMesPrueba, tareasPoes);
            Sheet hojaSPS = crearHoja(libroExcelSSOP, "SPS", numeroMesPrueba, tareasPoes);
            Sheet hojaBienestarAnimal = crearHoja(libroExcelSSOP, "Bienestar Animal",numeroMesPrueba, tareasPoes);*/

            // Configuración de la respuesta HTTP
            response.setContentType(
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            );
            response.setHeader(
                    "Content-Disposition",
                    "attachment; filename=planilla-mensual-de-prueba.xlsx"
            );

            libroExcelSSOP.write(response.getOutputStream());
            libroExcelSSOP.close();

        }catch (IOException e ){
            throw new IOException("Error al exportar la planilla a Excel");
        }
    }

    private Sheet crearHoja(Workbook libroExcel, String programa, int nroDelMes, List<String> listaDeTareas) {

        Sheet hoja = libroExcel.createSheet(programa);

        //Convertir cantidad total de dias del mes
        LocalDate fechaPrueba = LocalDate.of(2025,nroDelMes, 1);
        int cantDiasDelMes = fechaPrueba.lengthOfMonth();

        Month mesSinFormato = fechaPrueba.getMonth();
        String mes = mesEnEspaniol(mesSinFormato);
        int anio = fechaPrueba.getYear();

        //Creación de la tabla
        //-----------------------------
        //-----------------------------

        final int FILA_INICIAL_TABLA = 0;
        final int PRIMERA_FILA_TABLA = 1;
        final int COLUMNA_INICIO_TABLA = 0;

        final int OFFSET_FILAS = 8;
        final int OFFSET_COLUMNAS = 2;

        final int SEPARADOR_FILAS_OPE_PREOPE = 82;

        //Valores de configuración de tamaño para la tabla
        int cantTareas = listaDeTareas.size();
        int cantidadFilasTotalesParaTabla = (OFFSET_FILAS + cantTareas);

        final int ULTIMA_FILA_TABLA  = cantidadFilasTotalesParaTabla - 1;

        final int PRIMERA_COLUMNA_TABLA = COLUMNA_INICIO_TABLA + 1;
        final int ULTIMA_COLUMNA_TABLA  = (OFFSET_COLUMNAS + cantDiasDelMes);
        final int COLUMNA_BORDE_DERECHA = ULTIMA_COLUMNA_TABLA + 1;

        // Convención de posiciones de la tabla
        final int FILA_TITULO_DIAS = 2;       // fila de los números de días
        final int FILA_BARRA_GRIS_SUPERIOR = 3;
        final int FILA_BARRA_GRIS_INFERIOR = 4;
        final int FILA_PRIMERA_TAREA = 5;

        final int COLUMNA_TAREAS = 1;         // "ÁREAS DETERMINADAS"
        final int COLUMNA_FRECUENCIA = 2;     // columna "Frecuencia"
        final int COLUMNA_PRIMER_DIA = 3;     // primera columna de días (1..31)

        //Regiones mergeadas
        //-----------------------------

        //Mergeado celdas para el titulo de la tabla
        hoja.addMergedRegion(new CellRangeAddress(
                1, //first row (0-based)
                2, //last row  (0-based)
                1, //first column (0-based)
                1  //last column  (0-based)
        ));

        //Mergeado celdas para el título de frecuencia
        hoja.addMergedRegion(new CellRangeAddress(1,2,2,2));

        //Mergeado celdas para el título del mes
        hoja.addMergedRegion(new CellRangeAddress(1, 1, 3, ULTIMA_COLUMNA_TABLA));


        //Estilos
        //-----------------------------

        //Colores

        XSSFColor grisExcel = new XSSFColor(new java.awt.Color(166, 166, 166), null);
        XSSFColor amarillo = new XSSFColor(new java.awt.Color(255, 255, 0), null);

        //Bordes para tabla
        CellStyle estiloLineaSuperiorTabla = libroExcel.createCellStyle();
        estiloLineaSuperiorTabla.setBorderBottom(BorderStyle.THIN);
        estiloLineaSuperiorTabla.setBottomBorderColor(IndexedColors.BLACK.getIndex());

        CellStyle estiloLineaIzquierdaDeTabla = libroExcel.createCellStyle();
        estiloLineaIzquierdaDeTabla.setBorderRight(BorderStyle.THIN);
        estiloLineaIzquierdaDeTabla.setRightBorderColor(IndexedColors.BLACK.getIndex());

        CellStyle estiloLineaInferiorTabla = libroExcel.createCellStyle();
        estiloLineaInferiorTabla.setBorderTop(BorderStyle.THIN);
        estiloLineaInferiorTabla.setTopBorderColor(IndexedColors.BLACK.getIndex());

        CellStyle estiloLineaDerechaDeTabla = libroExcel.createCellStyle();
        estiloLineaDerechaDeTabla.setBorderLeft(BorderStyle.THIN);
        estiloLineaDerechaDeTabla.setLeftBorderColor(IndexedColors.BLACK.getIndex());

        //Estilo celda con todos los bordes fuente tamaño 9 negrita
        CellStyle estiloCeldaConTodosLosBordes = libroExcel.createCellStyle();
        estiloCeldaConTodosLosBordes.setBorderBottom(BorderStyle.THIN);
        estiloCeldaConTodosLosBordes.setBottomBorderColor(IndexedColors.BLACK.getIndex());
        estiloCeldaConTodosLosBordes.setBorderLeft(BorderStyle.THIN);
        estiloCeldaConTodosLosBordes.setLeftBorderColor(IndexedColors.BLACK.getIndex());
        estiloCeldaConTodosLosBordes.setBorderRight(BorderStyle.THIN);
        estiloCeldaConTodosLosBordes.setRightBorderColor(IndexedColors.BLACK.getIndex());
        estiloCeldaConTodosLosBordes.setBorderTop(BorderStyle.THIN);
        estiloCeldaConTodosLosBordes.setTopBorderColor(IndexedColors.BLACK.getIndex());

        //Estilo celda con todos los bordes fuente tamaño 11 negrita
        CellStyle estiloCeldaDias = libroExcel.createCellStyle();
        estiloCeldaDias.cloneStyleFrom(estiloCeldaConTodosLosBordes);

        //Estilo titulo de mes
        CellStyle estiloCeldaTituloDelMes = libroExcel.createCellStyle();
        estiloCeldaTituloDelMes.cloneStyleFrom(estiloCeldaConTodosLosBordes);

        //Estilo fondo gris
        CellStyle estiloFondoGris = libroExcel.createCellStyle();
        estiloFondoGris.setFillForegroundColor(grisExcel);
        estiloFondoGris.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        //Estilo fondo nombre de programa
        CellStyle estiloFondoGrisNombrePrograma = libroExcel.createCellStyle();
        estiloFondoGrisNombrePrograma.cloneStyleFrom(estiloFondoGris);

        //Estilo fondo descripcion de programa
        CellStyle estiloFondoGrisDescripcionDePrograma = libroExcel.createCellStyle();
        estiloFondoGrisDescripcionDePrograma.cloneStyleFrom(estiloFondoGris);

        //Fondo amarillo
        CellStyle estiloFondoAmarillo = libroExcel.createCellStyle();
        estiloFondoAmarillo.setFillForegroundColor(amarillo);
        estiloFondoAmarillo.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        //Estilo fondo amarillo títulos pre-operacionales y operativos
        CellStyle estiloFondoAmarilloTitulos = libroExcel.createCellStyle();
        estiloFondoAmarilloTitulos.cloneStyleFrom(estiloCeldaConTodosLosBordes);
        estiloFondoAmarilloTitulos.setFillForegroundColor(amarillo);
        estiloFondoAmarilloTitulos.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        //Estilo líneas cruzadas para fines de semana
        CellStyle estiloDiagonalCruces  = libroExcel.createCellStyle();
        estiloDiagonalCruces.cloneStyleFrom(estiloCeldaConTodosLosBordes);
/*        estiloDiagonalCruces.setFillPattern(FillPatternType.ALT_BARS);*/
        BordesDiagonalUtil.aplicarDiagonales(libroExcel, estiloDiagonalCruces, BorderStyle.THIN);

        //Fuentes
        //-----------------------------

        //Creación de fuentes
        Font fuenteNegrita9 = crearFuente(libroExcel, 9, true);
        Font fuenteNegrita11 = crearFuente(libroExcel, 11, true);
        Font fuenteNegrita12 = crearFuente(libroExcel, 12, true);
        Font fuenteNegrita18 = crearFuente(libroExcel, 18, true);

        //Estilo para celdas con todos los bordes
        estiloCeldaConTodosLosBordes.setFont(fuenteNegrita9);

        //Estilo para celdas de días
        estiloCeldaDias.setFont(fuenteNegrita11);
        estiloCeldaDias.setAlignment(HorizontalAlignment.CENTER);
        estiloCeldaDias.setVerticalAlignment(VerticalAlignment.CENTER);

        //Estilo para el título del mes
        estiloCeldaTituloDelMes.setFont(fuenteNegrita18);
        estiloCeldaTituloDelMes.setAlignment(HorizontalAlignment.CENTER);

        //Estilo fondo gris (Nombre de programas)
        estiloFondoGrisNombrePrograma.setFont(fuenteNegrita12);
        estiloFondoGrisNombrePrograma.setVerticalAlignment(VerticalAlignment.CENTER);

        //Fuente fondo gris (Descripción de programas)
        estiloFondoGrisDescripcionDePrograma.setFont(fuenteNegrita9);
        estiloFondoGrisDescripcionDePrograma.setVerticalAlignment(VerticalAlignment.CENTER);

        //Fuente fondo amarillo títulos pre-operacionales y operativos
        estiloFondoAmarilloTitulos.setFont(fuenteNegrita9);

        //Aplicando bordes
        //-----------------------------

        //Pintando las lineas bordes de la tabla
        for(int col = PRIMERA_COLUMNA_TABLA; col <= ULTIMA_COLUMNA_TABLA; col++){
            //linea superior de la tabla
            Cell celdaSuperiorDeTabla = obtenerOCrearCelda(hoja, FILA_INICIAL_TABLA, col);
            celdaSuperiorDeTabla.setCellStyle(estiloLineaSuperiorTabla);

            //linea inferior de la tabla
            Cell celdaInferiorDeTabla = obtenerOCrearCelda(hoja, ULTIMA_FILA_TABLA, col);
            celdaInferiorDeTabla.setCellStyle(estiloLineaInferiorTabla);
        }

        for(int fila = PRIMERA_FILA_TABLA; fila < ULTIMA_FILA_TABLA; fila++){
            //linea izquierda de la tabla
            Cell celdaLateralIzquierdaDeTabla = obtenerOCrearCelda(hoja, fila, COLUMNA_INICIO_TABLA);
            celdaLateralIzquierdaDeTabla.setCellStyle(estiloLineaIzquierdaDeTabla);

            //linea derecha de la tabla
            Cell celdaLateralDerechaDeTabla = obtenerOCrearCelda(hoja, fila, COLUMNA_BORDE_DERECHA);
            celdaLateralDerechaDeTabla.setCellStyle(estiloLineaDerechaDeTabla);
        }

        //Datos para completar tabla
        //-----------------------------

        //Título de tareas
        Cell celdaTituloTareas = obtenerOCrearCelda(hoja,1,1);
        celdaTituloTareas.setCellValue("ÁREAS DETERMINADAS");
        celdaTituloTareas.setCellStyle(estiloCeldaDias);
        Cell celdaTituloTareasMerge = obtenerOCrearCelda(hoja,2,1);
        celdaTituloTareasMerge.setCellStyle(estiloCeldaDias);

        //Título nombre de programa
        Cell celdaTituloPrograma = obtenerOCrearCelda(hoja,3,1);
        celdaTituloPrograma.setCellValue(programa);
        celdaTituloPrograma.setCellStyle(estiloFondoGrisNombrePrograma);

        //Descripción de programa
        Cell celdaDescripcionDePrograma = obtenerOCrearCelda(hoja,4,1);
        celdaDescripcionDePrograma.setCellValue("Todo asociado con los puntos del proceso más probables de contaminación");
        celdaDescripcionDePrograma.setCellStyle(estiloFondoGrisDescripcionDePrograma);

        //Título de la planilla
        Cell celdaTituloFrecuencia = obtenerOCrearCelda(hoja,1,2);
        celdaTituloFrecuencia.setCellValue("Frecuencia");
        celdaTituloFrecuencia.setCellStyle(estiloCeldaDias);
        Cell celdaTituloFrecuenciaMerge = obtenerOCrearCelda(hoja,2,2);
        celdaTituloFrecuenciaMerge.setCellStyle(estiloCeldaDias);

        //Título mes de la planilla
        Cell celdaTituloMesDeLaPlanilla = obtenerOCrearCelda(hoja,1,3);
        celdaTituloMesDeLaPlanilla.setCellValue("MES DE " + mes + " " + anio);
        celdaTituloMesDeLaPlanilla.setCellStyle(estiloCeldaTituloDelMes);

        //Llenando la fila de los números de dias
        for(int col = 0; col < cantDiasDelMes; col++){
            Cell nroDeDia = obtenerOCrearCelda(hoja, FILA_TITULO_DIAS,COLUMNA_PRIMER_DIA + col );
            nroDeDia.setCellValue(col+1);
            nroDeDia.setCellStyle(estiloCeldaDias);
        }

        //Barra gris separadora debajo de números de días
        for(int col = COLUMNA_TAREAS; col < ULTIMA_COLUMNA_TABLA; col++){
            Cell celdaBarraGris = obtenerOCrearCelda(hoja, FILA_BARRA_GRIS_SUPERIOR, COLUMNA_TAREAS + col);
            Cell celdaBarraGrisMerge = obtenerOCrearCelda(hoja, FILA_BARRA_GRIS_INFERIOR, COLUMNA_TAREAS + col);
            celdaBarraGris.setCellStyle(estiloFondoGris);
            celdaBarraGrisMerge.setCellStyle(estiloFondoGris);
        }

        //Barra gris separadora debajo de números de días
        for(int fila = FILA_PRIMERA_TAREA; fila < ULTIMA_FILA_TABLA; fila++){
            Cell celdaBarraGris = obtenerOCrearCelda(hoja, fila, COLUMNA_FRECUENCIA);
            celdaBarraGris.setCellStyle(estiloFondoGris);
        }

        //Llenando la primera columna con las tareas
        for(int fila = 0; fila < cantTareas; fila++){
            Cell celdaDeTarea = obtenerOCrearCelda(hoja,FILA_PRIMERA_TAREA + fila, COLUMNA_TAREAS );
            celdaDeTarea.setCellValue(listaDeTareas.get(fila));
            celdaDeTarea.setCellStyle(estiloCeldaConTodosLosBordes);
        }

        //Cuadros de tareas por día
        for(int fila = 0; fila < cantTareas; fila++){
            for(int col = 0; col < cantDiasDelMes; col++){
                Cell celdaCuadroDeTareaPorDia = obtenerOCrearCelda(hoja,FILA_PRIMERA_TAREA + fila, COLUMNA_PRIMER_DIA + col);
                celdaCuadroDeTareaPorDia.setCellStyle(estiloCeldaConTodosLosBordes);
            }
        }

        //Resaltar bordes cruzados para los fines de semana
        for(int fila = 0; fila < cantTareas; fila++){
            for(int col = 0; col < cantDiasDelMes; col++){
                LocalDate fechaActual = LocalDate.of(anio, nroDelMes, col+1);
                boolean esSabadoODomingo = fechaActual.getDayOfWeek().getValue() == 6 ||
                        fechaActual.getDayOfWeek().getValue() == 7;

                if(esSabadoODomingo){
                    Cell celdaFinDeSemana = obtenerOCrearCelda(hoja,FILA_PRIMERA_TAREA + fila,COLUMNA_PRIMER_DIA + col);
                    celdaFinDeSemana.setCellStyle(estiloDiagonalCruces);
                }
            }
        }

        //Resaltar en amarillo las filas de pre-operacionales y operativos
        for(int fila = 0; fila < cantTareas; fila++){
            String tareaActual = listaDeTareas.get(fila);
            boolean esFilaDeTituloPreoperacionalUOperacional = tareaActual.contains("OPERACIONAL:") ||
                    tareaActual.contains("OPERATIVIDAD:");

            if(esFilaDeTituloPreoperacionalUOperacional){
                Cell celdaDeTarea = obtenerOCrearCelda(hoja,FILA_PRIMERA_TAREA + fila, COLUMNA_TAREAS);
                celdaDeTarea.setCellStyle(estiloFondoAmarilloTitulos);

                for(int col = COLUMNA_PRIMER_DIA; col < COLUMNA_BORDE_DERECHA; col++){
                    Cell celdaAmarilla = obtenerOCrearCelda(hoja,FILA_PRIMERA_TAREA + fila, col);
                    celdaAmarilla.setCellStyle(estiloFondoAmarillo);
                }
            }
        }

        //Pintar de gris separación entre filas de pre-operacionales y operativos
        for(int col = 0; col < ULTIMA_COLUMNA_TABLA; col++){
            Cell celdaSeparadora = obtenerOCrearCelda(hoja, SEPARADOR_FILAS_OPE_PREOPE, COLUMNA_TAREAS + col);
            celdaSeparadora.setCellStyle(estiloFondoGris);
        }

        //Ajustes de ancho y alto de celdas
        //---------------------------------

        //Ajuste de ancho de la primera columna
        hoja.setColumnWidth(0,(int) ((4.8) * 256));

        //Ajuste de ancho de columna tareas
        hoja.setColumnWidth(1,55*256);

        //Ajuste de ancho de frecuencias
        hoja.setColumnWidth(2, 15 * 256);

        //Ajuste de ancho para dias del mes
        for(int col = 0; col < cantDiasDelMes; col++){
            hoja.setColumnWidth(COLUMNA_PRIMER_DIA + col, 256*5);
        }

        //Ajustar alto de filas
        hoja.getRow(4).setHeightInPoints(23);
        return hoja;
    }

    private Cell obtenerOCrearCelda(Sheet hoja, int indiceFila, int indiceCol) {
        Row fila = hoja.getRow(indiceFila);
        if (fila == null) fila = hoja.createRow(indiceFila);

        Cell celda = fila.getCell(indiceCol);
        if (celda == null) celda = fila.createCell(indiceCol);

        return celda;
    }

    private String mesEnEspaniol(Month mes){
        switch (mes){
            case JANUARY:
                return "ENERO";
            case FEBRUARY:
                return "FEBRERO";
            case MARCH:
                return "MARZO";
            case APRIL:
                return "ABRIL";
            case MAY:
                return "MAYO";
            case JUNE:
                return "JUNIO";
            case JULY:
                return "JULIO";
            case AUGUST:
                return "AGOSTO";
            case SEPTEMBER:
                return "SEPTIEMBRE";
            case OCTOBER:
                return "OCTUBRE";
            case NOVEMBER:
                return "NOVIEMBRE";
            case DECEMBER:
                return "DICIEMBRE";
            default:
                return "";
        }
    }

    private Font crearFuente(Workbook libroExcel, int tamanio, boolean negrita) {
        Font font = libroExcel.createFont();
        font.setFontName("Calibri");
        font.setFontHeightInPoints((short) tamanio);
        font.setBold(negrita);
        return font;
    }
}
