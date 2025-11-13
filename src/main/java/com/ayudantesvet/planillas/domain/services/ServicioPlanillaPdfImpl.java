package com.ayudantesvet.planillas.domain.services;

import org.apache.pdfbox.pdmodel.PDPageContentStream;
import org.apache.pdfbox.pdmodel.common.PDRectangle;
import org.apache.pdfbox.pdmodel.font.PDType1Font;
import org.apache.pdfbox.pdmodel.font.Standard14Fonts;
import org.springframework.stereotype.Service;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;

@Service
public class ServicioPlanillaPdfImpl implements ServicioPlanillaPdf {
    public void exportarPlanillaAPdf(HttpServletResponse response) throws IOException {
        try {
            //Creación del documento PDF (Sin hojas)
            PDDocument planillaDiaria = new PDDocument();
            // Creación de una página
            PDPage pagina = new PDPage(PDRectangle.A4);
            planillaDiaria.addPage(pagina);

            // Creación del contenido de la página
            PDPageContentStream contentStream = new PDPageContentStream(planillaDiaria, pagina);
            contentStream.beginText();

            // Añadir texto al PDF
            // Configuración de posición del texto
            contentStream.newLineAtOffset(25, 700);
            // Configuración de fuente y tamaño
            contentStream.setFont( new PDType1Font(Standard14Fonts.FontName.HELVETICA), 24 );
            String titulo = "Planilla diaria control";
            contentStream.showText(titulo);

            contentStream.newLineAtOffset(0, -30);
            contentStream.setFont( new PDType1Font(Standard14Fonts.FontName.HELVETICA_BOLD), 12 );
            String texto = "Circular N°4301";
            contentStream.showText(texto);

            contentStream.newLineAtOffset(0, -20);
            contentStream.setFont( new PDType1Font(Standard14Fonts.FontName.TIMES_ROMAN), 8 );
            LocalDateTime hoy = LocalDateTime.now();
            DateTimeFormatter formatoDeFecha = DateTimeFormatter.ofPattern("dd/MM/yyyy");
            String fecha = hoy.format(formatoDeFecha);
            contentStream.showText(fecha);

            contentStream.endText();
            contentStream.close();

            // Configuración de la respuesta HTTP
            response.setContentType(
                    "application/pdf"
            );
            response.setHeader(
                    "Content-Disposition",
                    "attachment; filename=planilla-diaria-de-prueba.pdf"
            );

            planillaDiaria.save(response.getOutputStream());
            planillaDiaria.close();

        } catch (IOException e) {
            throw new IOException("Error al exportar la planilla a PDF");
        }
    }
}
