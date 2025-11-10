package com.ayudantesvet.planillas.domain.services;

import org.apache.pdfbox.pdmodel.PDPageContentStream;
import org.apache.pdfbox.pdmodel.font.PDType1Font;
import org.apache.pdfbox.pdmodel.font.Standard14Fonts;
import org.springframework.stereotype.Service;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;

@Service
public class ServicioPlanillaPdfImpl implements ServicioPlanillaPdf {
    public void exportarPlanillaAPdf(HttpServletResponse response) throws IOException {
        try {
            PDDocument planillaDiaria = new PDDocument();
            PDPage pagina = new PDPage();
            planillaDiaria.addPage(pagina);

            PDPageContentStream contentStream = new PDPageContentStream( planillaDiaria, pagina);

            contentStream.beginText();

            // Configuración de posición del texto
            contentStream.newLineAtOffset(25, 700);
            // Configuración de fuente y tamaño
            contentStream.setFont( new PDType1Font(Standard14Fonts.FontName.HELVETICA), 12 );

            String text = "Prueba de PDF SSOP";
            contentStream.showText(text);
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
