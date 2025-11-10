package com.ayudantesvet.planillas.web;

import com.ayudantesvet.planillas.domain.services.ServicioPlanillaExcel;
import com.ayudantesvet.planillas.domain.services.ServicioPlanillaPdf;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;

@Controller
public class ControladorPlanillas {

    private final ServicioPlanillaExcel servicioPlanillaExcel;
    private final ServicioPlanillaPdf servicioPlanillaPdf;

    public ControladorPlanillas(ServicioPlanillaExcel servicioPlanillaExcel, ServicioPlanillaPdf servicioPlanillaPdf) {
        this.servicioPlanillaPdf = servicioPlanillaPdf;
        this.servicioPlanillaExcel = servicioPlanillaExcel;
    }

    @GetMapping("/exportar/planillassop-mensual/xlsx" )
   public  void exportarPlanillaAExcel(HttpServletResponse response) throws IOException{
        servicioPlanillaExcel.exportarPlanillaAExcel(response);
    }

    @GetMapping("/exportar/planillassop-diaria/pdf" )
    public  void exportarPlanillaAPdf(HttpServletResponse response) throws IOException{
        servicioPlanillaPdf.exportarPlanillaAPdf(response);
    }
}
