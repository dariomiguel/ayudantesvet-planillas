package com.ayudantesvet.planillas.domain.services;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;

public interface ServicioPlanillaExcel {
    void exportarPlanillaAExcel(HttpServletResponse response) throws IOException;
}
