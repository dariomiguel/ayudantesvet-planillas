package com.ayudantesvet.planillas.domain.services;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;

public interface ServicioPlanillaPdf {
    void exportarPlanillaAPdf (HttpServletResponse response) throws IOException;
}
