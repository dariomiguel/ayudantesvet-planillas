package com.ayudantesvet.planillas.domain.services;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import javax.transaction.Transactional;

@Service("servicioInicio")
@Transactional
public class ServicioInicioImpl implements ServicioInicio {

    @Autowired
    public ServicioInicioImpl(){

    }
}

