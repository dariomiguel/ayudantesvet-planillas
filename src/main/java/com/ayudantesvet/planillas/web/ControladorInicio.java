package com.ayudantesvet.planillas.web;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.servlet.ModelAndView;

@Controller
public class ControladorInicio {

    @Autowired
    public ControladorInicio(){
    }

    @RequestMapping(path = "/", method = RequestMethod.GET)
    public ModelAndView inicio() {
        return new ModelAndView("redirect:/inicio");
    }

    @RequestMapping(path = "/inicio", method = RequestMethod.GET)
    public ModelAndView irAInicio() {
        return new ModelAndView("inicio");
    }
}

