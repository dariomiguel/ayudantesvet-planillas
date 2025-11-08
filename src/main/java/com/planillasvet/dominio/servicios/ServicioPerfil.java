package com.planillasvet.dominio.servicios;

import com.planillasvet.dominio.entidades.Usuario;

public interface ServicioPerfil {
    Usuario obtenerPerfil(String email);
    void actualizarPerfil(Usuario usuario);
}