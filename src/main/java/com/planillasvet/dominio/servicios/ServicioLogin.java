package com.planillasvet.dominio.servicios;

import com.planillasvet.dominio.entidades.Usuario;
import com.planillasvet.dominio.UsuarioExistente;

public interface ServicioLogin {

    Usuario consultarUsuario(String email, String password);
    void registrar(Usuario usuario) throws UsuarioExistente;

}
