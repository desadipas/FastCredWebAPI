<Serializable()> Public Class ClsPersona

    Public Data As ComTipos.clsDataset
    Public oDirec As ComTipos.ClsPersDireccion
    Public Documentos As ComTipos.clsDataset
    Public Parentesco As ComTipos.clsDataset
    Public Dependientes As ComTipos.clsDataset
    Public Vivienda As ComTipos.clsDataset
    Public EstadosCivil As ComTipos.clsDataset
    Public ActividadEco As ComTipos.clsDataset
    Public RelaInst As ComTipos.clsDataset
    Public Nacionalidad As ComTipos.clsDataset
    Public TiposDoc As ComTipos.clsDataset
    Public TiposParent As ComTipos.clsDataset
    Public Estados As ComTipos.clsDataset
    Public GradosInt As ComTipos.clsDataset
    Public Profesiones As ComTipos.clsDataset
    Public oPropiedad As ComTipos.clsDataset
    Public oPropietario As ComTipos.clsDataset
    Public oZona As ComTipos.clsDataset
    Public oBarrio As ComTipos.clsDataset
    Public oTipoViv As ComTipos.clsDataset
    Public oMaterial As ComTipos.clsDataset
    Public oEstadoViv As ComTipos.clsDataset
    Public oDirecTrab As ComTipos.ClsPersDireccion
    Public oNegocio As ComTipos.ClsPersDireccion
    Public oSucursales As ComTipos.clsDataset
    Public oSiglas As ComTipos.clsDataset



    'PARA PERSONA JURIDICA
    Public TiposMagnEmp As ComTipos.clsDataset
    Public TiposPersonaJur As ComTipos.clsDataset
    Public EstadosPerJur As ComTipos.clsDataset
    Public oRepresentantes As ComTipos.clsDataset
    Public oAccionistas As ComTipos.clsDataset

    Public Sub New()

        Data = New ComTipos.clsDataset
        Documentos = New ComTipos.clsDataset
        Parentesco = New ComTipos.clsDataset
        EstadosCivil = New ComTipos.clsDataset
        ActividadEco = New ComTipos.clsDataset
        RelaInst = New ComTipos.clsDataset
        Nacionalidad = New ComTipos.clsDataset
        TiposDoc = New ComTipos.clsDataset
        TiposParent = New ComTipos.clsDataset
        Dependientes = New ComTipos.clsDataset
        Vivienda = New ComTipos.clsDataset
        Estados = New ComTipos.clsDataset
        TiposMagnEmp = New ComTipos.clsDataset
        TiposPersonaJur = New ComTipos.clsDataset
        oDirec = New ComTipos.ClsPersDireccion
        EstadosPerJur = New ComTipos.clsDataset
        GradosInt = New ComTipos.clsDataset
        Profesiones = New ComTipos.clsDataset
        oDirecTrab = New ComTipos.ClsPersDireccion
        oNegocio = New ComTipos.ClsPersDireccion
        oSucursales = New ComTipos.clsDataset
        oSiglas = New ComTipos.clsDataset

        oPropiedad = New ComTipos.clsDataset
        oPropietario = New ComTipos.clsDataset
        oZona = New ComTipos.clsDataset
        oBarrio = New ComTipos.clsDataset
        oTipoViv = New ComTipos.clsDataset
        oMaterial = New ComTipos.clsDataset
        oEstadoViv = New ComTipos.clsDataset
        oRepresentantes = New ComTipos.clsDataset
        oAccionistas = New ComTipos.clsDataset

    End Sub
End Class
