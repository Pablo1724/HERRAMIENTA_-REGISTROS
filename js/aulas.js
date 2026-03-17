const configuracionAulas = {
    modalidades: {
        SEMANAL: {
            dias: "Lunes a Viernes",
            turnos: [
                { id: "S_MAÑANA", nombre: "Mañana", horario: "07:15 - 10:15" },
                { id: "S_MEDIODIA", nombre: "Mediodía", horario: "10:30 - 13:00" },
                { id: "S_TARDE", nombre: "Tarde", horario: "15:00 - 18:00" },
                { id: "S_NOCHE", nombre: "Noche", horario: "19:00 - 22:00" }
            ]
        },
        SEMIPRESENCIAL: {
            turnos: [
                { id: "SP_MAÑANA", nombre: "Mañana (Sábados)", horario: "08:15 - 12:15" },
                { id: "SP_TARDE", nombre: "Tarde (Sábados)", horario: "15:00 - 19:00" },
                { id: "SP_NOCHE", nombre: "Noche (Lunes)", horario: "19:00 - 22:15" }
            ]
        }
    },
    listadoAulas: [
        { id: "A101", nombre: "Aula 101", piso: "PB", capacidad: 50 },
        { id: "A102", nombre: "Aula 102", piso: "PB", capacidad: 45 },
        { id: "L001", nombre: "Laboratorio I", piso: "P1", capacidad: 30 },
        // ... agregar más aulas según la sede
    ]
};