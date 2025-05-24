
CREATE TABLE RespuestasIA (
    Id INT PRIMARY KEY IDENTITY,
    Pregunta NVARCHAR(500) NOT NULL,
    Respuesta NVARCHAR(MAX) NOT NULL,
    Tokens INT NOT NULL,
    TiempoRespuesta FLOAT NOT NULL,
    Fecha DATETIME NOT NULL DEFAULT GETDATE()
);

use dbInvestigacion;


select * from RespuestasIA;

