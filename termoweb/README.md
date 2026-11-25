# Calefacción TermoWeb

Programa los radiadores de la cuenta TermoWeb con un horario semanal hora a hora y ajusta cada día las temperaturas según la previsión de Zarautz (Open-Meteo).

## Horario

| Días | Confort | Eco |
|---|---|---|
| Lunes a viernes | 7:00–9:00 y 12:00–21:00 | el resto (noche y de 9:00 a 12:00, cuando no hay nadie) |
| Sábado y domingo | 7:00–21:00 | noche |

El confort empieza una hora antes de levantaros (7:00) y de volver (12:00) para que la casa ya esté caliente. Pasa a eco a las 21:00: el radiador guarda calor y la casa se enfría despacio hasta las 21:30. Se cambia en `config.json`, en `horario` (horas enteras, `"7-9"` = de 7:00 a 9:00).

## Temperaturas

Base: **confort 20 °C, eco 17 °C, antihielo 7 °C**. Cada mañana se ajustan con la previsión:

- **Confort**, según la sensación térmica media en las horas de confort: +1 °C si es menos de 5 °C, +0,5 por debajo de 10, sin cambio hasta 15, −1 hasta 18 y −2 por encima. Siempre queda entre 18 y 22 °C.
- **Eco**, según la mínima de la noche: +1 °C si baja de 2 °C y +0,5 si baja de 6. Queda entre 15 y 18,5 °C.

Si pasáis frío o calor, cambia `temperaturas` en `config.json`.

## Puesta en marcha (Windows)

1. Instala Python desde python.org (marca «Add python.exe to PATH»). No hace falta nada más.
2. Copia `credenciales.ejemplo.json` como `credenciales.json` y pon tu correo y tu contraseña de TermoWeb. Ese archivo no se sube nunca (`.gitignore`).
3. `termoweb.cmd ver`: muestra los radiadores y el horario que tienen ahora. **Compáralo con la app** para comprobar que los días y las horas se leen igual.
4. `termoweb.cmd horario`: envía el horario y pone los radiadores en automático. Solo hace falta una vez, o cuando cambies el horario.
5. `instalar-tarea.cmd`: crea una tarea de Windows que ejecuta `ajustar` cada día a las 6:00. El PC tiene que estar encendido a esa hora. Si no lo está, el ajuste se hace la siguiente vez que se ejecute.

Añade `--prueba` a cualquier orden para ver lo que haría sin cambiar nada. `termoweb.cmd tiempo` muestra la previsión y las temperaturas de hoy. Lo que se cambia queda anotado en `registro.txt`.

TermoWeb no tiene API oficial: se usa la misma que la app, documentada por la comunidad. Si un día cambian algo y deja de funcionar, la app sigue funcionando como siempre.
