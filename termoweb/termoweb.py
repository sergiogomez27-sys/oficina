"""Programa los radiadores TermoWeb: horario semanal hora a hora y temperaturas segun el tiempo.

Uso:
    python termoweb.py ver                 muestra los radiadores, sus temperaturas y su horario actual
    python termoweb.py horario [--prueba]  envia el horario semanal de config.json y pone los radiadores en automatico
    python termoweb.py ajustar [--prueba]  calcula las temperaturas de hoy con la prevision y las envia
    python termoweb.py tiempo              solo muestra la prevision y las temperaturas que pondria

Con --prueba no se cambia nada: solo se muestra lo que se enviaria.
Solo usa la biblioteca estandar de Python (no hace falta instalar nada).
"""

import base64
import datetime as dt
import json
import sys
import urllib.error
import urllib.parse
import urllib.request
from pathlib import Path

CARPETA = Path(__file__).resolve().parent
SERVIDOR = "https://control.termoweb.net"
# Identificador de la aplicacion oficial (el mismo para todos los usuarios, no es un secreto tuyo).
CLIENTE = base64.b64encode(b"52172dc84f63d6c759000005:bxv4Z3xUSe").decode()

DIAS = ["lunes", "martes", "miercoles", "jueves", "viernes", "sabado", "domingo"]
# Valores de cada hora en el programa de TermoWeb
ANTIHIELO, ECO, CONFORT = 0, 1, 2
LETRA = {ANTIHIELO: "a", ECO: "·", CONFORT: "C"}


def registrar(texto):
    print(texto)
    try:
        with open(CARPETA / "registro.txt", "a", encoding="utf-8") as f:
            f.write(f"{dt.datetime.now():%Y-%m-%d %H:%M}  {texto}\n")
    except OSError:
        pass


def leer_json(nombre):
    with open(CARPETA / nombre, encoding="utf-8") as f:
        return json.load(f)


# ---------------------------------------------------------------- TermoWeb

class TermoWeb:
    def __init__(self, usuario, contrasena):
        datos = urllib.parse.urlencode(
            {"grant_type": "password", "username": usuario, "password": contrasena}
        ).encode()
        peticion = urllib.request.Request(
            SERVIDOR + "/client/token",
            data=datos,
            headers={
                "Authorization": "Basic " + CLIENTE,
                "Content-Type": "application/x-www-form-urlencoded",
            },
        )
        try:
            with urllib.request.urlopen(peticion, timeout=20) as r:
                self.token = json.load(r)["access_token"]
        except urllib.error.HTTPError as e:
            if e.code in (400, 401):
                sys.exit("TermoWeb no acepta el usuario o la contrasena de credenciales.json.")
            raise

    def _pedir(self, ruta, cuerpo=None):
        peticion = urllib.request.Request(
            SERVIDOR + ruta,
            data=None if cuerpo is None else json.dumps(cuerpo).encode(),
            headers={
                "Authorization": "Bearer " + self.token,
                "Content-Type": "application/json",
                "Accept": "application/json",
            },
            method="GET" if cuerpo is None else "POST",
        )
        with urllib.request.urlopen(peticion, timeout=20) as r:
            texto = r.read().decode() or "{}"
        return json.loads(texto)

    def radiadores(self):
        """Devuelve [(dev_id, tipo, addr, nombre)] de todas las centrales de la cuenta."""
        lista = []
        for dev in self._pedir("/api/v2/connected_devices").get("devs", []):
            dev_id = dev["dev_id"]
            for nodo in self._pedir(f"/api/v2/devs/{dev_id}/mgr/nodes").get("nodes", []):
                if nodo.get("type") in ("htr", "acm"):
                    nombre = nodo.get("name") or f"{nodo['type']} {nodo['addr']}"
                    lista.append((dev_id, nodo["type"], nodo["addr"], nombre.strip()))
        return lista

    def ajustes(self, dev_id, tipo, addr):
        return self._pedir(f"/api/v2/devs/{dev_id}/{tipo}/{addr}/settings")

    def cambiar(self, dev_id, tipo, addr, cambios):
        return self._pedir(f"/api/v2/devs/{dev_id}/{tipo}/{addr}/settings", cambios)


def conectar():
    try:
        cred = leer_json("credenciales.json")
    except FileNotFoundError:
        sys.exit(
            "Falta credenciales.json. Copia credenciales.ejemplo.json como credenciales.json "
            "y pon tu correo y contrasena de TermoWeb."
        )
    return TermoWeb(cred["usuario"], cred["contrasena"])


def elegidos(tw, config):
    todos = tw.radiadores()
    filtro = config.get("radiadores", "todos")
    if filtro == "todos":
        return todos
    nombres = {n.lower() for n in filtro}
    return [r for r in todos if r[3].lower() in nombres]


# ---------------------------------------------------------------- Horario

def programa_semanal(horario):
    """Convierte {"lunes": ["7-9", "12-21"], ...} en las 168 horas (lunes 0:00 primero)."""
    prog = []
    for dia in DIAS:
        horas = [ECO] * 24
        for tramo in horario.get(dia, []):
            desde, hasta = (int(x) for x in tramo.split("-"))
            for h in range(desde, hasta):
                horas[h] = CONFORT
        prog.extend(horas)
    return prog


def dibujar(prog):
    print("            " + "".join(f"{h:<3}" for h in range(0, 24, 3)).rstrip())
    for i, dia in enumerate(DIAS):
        fila = "".join(LETRA.get(v, "?") for v in prog[i * 24:(i + 1) * 24])
        print(f"  {dia:<10}{fila}")
    print("  C = confort   · = eco   a = antihielo")


# ---------------------------------------------------------------- Tiempo

def prevision(ubicacion):
    q = urllib.parse.urlencode({
        "latitude": ubicacion["lat"],
        "longitude": ubicacion["lon"],
        "hourly": "apparent_temperature,temperature_2m",
        "timezone": ubicacion.get("zona_horaria", "Europe/Madrid"),
        "forecast_days": 2,
    })
    with urllib.request.urlopen("https://api.open-meteo.com/v1/forecast?" + q, timeout=20) as r:
        datos = json.load(r)["hourly"]
    return [
        (dt.datetime.fromisoformat(t), sens, temp)
        for t, sens, temp in zip(datos["time"], datos["apparent_temperature"], datos["temperature_2m"])
    ]


def sumar_segun(tabla, clave, valor):
    for fila in tabla:
        if valor < fila[clave]:
            return fila["sumar"]
    return 0


def temperaturas_de_hoy(config, hoy=None):
    """Devuelve (confort, eco, explicacion) para hoy segun la prevision."""
    base = config["temperaturas"]
    ajuste = config.get("ajuste_clima", {})
    if not ajuste.get("activo"):
        return base["confort"], base["eco"], "ajuste por clima desactivado"

    hoy = hoy or dt.date.today()
    prog = programa_semanal(config["horario"])
    dia = hoy.weekday()
    horas_confort = {h for h in range(24) if prog[dia * 24 + h] == CONFORT}

    datos = prevision(config["ubicacion"])
    manana = hoy + dt.timedelta(days=1)
    despierto = [s for t, s, _ in datos if t.date() == hoy and t.hour in horas_confort]
    # La noche que viene: desde las 21 de hoy hasta las 9 de manana
    noche = [temp for t, _, temp in datos
             if (t.date() == hoy and t.hour >= 21) or (t.date() == manana and t.hour < 9)]
    if not despierto or not noche:
        return base["confort"], base["eco"], "sin prevision para hoy, uso las temperaturas base"

    sensacion = sum(despierto) / len(despierto)
    minima = min(noche)

    lim = ajuste.get("limites", {})
    confort = base["confort"] + sumar_segun(ajuste["confort"], "si_sensacion_menor_de", sensacion)
    eco = base["eco"] + sumar_segun(ajuste["eco"], "si_minima_menor_de", minima)
    confort = min(max(confort, lim.get("confort", [0, 99])[0]), lim.get("confort", [0, 99])[1])
    eco = min(max(eco, lim.get("eco", [0, 99])[0]), lim.get("eco", [0, 99])[1])
    eco = min(eco, confort - 1)

    explicacion = (f"{config['ubicacion']['nombre']}: sensacion media en horas de confort "
                   f"{sensacion:.1f} °C, minima esta noche {minima:.1f} °C")
    return round(confort * 2) / 2, round(eco * 2) / 2, explicacion


# ---------------------------------------------------------------- Ordenes

def orden_ver(config):
    tw = conectar()
    radiadores = tw.radiadores()
    if not radiadores:
        print("La cuenta no tiene radiadores.")
    for dev_id, tipo, addr, nombre in radiadores:
        a = tw.ajustes(dev_id, tipo, addr)
        print(f"\n{nombre}  (central {dev_id}, {tipo} {addr})")
        print(f"  modo: {a.get('mode')}   ahora: {a.get('mtemp')} °C   objetivo: {a.get('stemp')} °C")
        pt = a.get("ptemp")
        if pt and len(pt) == 3:
            print(f"  temperaturas: antihielo {pt[0]}  eco {pt[1]}  confort {pt[2]}  ({a.get('units', 'C')})")
        if a.get("prog") and len(a["prog"]) == 168:
            dibujar(a["prog"])
    print("\nCompara este horario con el que ves en la app. Si los dias o las horas no "
          "coinciden, avisa antes de usar 'horario'.")


def orden_horario(config, prueba):
    prog = programa_semanal(config["horario"])
    print("Horario que se envia:")
    dibujar(prog)
    if prueba:
        print("\n--prueba: no se ha cambiado nada.")
        return
    tw = conectar()
    for dev_id, tipo, addr, nombre in elegidos(tw, config):
        tw.cambiar(dev_id, tipo, addr, {"prog": prog})
        tw.cambiar(dev_id, tipo, addr, {"mode": "auto"})
        registrar(f"horario enviado a {nombre} y puesto en automatico")


def orden_ajustar(config, prueba):
    confort, eco, explicacion = temperaturas_de_hoy(config)
    antihielo = config["temperaturas"]["antihielo"]
    print(explicacion)
    print(f"Temperaturas de hoy: confort {confort} °C, eco {eco} °C, antihielo {antihielo} °C")
    if prueba:
        print("--prueba: no se ha cambiado nada.")
        return
    tw = conectar()
    ptemp = [f"{antihielo:.1f}", f"{eco:.1f}", f"{confort:.1f}"]
    for dev_id, tipo, addr, nombre in elegidos(tw, config):
        tw.cambiar(dev_id, tipo, addr, {"ptemp": ptemp, "units": "C"})
        registrar(f"{nombre}: confort {confort}, eco {eco}  ({explicacion})")


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    prueba = "--prueba" in sys.argv
    orden = args[0] if args else "ver"
    config = leer_json("config.json")
    try:
        if orden == "ver":
            orden_ver(config)
        elif orden == "horario":
            orden_horario(config, prueba)
        elif orden == "ajustar":
            orden_ajustar(config, prueba)
        elif orden == "tiempo":
            orden_ajustar(config, prueba=True)
        else:
            print(__doc__)
    except urllib.error.HTTPError as e:
        registrar(f"error del servidor ({e.code}) en {e.url}: {e.read().decode(errors='replace')[:300]}")
        sys.exit(1)
    except urllib.error.URLError as e:
        registrar(f"sin conexion: {e.reason}")
        sys.exit(1)


if __name__ == "__main__":
    main()
