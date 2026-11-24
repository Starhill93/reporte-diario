#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Consulta Amigable (MEF) -> datos/consulta_amigable.json
=======================================================
Lo corre GitHub Actions (.github/workflows/consulta_amigable.yml) en los servidores de GitHub;
la página del reporte diario lee el .json resultante. Solo usa cifras públicas del MEF.

  ues        MIMP (pliego 039) por unidad ejecutora, año actual
  fen        INABIF (UE 1160) actividad FEN 5006412, año actual
  fen_ues    MIMP: actividad FEN por unidad ejecutora, año actual
  historico  INABIF PIM y devengado desde 2023 (los años cerrados se reutilizan del .json anterior)

Montos: pia, pim, cert, comp (compromiso anual), dev (devengado).

Uso:  python ca/consulta_amigable.py [--salida datos/consulta_amigable.json] [--desde 2023]
Requiere: requests
"""

from __future__ import annotations

import argparse
import json
import re
import sys
import time
from datetime import datetime, timedelta, timezone
from html import unescape
from pathlib import Path

import requests

URL = "https://apps5.mineco.gob.pe/transparencia/Navegador/Navegar_7.aspx?y={anio}&ap=ActProy"
SECTOR, PLIEGO, UE = "39", "039", "1160"
FEN_CATEGORIA, FEN_PRODUCTO, FEN_ACTIVIDAD = "9002", "3999999", "5006412"
LIMA = timezone(timedelta(hours=-5))
RAIZ = Path(__file__).resolve().parent.parent


class Consulta:
    def __init__(self, anio: int):
        self.url = URL.format(anio=anio)
        self.s = requests.Session()
        self.s.headers["User-Agent"] = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                                        "(KHTML, like Gecko) Chrome/124.0 Safari/537.36")

    def _get(self, metodo, **kw) -> str:
        for intento in range(1, 5):
            try:
                r = metodo(self.url, timeout=120, **kw)
                r.raise_for_status()
                return r.text
            except requests.RequestException:
                if intento == 4:
                    raise
                time.sleep(3 * intento)

    @staticmethod
    def filas(t: str) -> list[tuple[str, str, str, dict]]:
        """(valor del radio, código, nombre, montos) de cada fila de la tabla."""
        out = []
        for m in re.finditer(r'<input type="radio" name="grp1" value="([^"]*)".*?<td[^>]*align="left"[^>]*>(.*?)</td>',
                             t, re.S):
            val = unescape(m.group(1))
            p = val.split("/")
            if len(p) != 8:
                continue
            etiqueta = re.sub(r"\s+", " ", unescape(re.sub(r"<[^>]+>", "", m.group(2)))).strip()
            cod, _, nom = etiqueta.partition(":")
            pia, pim, comp, _cm, dev, _gi, cert = (float(x or 0) for x in p[1:])
            out.append((val, cod.strip(), nom.strip(),
                        {"pia": pia, "pim": pim, "cert": cert, "comp": comp, "dev": dev}))
        return out

    def inicio(self) -> str:
        return self._get(self.s.get)

    def elegir(self, t: str, codigo: str, boton: str) -> str:
        val = next((v for v, *_ in self.filas(t) if v.split("/")[0].strip() == codigo), None)
        if val is None:
            raise RuntimeError(f"La Consulta Amigable no muestra el código {codigo}")
        campos = {n: unescape(v) for n, v in
                  re.findall(r'<input type="hidden" name="([^"]+)" id="[^"]*" value="([^"]*)"', t)}
        campos["grp1"] = val
        campos[f"ctl00$CPH1${boton}"] = "x"
        return self._get(self.s.post, data=campos)


def fila(cod: str, nom: str, m: dict) -> dict:
    return {"codigo": cod, "nombre": nom, **m}


def ues_del_pliego(anio: int, con_fen: bool):
    ca = Consulta(anio)
    t = ca.inicio()
    t = ca.elegir(t, "", "BtnTipoGobierno")
    t = ca.elegir(t, "E", "BtnSector")
    t = ca.elegir(t, SECTOR, "BtnPliego")
    t = ca.elegir(t, PLIEGO, "BtnEjecutora")
    ues = [fila(c, n, m) for _, c, n, m in ca.filas(t)]
    if not ues:
        raise RuntimeError(f"Sin UEs del pliego {PLIEGO} en {anio}")
    fen = None
    if con_fen:
        t = ca.elegir(t, UE, "BtnProdProy")
        t = ca.elegir(t, FEN_PRODUCTO, "BtnAAO")
        fen = next((fila(c, n, m) for v, c, n, m in ca.filas(t) if v.split("/")[0] == FEN_ACTIVIDAD), None)
    return ues, fen


def fen_por_ue(anio: int) -> list[dict]:
    ca = Consulta(anio)
    t = ca.inicio()
    t = ca.elegir(t, "", "BtnTipoGobierno")
    t = ca.elegir(t, "E", "BtnSector")
    t = ca.elegir(t, SECTOR, "BtnPliego")
    t = ca.elegir(t, PLIEGO, "BtnProgramaPpto")
    t = ca.elegir(t, FEN_CATEGORIA, "BtnProdProy")
    t = ca.elegir(t, FEN_PRODUCTO, "BtnAAO")
    if not any(v.split("/")[0] == FEN_ACTIVIDAD for v, *_ in ca.filas(t)):
        return []
    t = ca.elegir(t, FEN_ACTIVIDAD, "BtnEjecutora")
    return [fila(c, n, m) for _, c, n, m in ca.filas(t)]


def es_inabif(f: dict) -> bool:
    return f["codigo"].endswith(f"-{UE}") or f["codigo"] == UE


def actualizar(salida: Path, js: Path | None = None, desde: int = 2023, anio: int | None = None,
               siempre: bool = False, log=print) -> dict:
    """Consulta la Consulta Amigable y guarda el .json (y el .js si se pide). Devuelve los datos."""
    anio = anio or datetime.now(LIMA).year
    previo = {}
    if salida.exists():
        try:
            previo = json.loads(salida.read_text(encoding="utf-8"))
        except ValueError:
            pass
    cerrados = {h["anio"]: h for h in previo.get("historico", []) if h["anio"] < anio}

    ues, fen = ues_del_pliego(anio, con_fen=True)
    inabif = next((u for u in ues if es_inabif(u)), None)
    if inabif is None:
        raise RuntimeError(f"No aparece la UE {UE}")
    log(f"{anio}: {len(ues)} UEs · INABIF PIM {inabif['pim']:,.0f} · FEN PIM {(fen or {}).get('pim', 0):,.0f}")
    fen_ues = fen_por_ue(anio)
    log(f"FEN por UE: {len(fen_ues)}")

    historico = []
    for a in range(desde, anio):
        if a not in cerrados:            # los años cerrados no cambian: se consultan una sola vez
            u, _ = ues_del_pliego(a, con_fen=False)
            f = next((x for x in u if es_inabif(x)), None)
            if f is None:
                log(f"{a}: sin INABIF, se omite")
                continue
            cerrados[a] = {"anio": a, "pim": f["pim"], "dev": f["dev"]}
            log(f"{a}: PIM {f['pim']:,.0f} · dev {f['dev']:,.0f}")
        historico.append(cerrados[a])
    historico.append({"anio": anio, "pim": inabif["pim"], "dev": inabif["dev"]})

    datos = {
        "consultado": datetime.now(LIMA).strftime("%Y-%m-%dT%H:%M"),
        "anio": anio,
        "ues": ues,
        "fen": fen,
        "fen_ues": fen_ues,
        "historico": historico,
    }
    # si las cifras no cambiaron, no se reescribe (así GitHub no hace un commit cada hora)
    sin_hora = lambda d: {k: v for k, v in d.items() if k != "consultado"}
    if not siempre and previo and sin_hora(previo) == sin_hora(datos):
        log("Sin cambios")
        return datos
    texto = json.dumps(datos, ensure_ascii=False, indent=1)
    salida.parent.mkdir(parents=True, exist_ok=True)
    salida.write_text(texto, encoding="utf-8")
    log(f"Guardado {salida}")
    if js:
        js.parent.mkdir(parents=True, exist_ok=True)
        js.write_text(f"window.CA_LOCAL = {texto};\n", encoding="utf-8")
        log(f"Guardado {js}")
    return datos


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--salida", type=Path, default=RAIZ / "datos" / "consulta_amigable.json")
    ap.add_argument("--desde", type=int, default=2023)
    ap.add_argument("--anio", type=int, default=datetime.now(LIMA).year)
    ap.add_argument("--js", type=Path, default=None,
                    help="Además guarda una copia .js (window.CA_LOCAL) para abrir la página con doble clic")
    ap.add_argument("--siempre", action="store_true", help="Guardar aunque las cifras no hayan cambiado")
    args = ap.parse_args()
    actualizar(args.salida, args.js, args.desde, args.anio, args.siempre)


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"ERROR: {e}", file=sys.stderr)
        sys.exit(1)