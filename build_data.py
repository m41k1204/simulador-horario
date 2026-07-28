#!/usr/bin/env python3
"""Convierte el Excel de Consulta de Horario de UTEC en courses.json.

Cruza la oferta del semestre con carga_habil.tsv (los cursos que al alumno
todavia le falta llevar) y se queda solo con la interseccion: lo que esta
ofertado Y puede matricular. Cada curso queda marcado como Obligatorio o
Electivo, con su ciclo y creditos.

Uso:  python build_data.py [ruta_al_excel]
"""
import json
import re
import sys
from collections import OrderedDict
from pathlib import Path

import openpyxl

# prefijo de codigo -> departamento (solo para mostrar)
DEPTOS = {
    "CS": "Ciencia de la Computación", "DS": "Ciencia de Datos",
    "IS": "Sistemas de Información", "CY": "Ciberseguridad",
    "CC": "Ciencias", "HH": "Humanidades", "PI": "Proyectos Interdisciplinarios",
    "AD": "Administración", "GI": "Gestión e Innovación", "BA": "Business Analytics",
    "AM": "Ing. Ambiental", "BI": "Bioingeniería", "CI": "Ing. Civil",
    "EL": "Electrónica", "EN": "Ing. de la Energía", "IN": "Ing. Industrial",
    "IQ": "Ing. Química", "ME": "Ing. Mecánica", "MT": "Mecatrónica",
    "GE": "Global Education", "CL": "Curso Libre",
}

# afinidad con la carrera: define el orden y el filtro "de mi facultad"
FACULTAD = {"CS", "DS", "IS", "CY", "CC"}

DIAS = {"Lun.": 1, "Mar.": 2, "Mie.": 3, "Jue.": 4, "Vie.": 5, "Sab.": 6, "Dom.": 7}

HORARIO_RE = re.compile(r"^(\w+\.)\s*(\d{2}):(\d{2})\s*-\s*(\d{2}):(\d{2})$")

FILA_DATOS = 9  # las 8 primeras filas son cabecera del reporte


def parse_horario(texto):
    m = HORARIO_RE.match(str(texto).strip())
    if not m:
        raise ValueError("horario no reconocido: %r" % texto)
    dia, h1, m1, h2, m2 = m.groups()
    return {
        "dia": DIAS[dia],
        "inicio": int(h1) * 60 + int(m1),
        "fin": int(h2) * 60 + int(m2),
    }


def leer_filas(path):
    ws = openpyxl.load_workbook(path, data_only=True)["Hoja 1"]
    for fila in ws.iter_rows(min_row=FILA_DATOS, values_only=True):
        if fila[0]:
            yield fila


def leer_carga_habil(path):
    """codigo -> {tipo, ciclo, creditos} de los cursos que aun puede llevar."""
    habil = {}
    for linea in path.read_text(encoding="utf-8").splitlines():
        if not linea.strip() or linea.startswith("#"):
            continue
        cod, tipo, ciclo, creditos = [c.strip() for c in linea.split("\t")]
        habil[cod] = {"tipo": tipo, "ciclo": ciclo, "creditos": int(creditos)}
    return habil


def main():
    xlsx = Path(sys.argv[1]) if len(sys.argv) > 1 else next(
        Path(__file__).resolve().parent.parent.glob("Consulta_Horario-*.xlsx")
    )
    habil = leer_carga_habil(Path(__file__).resolve().parent / "carga_habil.tsv")
    cursos = OrderedDict()
    descartados = set()

    for (codigo, nombre, seccion, grupo, modalidad, horario, frecuencia,
         ubicacion, vacantes, matriculados, docente, correo) in leer_filas(xlsx):
        if codigo not in habil:
            descartados.add(codigo)
            continue
        prefijo = re.match(r"^[A-Z]+", codigo).group(0)

        curso = cursos.setdefault(codigo, {
            "codigo": codigo,
            "nombre": nombre.strip(),
            "prefijo": prefijo,
            "departamento": DEPTOS.get(prefijo, prefijo),
            "facultad": prefijo in FACULTAD,
            **habil[codigo],
            "secciones": OrderedDict(),
        })
        sec = curso["secciones"].setdefault(str(seccion), {
            "seccion": str(seccion),
            "vacantes": vacantes,
            "matriculados": matriculados,
            "docentes": [],
            "fijas": [],       # teoria: se llevan si o si con la seccion
            "labs": OrderedDict(),  # grupos de laboratorio: se elige uno
        })

        docente = (docente or "").strip()
        if docente and docente not in sec["docentes"]:
            sec["docentes"].append(docente)

        sesion = parse_horario(horario)
        sesion.update({
            "grupo": grupo,
            "tipo": "lab" if str(grupo).upper().startswith("LABORATORIO") else "teoria",
            "modalidad": modalidad,
            "frecuencia": frecuencia,
            "ubicacion": ubicacion,
            "docente": docente,
            # el correo del docente viene en el Excel pero no se exporta:
            # el repo es publico y no aporta nada para armar el horario
        })

        if sesion["tipo"] == "lab":
            sec["labs"].setdefault(grupo, []).append(sesion)
        else:
            sec["fijas"].append(sesion)

    salida = []
    for curso in cursos.values():
        curso["secciones"] = [
            {**sec, "labs": [{"grupo": g, "sesiones": s} for g, s in sec["labs"].items()]}
            for sec in curso["secciones"].values()
        ]
        salida.append(curso)
    # primero los obligatorios (por ciclo), luego electivos de la facultad, luego el resto
    salida.sort(key=lambda c: (
        c["tipo"] != "Obligatorio",
        int(c["ciclo"]) if c["ciclo"].isdigit() else 99,
        not c["facultad"],
        c["codigo"],
    ))

    datos = json.dumps({"fuente": xlsx.name, "cursos": salida}, ensure_ascii=False)
    carpeta = Path(__file__).resolve().parent
    destino = carpeta / "courses.json"
    destino.write_text(datos, encoding="utf-8")
    (carpeta / "data.js").write_text("window.DATA = %s;\n" % datos, encoding="utf-8")

    # horario.html: todo en un solo archivo (sin data.js) para abrirlo con doble clic
    plantilla = (carpeta / "index.html").read_text(encoding="utf-8")
    marca = '<script src="data.js"></script>'
    if marca not in plantilla:
        raise SystemExit("index.html ya no incluye %s" % marca)
    (carpeta / "horario.html").write_text(
        plantilla.replace(marca, "<script>window.DATA = %s;</script>"
                          % datos.replace("</", "<\\/")), encoding="utf-8")

    secciones = sum(len(c["secciones"]) for c in salida)
    obl = [c["codigo"] for c in salida if c["tipo"] == "Obligatorio"]
    print("%s: %d cursos (%d obligatorios), %d secciones"
          % (destino.name, len(salida), len(obl), secciones))
    print("obligatorios ofertados:", " ".join(obl))
    faltan = sorted(c for c in habil if c not in cursos)
    print("de tu carga hábil, sin oferta este semestre: %d" % len(faltan))
    print("cursos ofertados que ya no puedes llevar: %d" % len(descartados))


if __name__ == "__main__":
    main()
