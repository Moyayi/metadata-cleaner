#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Genera el icono icon.ico para MetaCleaner.
Ejecutar una sola vez antes de compilar:
    python assets/create_icon.py
"""

from PIL import Image, ImageDraw, ImageFont
import os

SIZES = [256, 128, 64, 48, 32, 16]

BG      = (13,  13,  13)       # #0D0D0D
ACCENT  = (124, 106, 247)      # #7C6AF7
WHITE   = (239, 239, 239)      # #EFEFEF


def dibujar_icono(size: int) -> Image.Image:
    img  = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)

    # Fondo redondeado
    radio = size // 6
    draw.rounded_rectangle([0, 0, size - 1, size - 1], radius=radio, fill=BG)

    # Escudo simplificado
    pad  = size // 8
    cx   = size // 2
    top  = pad
    bot  = size - pad
    ancho = int(size * 0.55)
    izq  = cx - ancho // 2
    der  = cx + ancho // 2

    # Polígono del escudo (5 puntos)
    mid_y = top + (bot - top) * 55 // 100
    escudo = [
        (izq, top),
        (der, top),
        (der, mid_y),
        (cx,  bot),
        (izq, mid_y),
    ]
    draw.polygon(escudo, fill=ACCENT)

    # Cerradura (candado)
    cy    = top + (bot - top) * 42 // 100
    r_arco = max(2, size // 10)
    grosor = max(1, size // 24)
    bbox_arco = [cx - r_arco, cy - r_arco, cx + r_arco, cy]
    draw.arc(bbox_arco, start=0, end=180, fill=WHITE, width=grosor)

    # Cuerpo del candado
    cw = max(2, r_arco * 2)
    ch = max(2, int(r_arco * 1.4))
    draw.rounded_rectangle(
        [cx - cw // 2, cy, cx + cw // 2, cy + ch],
        radius=max(1, cw // 6),
        fill=WHITE,
    )

    return img


def main():
    iconos = [dibujar_icono(s) for s in SIZES]
    salida = os.path.join(os.path.dirname(__file__), "icon.ico")
    iconos[0].save(
        salida,
        format="ICO",
        sizes=[(s, s) for s in SIZES],
        append_images=iconos[1:],
    )
    print(f"Icono generado: {salida}")


if __name__ == "__main__":
    main()
