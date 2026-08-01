#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Léxicos da extracção (Phase 1 home; Phase 2 will load from dados/)."""

from __future__ import annotations

from textura.config import NOS  # re-export convenience
from textura.tokenizacao import normaliza




# Léxico mínimo, por janela, para usos extramusicais de «textura» que a
# triagem documental (por ficheiro) não consegue ver. Aditivo e conservador:
# apenas sinaliza para revisão; nunca altera ``nuclear``.
DOMINIO_JANELA_LEXICO: dict[str, tuple[str, ...]] = {
    "geologia": ("dolomite", "breccia", "facies", "bioclast", "waulsortian",
                 "sedimentar", "mineralog"),
    "artes_visuais": ("painting", "paintings", "mural", "canvas",
                      "flat plane"),
    "haptica_materiais": ("haptic", "tactile", "knitting needles",
                          "materials library"),
    "ecolocalizacao": ("echolocation", "bat-inspired", "bat inspired"),
    "fala": ("speech recognition", "speech waveform", "voice conversion"),
    "texto_social": ("social texture", "reading and writing"),
}


def dominio_janela(contexto: str) -> str:
    """Domínio extramusical sugerido pela própria janela ('' se nenhum)."""
    c = normaliza(str(contexto or "")).lower()
    for dom, pistas in DOMINIO_JANELA_LEXICO.items():
        for p in pistas:
            if p in c:
                return dom
    return ""
