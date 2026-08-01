# -*- coding: utf-8 -*-
"""Phase 5: revisao_sugerida vocabulary is single-source and used by emitters."""

from __future__ import annotations

import re
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]


class TestRevisaoVocabulario(unittest.TestCase):
    def test_helpers_roundtrip(self):
        from textura.revisao import (
            etiqueta_exacta,
            etiqueta_prefixada,
            juntar_etiquetas,
            validar_revisao_sugerida,
        )

        a = etiqueta_exacta("atributiva_coordenada")
        b = etiqueta_prefixada("dominio_janela", "geologia")
        blob = juntar_etiquetas(a, b)
        self.assertEqual(validar_revisao_sugerida(blob), [])
        self.assertEqual(
            validar_revisao_sugerida("etiqueta_inventada"),
            ["etiqueta_inventada"],
        )

    def test_emitters_draw_from_helpers(self):
        """Ban ad-hoc f-string / concat construction outside the helpers."""
        ban_prefix = (
            'f"coordenacao_heterogenea:',
            "f'coordenacao_heterogenea:",
            'f"associativa_com_nao_textural:',
            "f'associativa_com_nao_textural:",
            'f"dominio_janela:',
            "f'dominio_janela:",
            '"dominio_janela:" +',
            "'dominio_janela:' +",
            '+ "dominio_janela:"',
            "+ 'dominio_janela:'",
        )
        allowed = {ROOT / "textura" / "revisao.py"}
        offenders = []
        for path in (ROOT / "textura").rglob("*.py"):
            if path.resolve() in {p.resolve() for p in allowed}:
                continue
            text = path.read_text(encoding="utf-8")
            for needle in ban_prefix:
                if needle in text:
                    offenders.append(f"{path.relative_to(ROOT)}: {needle}")
            # Exact tags must go through etiqueta_exacta(...)
            for tag in (
                "genitiva_por_complemento",
                "atributiva_via_conj",
                "atributiva_coordenada",
            ):
                for m in re.finditer(re.escape(f'"{tag}"'), text):
                    start = max(0, m.start() - 40)
                    ctx = text[start:m.end() + 5]
                    if "etiqueta_exacta" not in ctx:
                        offenders.append(
                            f"{path.relative_to(ROOT)}: bare {tag!r}"
                        )
        self.assertEqual(
            offenders, [], "raw revisao construction:\n" + "\n".join(offenders)
        )

    def test_no_unknown_emission_sites(self):
        """Every etiqueta_* call site is in known modules."""
        rx = re.compile(r"etiqueta_(?:exacta|prefixada)\(")
        sites = []
        for path in (ROOT / "textura").rglob("*.py"):
            if path.name == "revisao.py":
                continue
            for i, line in enumerate(path.read_text(encoding="utf-8").splitlines(), 1):
                if rx.search(line):
                    sites.append(f"{path.relative_to(ROOT)}:{i}")
        self.assertTrue(sites, "expected at least one emitter")
        for s in sites:
            self.assertTrue(
                s.startswith("textura/relacoes.py")
                or s.startswith("textura\\relacoes.py")
                or s.startswith("textura/pipeline.py")
                or s.startswith("textura\\pipeline.py"),
                f"unexpected emitter site: {s}",
            )


if __name__ == "__main__":
    unittest.main()
