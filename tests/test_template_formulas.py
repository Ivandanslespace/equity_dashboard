"""Tests des recherches par en-tête dans le modèle Excel."""

import posixpath
import re
import zipfile

from lxml import etree

from ooxml_writer import NS_MAIN, NS_REL

from test_ooxml_writer import TEMPLATE


NS_PKG_REL = (
    "http://schemas.openxmlformats.org/package/2006/relationships"
)


def _parties_feuilles(archive):
    """Retourne les chemins OOXML indexés par nom de feuille."""
    classeur = etree.fromstring(archive.read("xl/workbook.xml"))
    relations = etree.fromstring(
        archive.read("xl/_rels/workbook.xml.rels")
    )
    cibles = {
        item.get("Id"): item.get("Target")
        for item in relations
    }
    feuilles = classeur.find(f"{{{NS_MAIN}}}sheets")
    return {
        feuille.get("name"): posixpath.normpath(
            posixpath.join(
                "xl",
                cibles[feuille.get(f"{{{NS_REL}}}id")],
            )
        )
        for feuille in feuilles
    }


def _formule(archive, partie, adresse):
    """Lit la formule OOXML d'une cellule précise."""
    root = etree.fromstring(archive.read(partie))
    cellule = root.find(
        f".//{{{NS_MAIN}}}c[@r='{adresse}']"
    )
    return cellule.find(f"{{{NS_MAIN}}}f").text


def test_formules_data_recherchent_les_colonnes_par_entete():
    """Les feuilles clés ne doivent plus dépendre des lettres de métriques."""
    with zipfile.ZipFile(TEMPLATE) as archive:
        parties = _parties_feuilles(archive)
        formules = {
            "Analyse": _formule(
                archive, parties["Analyse"], "I2"
            ),
            "Screening": _formule(
                archive, parties["Screening"], "E2"
            ),
            "Proposition": _formule(
                archive, parties["Proposition"], "C6"
            ),
            "Optim": _formule(
                archive, parties["Optim"], "AF2"
            ),
            "Explicabilite": _formule(
                archive, parties["Explicabililité"], "A3"
            ),
        }

    assert all("MATCH(" in formule for formule in formules.values())
    assert 'MATCH("Weight PTF"' in formules["Analyse"]
    assert 'MATCH("Name"' in formules["Proposition"]
    assert 'MATCH("ISIN"' in formules["Explicabilite"]


def test_analyse_ne_reference_plus_les_lettres_des_metriques_data():
    """Analyse doit utiliser les libellés, jamais les anciennes lettres DATA."""
    motif = re.compile(
        r"DATA!\$?(?:DZ|EA|AS|CC|BE|I|DA|CO|DS|F)\$?2:"
    )
    with zipfile.ZipFile(TEMPLATE) as archive:
        partie = _parties_feuilles(archive)["Analyse"]
        root = etree.fromstring(archive.read(partie))
        formules = [
            item.text or ""
            for item in root.findall(f".//{{{NS_MAIN}}}f")
        ]

    assert not any(motif.search(formule) for formule in formules)
    assert not any("_xlfn." in formule for formule in formules)


def test_analyse_compte_uniquement_les_titres_non_vides():
    """Le compteur de titres doit ignorer les cellules contenant une formule vide."""
    with zipfile.ZipFile(TEMPLATE) as archive:
        parties = _parties_feuilles(archive)
        formules = {
            adresse: _formule(archive, parties["Analyse"], adresse)
            for adresse in ("C6", "D6")
        }

    assert all('COUNTIF(' in formule and '"<>"' in formule for formule in formules.values())
    assert not any("COUNTA(" in formule for formule in formules.values())


def test_analyse_contient_la_table_de_sensibilite_te():
    """Le modèle réserve une zone stable pour les scénarios de Tracking Error."""
    with zipfile.ZipFile(TEMPLATE) as archive:
        parties = _parties_feuilles(archive)
        root = etree.fromstring(archive.read(parties["Analyse"]))

        def valeur(adresse):
            cellule = root.find(f".//{{{NS_MAIN}}}c[@r='{adresse}']")
            inline = cellule.find(f"{{{NS_MAIN}}}is/{{{NS_MAIN}}}t")
            return inline.text if inline is not None else None

    assert valeur("P1") == "Sensibilité Tracking Error"
    assert valeur("P2") == "Scénario"
    assert valeur("U2") == "TE annualisé"
    assert valeur("V2") == "Poids PTF manquant / retiré"
    assert valeur("W2") == "Poids Bench manquant / retiré"
    assert valeur("P10") == "EWMA lambda 0,97"
