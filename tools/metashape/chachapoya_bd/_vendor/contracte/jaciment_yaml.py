"""
Contracte de ``jaciment.yaml`` (versió 2): el carnet d'identitat compartit del jaciment, per blocs, i tot el que el
rodeja: ``Jaciment``/``Campanya``, lectura i escriptura de ``jaciment.yaml`` i ``shift.txt``, les campanyes a disc i les
claus ``jaciment/*`` del document de Metashape.

Mòdul de contracte (``docs/CONTRACTE_JACIMENT.md``): còpia idèntica a Oculat (``pipeline/contracte/``). Sense PyYAML,
sense Metashape, sense bpy; **només importa altres mòduls de contracte** (``claus``, ``utm``) i la biblioteca
estàndard (``comu/tests/test_contracte_imports.py`` ho comprova). Si canvia, canvia a tots dos repositoris alhora.
``lliurament.py`` ho reexporta tot perquè cap consumidor haja de canviar.

El fitxer té **propietat per blocs**, com ``[tool.*]`` de ``pyproject.toml``::

    versio: 2                        ┐
    nom: "Sanxo Llop"                │ capçalera: la poden escriure els dos programes
    codi: "SLL"                      │ (versio, nom, codi, epsg, shift, campanyes)
    epsg: 25830                      │
    shift: [745000, 4316900, 0]      │
    campanyes:                       │ una per projecte .psx
      - carpeta: "2024_sll"          ┘
    nomenclatura:                    compartit (separador, camp_data_format)
      separador: "_"
    oculat:                          bloc d'Oculat: Pontarch NO el toca, NO el valida, el preserva byte a byte
      ...
    pontarch:                        bloc de Pontarch: Oculat ídem
      plantilles: [...]

Qui escriu passa sempre per ``reemplacar_bloc`` (o ``reemplacar_capcalera``): es reescriu només el bloc propi i la
resta del text es copia **verbatim**, comentaris inclosos. Ordre canònic en crear: capçalera, ``nomenclatura``,
``oculat``, ``pontarch``; els blocs desconeguts es preserven on estaven. La versió 1 (``projectes: [...]``) es llig i
es migra a ``campanyes`` (només ``carpeta``) quan algú escriu la capçalera, mai en silenci (qui escriu rep l'avís).

Coordenada al fitxer = coordenada de mapa − shift. El shift és «Shift X/Y/Z (origen local)», com al diàleg *Export
Model* de Metashape: enters en metres arredonits avall a centenars, ``Z0 = 0`` per defecte. Un sol shift per
jaciment: si ``jaciment.yaml`` ja existeix, mana ell.

**Shift pendent** (contracte 2.1): ``Jaciment.shift`` pot ser ``None``. Un ``jaciment.yaml`` sense ``shift:`` o amb
``shift: null`` es llig com a pendent; ``[0, 0, 0]`` explícit és pendent **només** si cap campanya té fitxers a disc
(``projectes_de`` buit; si en té, és un valor real i ``shift_plausible`` decideix si avisa). La capçalera s'escriu
``shift: null   # pendent: es fixa en el primer vincle amb chunks``; ``shift.txt`` no s'escriu mentre siga pendent.
Fixar un shift pendent no és canviar el marc (``marc_compatible``): no cal «Editar la identitat…».

API del text: ``blocs``, ``escalars`` (lector d'un nivell: escalars, mapes d'escalars i llistes de mapes),
``reemplacar_bloc``, ``eliminar_bloc``, ``reemplacar_capcalera``, ``text_capcalera``, ``text_bloc``, ``text_nou``,
``VERSIO_FITXER``. API del jaciment: ``Jaciment``, ``Campanya``, ``normalitzar_*``, ``proposar_shift``,
``shift_plausible``, ``epsg_inicial``, ``llegir_lliurament``, ``desar_lliurament`` (capçalera + ``shift.txt``),
``projectes_de``, ``llegir_jaciment_meta``/``escriure_jaciment_meta`` (claus compartides ``jaciment/*``),
``resoldre_any``.
"""

from __future__ import annotations

import math
import re
import unicodedata
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Iterable, Sequence

from . import claus as _claus
from .utm import candidats_zona, comprovar_epsg, zona_utm_des_de_lonlat

VERSIO_FITXER = 2
CLAUS_CAPCALERA = ("versio", "nom", "codi", "epsg", "shift", "campanyes")
CLAUS_CAPCALERA_V1 = ("projectes",)               # versió 1: es migra a campanyes
CLAU_NOMENCLATURA = "nomenclatura"
BLOC_OCULAT = "oculat"
BLOC_PONTARCH = "pontarch"
ORDRE_CANONIC = ("capcalera", CLAU_NOMENCLATURA, BLOC_OCULAT, BLOC_PONTARCH)
CAPCALERA = "capcalera"                           # nom simbòlic del conjunt de CLAUS_CAPCALERA a ``blocs``/``ordre``
FONTS_ANY = ("exif", "nom_projecte", "manual")
COMENTARI_CAPCALERA = (
    "# Carnet d'identitat del jaciment, compartit per Oculat i Pontarch. Coordenada als fitxers = coordenada de mapa - shift.",
    "# No canvies l'EPSG ni el shift quan ja hi ha lliuraments: deixarien de coincidir. Cada programa escriu només el seu bloc.",
)

FITXER_JACIMENT = "jaciment.yaml"
FITXER_SHIFT = "shift.txt"
SUBCARPETA_3D = "3D"
SUBCARPETA_ORTO = "ORTO"
SUBCARPETA_DEM = "DEM"
SUBCARPETES_PROJECTE = (SUBCARPETA_3D, SUBCARPETA_ORTO, SUBCARPETA_DEM)   # <jaciment>/<campanya>/: només el que ix de Metashape
NOM_PROJECTE_DEFECTE = "projecte"
PAS_SHIFT = 100

CLAU_META_LLIURAMENT = "lliurament"   # doc.meta['pontarch/lliurament']: heretada (0.5–0.6); es llig, no s'escriu
CLAU_CARPETA = "carpeta"              # doc.meta['jaciment/carpeta']: ruta de la carpeta de jaciment (compartida amb Oculat)
CLAU_CAMPANYA = "campanya"            # doc.meta['jaciment/campanya']: subcarpeta de lliurament del projecte
FONTS_VINCLE = ("compartida", "heretada", "cap")

FONTS_EPSG = {                        # les que mostra qt_contracte.SelectorCrsUtm
    "jaciment": f"de {FITXER_JACIMENT}",
    "chunk": "CRS del chunk",
    "detectat": "zona detectada",
    "manual": "manual",
    "cap": "",
}
FONTS_EPSG_SELECTOR = tuple(FONTS_EPSG)

E_UTM_MIN, E_UTM_MAX = 100_000.0, 900_000.0
N_UTM_MIN, N_UTM_MAX = 0.0, 10_000_000.0
RADI_PLAUSIBLE_M = 50_000.0

Num = int | float
TEXT_SHIFT_PENDENT = "pendent"
COMENTARI_SHIFT_PENDENT = "# pendent: es fixa en el primer vincle amb chunks"

_RE_CLAU = re.compile(r"^([A-Za-z_][\w\-]*)\s*:(?:\s|$)")


# ── Text: comentaris i valors ────────────────────────────────────────────────


def _sense_comentari(linia: str) -> str:
    """La línia sense el comentari ``#`` (respecta les cometes)."""
    dins = None
    for i, ch in enumerate(linia):
        if dins:
            if ch == dins:
                dins = None
        elif ch in "\"'":
            dins = ch
        elif ch == "#" and (i == 0 or linia[i - 1] in " \t"):
            return linia[:i]
    return linia


def _trossos_inline(cos: str) -> list[str]:
    """Elements d'una llista en línia, separats per comes fora de claudàtors i cometes."""
    out: list[str] = []
    actual: list[str] = []
    prof = 0
    dins = None
    for ch in cos:
        if dins:
            actual.append(ch)
            if ch == dins:
                dins = None
            continue
        if ch in "\"'":
            dins = ch
        elif ch == "[":
            prof += 1
        elif ch == "]":
            prof -= 1
        if ch == "," and prof == 0:
            out.append("".join(actual))
            actual = []
        else:
            actual.append(ch)
    if "".join(actual).strip():
        out.append("".join(actual))
    return [t for t in out if t.strip()]


def _valor_simple(v: str) -> Any:
    """Escalar YAML (text amb o sense cometes, enter, decimal, booleà, null) o llista en línia (niada si cal)."""
    v = v.strip()
    if v.startswith("[") and v.endswith("]"):
        return [_valor_simple(x) for x in _trossos_inline(v[1:-1])]
    if len(v) >= 2 and v[0] == v[-1] and v[0] in "\"'":
        cos = v[1:-1]
        return cos.replace('\\"', '"').replace("\\\\", "\\") if v[0] == '"' else cos.replace("''", "'")
    if v in ("", "~", "null", "Null", "NULL"):
        return None
    if v in ("true", "True", "TRUE"):
        return True
    if v in ("false", "False", "FALSE"):
        return False
    try:
        return int(v)
    except ValueError:
        pass
    try:
        return float(v)
    except ValueError:
        return v


def _yaml_str(s: str) -> str:
    return '"' + str(s).replace("\\", "\\\\").replace('"', '\\"') + '"'


def _num(v: Any) -> str:
    f = float(v)
    return str(int(f)) if f.is_integer() else f"{f:g}"


def _yaml_valor(v: Any) -> str:
    """Valor escalar o llista en línia tal com s'escriu."""
    if v is None:
        return "null"
    if isinstance(v, bool):
        return "true" if v else "false"
    if isinstance(v, (int, float)):
        return _num(v) if isinstance(v, float) else str(v)
    if isinstance(v, (list, tuple)):
        return "[" + ", ".join(_yaml_valor(x) for x in v) + "]"
    return _yaml_str(str(v))


def _indent(linia: str) -> int:
    return len(linia) - len(linia.lstrip(" "))


# ── Blocs de primer nivell ───────────────────────────────────────────────────


def blocs(text: str) -> list[tuple[str, str]]:
    """``[(clau, text del bloc), …]`` en l'ordre del fitxer: cada clau de primer nivell amb totes les seues línies
    (les indentades, els ítems de llista, els comentaris que la precedeixen). Les línies anteriors al primer bloc
    (comentaris de capçalera del fitxer) van amb la clau ``""``. Els blancs que separen blocs queden al final del
    bloc anterior, perquè reemplaçar un bloc no els faça desaparéixer. ``"".join(textos) == text``."""
    out: list[tuple[str, str]] = []
    clau_actual: str | None = None
    linies_actual: list[str] = []
    pendents: list[str] = []            # blancs i comentaris a columna 0: es reparteixen quan arriba la clau següent

    def tancar(nou_inici: list[str]) -> None:
        nonlocal linies_actual
        if clau_actual is not None or linies_actual:
            out.append((clau_actual or "", "".join(linies_actual)))
        linies_actual = list(nou_inici)

    for linia in text.splitlines(keepends=True):
        m = _RE_CLAU.match(linia)
        if m:
            # blancs inicials dels pendents → bloc anterior; des del primer comentari → bloc nou
            i = 0
            while i < len(pendents) and not pendents[i].strip():
                i += 1
            if clau_actual is None and not linies_actual and pendents:
                # abans del primer bloc tot és capçalera del fitxer (clau "")
                linies_actual = list(pendents)
                tancar([])
            else:
                linies_actual += pendents[:i]
                tancar(pendents[i:])
            pendents = []
            clau_actual = m.group(1)
            linies_actual.append(linia)
        elif not linia.strip() or linia.startswith("#"):
            pendents.append(linia)
        else:
            linies_actual += pendents
            pendents = []
            linies_actual.append(linia)
    linies_actual += pendents
    tancar([])
    return out


def claus(text: str) -> list[str]:
    """Claus de primer nivell del text, en ordre."""
    return [c for c, _ in blocs(text) if c]


def _rang(clau: str) -> int | None:
    """Posició en l'ordre canònic (capçalera 0, nomenclatura 1, oculat 2, pontarch 3); None si és desconeguda."""
    if clau in CLAUS_CAPCALERA or clau in CLAUS_CAPCALERA_V1:
        return 0
    try:
        return ORDRE_CANONIC.index(clau)
    except ValueError:
        return None


def _amb_salt(t: str) -> str:
    return t if (not t or t.endswith("\n")) else t + "\n"


def _separat(anterior: str, nou: str) -> str:
    """Una línia en blanc entre un bloc i el següent si l'anterior no en deixa."""
    if anterior and not anterior.endswith("\n\n") and anterior.strip():
        return "\n" + nou
    return nou


def _inserir(parts: list[tuple[str, str]], rang_nou: int, textos_nous: list[str]) -> list[tuple[str, str]]:
    """Insereix els textos nous després de l'últim bloc conegut de rang ≤ ``rang_nou`` (o de la capçalera del fitxer)."""
    pos = 0
    for i, (c, _t) in enumerate(parts):
        r = _rang(c) if c else -1
        if r is not None and r <= rang_nou:
            pos = i + 1
    anterior = parts[pos - 1][1] if pos else ""
    nous = []
    for k, t in enumerate(textos_nous):
        nous.append(("*", _separat(anterior if k == 0 else nous[-1][1], _amb_salt(t))))
    # el bloc que ve després també vol quedar separat
    if pos < len(parts) and nous:
        seguent_clau, seguent_text = parts[pos]
        if seguent_clau and not nous[-1][1].endswith("\n\n") and not seguent_text.startswith(("\n", "#")):
            nous[-1] = (nous[-1][0], nous[-1][1] + "\n")
    return parts[:pos] + nous + parts[pos:]


def reemplacar_bloc(text: str, clau: str, nou_bloc: str) -> str:
    """El text amb el bloc ``clau`` reescrit amb ``nou_bloc`` (el text sencer del bloc, ``clau:`` inclosa) i **tota la
    resta verbatim**. Si el bloc no hi era, s'afig en la seua posició canònica (una clau desconeguda, al final)."""
    if clau in CLAUS_CAPCALERA or clau in CLAUS_CAPCALERA_V1:
        raise ValueError(f"'{clau}' és de la capçalera: usa reemplacar_capcalera")
    parts = blocs(text)
    nou_bloc = _amb_salt(nou_bloc)
    if any(c == clau for c, _ in parts):
        out = []
        for c, t in parts:
            if c == clau:
                # conserva els blancs finals del bloc vell (separen el següent)
                cua = t[len(t.rstrip("\n")):]
                cua = cua[1:] if cua.startswith("\n") else cua
                out.append((c, nou_bloc + cua))
            else:
                out.append((c, t))
        return "".join(t for _, t in out)
    rang = _rang(clau)
    if rang is None:
        rang = len(ORDRE_CANONIC)
    return "".join(t for _, t in _inserir(parts, rang, [nou_bloc]))


def eliminar_bloc(text: str, clau: str) -> str:
    return "".join(t for c, t in blocs(text) if c != clau)


def reemplacar_capcalera(text: str, text_cap: str) -> str:
    """Reescriu la capçalera (``CLAUS_CAPCALERA`` i la ``projectes`` de la v1) amb ``text_cap`` al lloc de la primera
    clau de capçalera que hi havia; els altres blocs (``nomenclatura``, ``oculat``, ``pontarch``, desconeguts) i els
    comentaris del fitxer no es toquen."""
    parts = blocs(text)
    text_cap = _amb_salt(text_cap)
    es_cap = [bool(c) and _rang(c) == 0 for c, _ in parts]
    if not any(es_cap):
        return "".join(t for _, t in _inserir(parts, 0, [text_cap]))
    out: list[tuple[str, str]] = []
    posada = False
    for (c, t), cap in zip(parts, es_cap):
        if not cap:
            out.append((c, t))
            continue
        if not posada:
            # els comentaris que precedien la primera clau de capçalera (del fitxer, no del nostre text) es conserven
            inici = re.match(r"((?:[ \t]*\n|#[^\n]*\n)*)", t).group(1)
            out.append((c, inici + text_cap))
            posada = True
    # el blanc que separava la capçalera del següent bloc
    if out and posada:
        i = next(i for i, (c, _t) in enumerate(out) if c and _rang(c) == 0)
        if i + 1 < len(out) and not out[i][1].endswith("\n\n"):
            seguent = out[i + 1][1]
            if not seguent.startswith("\n"):
                out[i] = (out[i][0], out[i][1] + "\n")
    return "".join(t for _, t in out)


# ── Lector d'un nivell ───────────────────────────────────────────────────────


def escalars(text: str) -> dict[str, Any]:
    """``dict`` de les claus de primer nivell: escalars i llistes en línia, **i un nivell més**: un bloc niat
    d'escalars és un ``dict``, una llista de bloc (``- clau: valor`` amb continuacions, o ``- escalar``) és una
    ``list``. Més fondària s'ignora. Sense PyYAML (és el mateix lector que vendoritza Blender)."""
    out: dict[str, Any] = {}
    clau_top: str | None = None
    ind_fill: int | None = None                  # indentació dels fills directes del bloc obert
    item: dict | None = None                     # ítem de llista (mapa) en construcció
    ind_item: int | None = None                  # indentació de les claus de l'ítem
    for raw in text.splitlines():
        linia = _sense_comentari(raw.replace("\t", "  ")).rstrip()
        if not linia.strip():
            continue
        ind = _indent(linia)
        s = linia.strip()
        if ind == 0:
            m = _RE_CLAU.match(linia)
            if m:
                clau, _, valor = linia.partition(":")
                clau = clau.strip()
                if valor.strip():
                    out[clau] = _valor_simple(valor)
                    clau_top = None
                else:
                    out[clau] = None
                    clau_top, ind_fill, item, ind_item = clau, None, None, None
            elif s.startswith("-") and clau_top is not None:
                _item_llista(out, clau_top, s)
                item, ind_item = (out[clau_top][-1], None) if isinstance(out[clau_top][-1], dict) else (None, None)
            continue
        if clau_top is None:
            continue
        if s == "-" or s.startswith("- "):
            if ind_fill is None:
                ind_fill = ind
            if ind != ind_fill:
                continue                          # llista dins d'un nivell més fondo: s'ignora
            _item_llista(out, clau_top, s)
            ultim = out[clau_top][-1]
            item, ind_item = (ultim, None) if isinstance(ultim, dict) else (None, None)
            continue
        m = _RE_CLAU.match(s)
        if not m:
            continue
        k, _, v = s.partition(":")
        k = k.strip()
        if isinstance(out[clau_top], list):
            if item is None or ind <= (ind_fill or 0):
                continue
            if ind_item is None:
                ind_item = ind
            if ind == ind_item:
                item[k] = _valor_simple(v)
            continue
        if ind_fill is None:
            ind_fill = ind
        if ind != ind_fill:
            continue                              # tercer nivell: s'ignora
        if out[clau_top] is None:
            out[clau_top] = {}
        out[clau_top][k] = _valor_simple(v)
    return out


def _item_llista(out: dict, clau_top: str, s: str) -> None:
    if not isinstance(out[clau_top], list):
        out[clau_top] = []
    cos = s[1:].strip()
    if cos and _RE_CLAU.match(cos):
        k, _, v = cos.partition(":")
        out[clau_top].append({k.strip(): _valor_simple(v)})
    else:
        out[clau_top].append(_valor_simple(cos))


# ── Escriptura ───────────────────────────────────────────────────────────────


def text_bloc(clau: str, linies: Iterable[str] | dict[str, Any]) -> str:
    """``clau:`` + les línies indentades dues posicions. ``linies`` pot ser un ``dict`` d'escalars/llistes en línia
    (una línia per clau) o les línies ja escrites (sense la indentació del bloc)."""
    if isinstance(linies, dict):
        linies = [f"{k}: {_yaml_valor(v)}" for k, v in linies.items()]
    cos = "".join(f"  {ln}\n" if ln.strip() else "\n" for ln in linies)
    return f"{clau}:\n{cos}"


def _linies_campanya(c: Any) -> list[str]:
    out = [f"- carpeta: {_yaml_str(c.carpeta)}"]
    if getattr(c, "nom", None):
        out.append(f"  nom: {_yaml_str(c.nom)}")
    if getattr(c, "any", None) is not None:
        out.append(f"  any: {int(c.any)}")
    if getattr(c, "any_font", None):
        out.append(f"  any_font: {c.any_font}")
    if getattr(c, "dates", None):
        out.append("  dates: [" + ", ".join(_yaml_str(d) for d in c.dates) + "]")
    return out


def text_capcalera(j: Any) -> str:
    """Capçalera (sempre ``versio: 2``) d'un ``Jaciment`` (``nom``, ``codi``, ``epsg``, ``shift``, ``campanyes``)."""
    linies = [f"versio: {VERSIO_FITXER}", f"nom: {_yaml_str(j.nom)}"]
    if getattr(j, "codi", None):
        linies.append(f"codi: {_yaml_str(j.codi)}")
    linies.append(f"epsg: {int(j.epsg)}   # CRS projectat (UTM)")
    if j.shift is None:
        linies.append(f"shift: null   {COMENTARI_SHIFT_PENDENT}")
    else:
        x, y, z = (_num(v) for v in j.shift)
        linies.append(f"shift: [{x}, {y}, {z}]   # Shift X/Y/Z (origen local), metres")
    campanyes = list(getattr(j, "campanyes", None) or [])
    if campanyes:
        linies.append("campanyes:   # una per projecte .psx; carpeta = subcarpeta de lliurament")
        for c in campanyes:
            linies += ["  " + ln for ln in _linies_campanya(c)]
    else:
        linies.append("campanyes: []   # una per projecte .psx; carpeta = subcarpeta de lliurament")
    return "\n".join(linies) + "\n"


def text_nou(j: Any, nomenclatura: dict[str, Any] | None = None, pontarch: dict[str, Any] | None = None) -> str:
    """Fitxer nou en ordre canònic: comentaris, capçalera, ``nomenclatura`` i ``pontarch`` si es donen (``oculat``
    només l'escriu Oculat)."""
    text = "\n".join(COMENTARI_CAPCALERA) + "\n" + text_capcalera(j)
    if nomenclatura:
        text += "\n" + text_bloc(CLAU_NOMENCLATURA, nomenclatura)
    if pontarch:
        text += "\n" + text_bloc(BLOC_PONTARCH, pontarch)
    return text


# ── Any de la campanya ───────────────────────────────────────────────────────


def resoldre_any(dates_iso: Sequence[str] | None, nom_projecte: str | None) -> tuple[int | None, str, tuple[str, str] | None]:
    """``(any, font, (data mínima, data màxima))`` d'una campanya: l'any més freqüent de ``dates_iso`` (``AAAA-MM-DD``;
    empat → el més antic; font ``exif``), si no un ``20\\d\\d`` al nom del projecte (font ``nom_projecte``), si no
    ``(None, "cap", None)``. Les dates no ISO s'ignoren."""
    dates = sorted({str(d).strip()[:10] for d in (dates_iso or []) if re.match(r"^\d{4}-\d{2}-\d{2}", str(d).strip())})
    if dates:
        recompte: dict[int, int] = {}
        for d in dates:
            a = int(d[:4])
            recompte[a] = recompte.get(a, 0) + 1
        any_ = sorted(recompte, key=lambda a: (-recompte[a], a))[0]
        return any_, "exif", (dates[0], dates[-1])
    m = re.search(r"(?<!\d)(20\d\d)(?!\d)", nom_projecte or "")
    if m:
        return int(m.group(1)), "nom_projecte", None
    return None, "cap", None


# ── Campanya i Jaciment ──────────────────────────────────────────────────────


@dataclass
class Campanya:
    """Una campanya = un projecte ``.psx``: ``carpeta`` és la seua subcarpeta de lliurament; ``any`` amb la seua font
    (``exif`` | ``nom_projecte`` | ``manual``) i ``dates`` ``[inici, fi]`` ISO són opcionals."""
    carpeta: str
    nom: str | None = None
    any: int | None = None
    any_font: str | None = None
    dates: list[str] | None = None

    def __post_init__(self):
        self.carpeta = str(self.carpeta or "").strip()
        if not self.carpeta:
            raise ValueError("Una campanya necessita la carpeta (subcarpeta de lliurament del projecte).")
        self.nom = (str(self.nom).strip() or None) if self.nom else None
        try:
            self.any = int(self.any) if self.any is not None and not isinstance(self.any, bool) else None
        except (TypeError, ValueError):
            self.any = None
        self.any_font = str(self.any_font) if self.any_font in FONTS_ANY else None
        if self.dates:
            ds = [str(d).strip()[:10] for d in (self.dates if isinstance(self.dates, (list, tuple)) else [self.dates])]
            ds = [d for d in ds if d]
            self.dates = [ds[0], ds[-1]] if ds else None
        else:
            self.dates = None

    def fusionada(self, altra: "Campanya") -> "Campanya":
        """Esta campanya amb els camps no buits d'``altra`` (mateixa carpeta) per damunt."""
        return Campanya(self.carpeta, altra.nom or self.nom,
                        altra.any if altra.any is not None else self.any,
                        altra.any_font or (self.any_font if altra.any is None else None),
                        altra.dates or self.dates)


def campanya_de(v: Any) -> Campanya | None:
    """``Campanya`` d'un ``Campanya``, un ``dict`` (``carpeta``, ``nom``, ``any``, ``any_font``, ``dates``) o un nom de
    carpeta (v1); None si no en queda res."""
    if v is None:
        return None
    if isinstance(v, Campanya):
        return v
    if isinstance(v, dict):
        if not str(v.get("carpeta") or "").strip():
            return None
        return Campanya(v.get("carpeta"), v.get("nom"), v.get("any"), v.get("any_font"), v.get("dates"))
    s = str(v or "").strip()
    return Campanya(s) if s else None


def normalitzar_campanyes(campanyes: Any) -> list[Campanya]:
    """Llista sense carpetes repetides (la darrera entrada d'una carpeta completa l'anterior), en l'ordre donat; un
    escalar sol també val."""
    if campanyes is None:
        return []
    if isinstance(campanyes, (str, dict, Campanya)):
        campanyes = [campanyes]
    out: list[Campanya] = []
    for v in campanyes:
        c = campanya_de(v)
        if c is None:
            continue
        for i, prev in enumerate(out):
            if prev.carpeta == c.carpeta:
                out[i] = prev.fusionada(c)
                break
        else:
            out.append(c)
    return out


@dataclass
class Jaciment:
    nom: str
    epsg: int
    shift: list[Num] | None                    # None = pendent (es fixa en el primer vincle amb chunks)
    codi: str | None = None
    versio: int = VERSIO_FITXER
    campanyes: list[Campanya] = field(default_factory=list)    # una per projecte .psx, en ordre d'alta

    def __post_init__(self):
        self.epsg = int(self.epsg)
        self.shift = None if self.shift is None else normalitzar_shift(self.shift)
        self.nom = str(self.nom or "").strip()
        self.codi = (str(self.codi).strip() or None) if self.codi else None
        self.campanyes = normalitzar_campanyes(self.campanyes)

    @property
    def shift_pendent(self) -> bool:
        return self.shift is None

    @property
    def projectes(self) -> list[str]:
        """Subcarpetes de projecte (``carpeta`` de cada campanya): el que llegien els consumidors de la v1."""
        return [c.carpeta for c in self.campanyes]

    def campanya(self, carpeta: str | None) -> Campanya | None:
        return next((c for c in self.campanyes if c.carpeta == carpeta), None)

    def amb_campanya(self, campanya: Campanya | str | dict | None) -> "Jaciment":
        """Còpia amb la campanya afegida (o completada, si la carpeta ja hi era)."""
        c = campanya_de(campanya)
        if c is None:
            return self
        return Jaciment(self.nom, self.epsg, _copia_shift(self.shift), self.codi, self.versio, self.campanyes + [c])

    def amb_projecte(self, projecte: str | None) -> "Jaciment":
        """Embolcall de ``amb_campanya`` per als consumidors de la v1 (només la carpeta)."""
        return self.amb_campanya(projecte)

    def amb_campanyes_de_disc(self, a_disc: Iterable[str]) -> "Jaciment":
        """Còpia amb les subcarpetes de campanya trobades a disc (``projectes_de``) que no estaven declarades,
        afegides al final amb només ``carpeta`` (sense any: ``any_font`` queda buida). Les declarades no es toquen."""
        noves = [Campanya(n) for n in a_disc if n and self.campanya(n) is None]
        if not noves:
            return self
        return Jaciment(self.nom, self.epsg, _copia_shift(self.shift), self.codi, self.versio, self.campanyes + noves)

    def amb_shift(self, shift: Sequence[Num] | None) -> "Jaciment":
        """Còpia amb un altre shift (None = pendent)."""
        return Jaciment(self.nom, self.epsg, _copia_shift(shift), self.codi, self.versio, list(self.campanyes))


def _copia_shift(shift: Sequence[Num] | None) -> list[Num] | None:
    return None if shift is None else list(shift)


def normalitzar_projectes(projectes: Any) -> list[str]:
    """Llista de noms no buits, sense repetits, en l'ordre donat (un escalar sol també val)."""
    return [c.carpeta for c in normalitzar_campanyes(projectes)]


# ── Shift ────────────────────────────────────────────────────────────────────


def normalitzar_shift(shift: Sequence[Any] | None) -> list[Num]:
    """Tres nombres; els enters queden com a ``int`` (perquè s'escriguen sense decimals)."""
    if shift is None or len(shift) != 3:
        raise ValueError("El shift ha de tenir tres components [X0, Y0, Z0].")
    out: list[Num] = []
    for v in shift:
        f = float(v)
        if not math.isfinite(f):
            raise ValueError(f"Component de shift no vàlid: {v!r}")
        out.append(int(f) if f.is_integer() else f)
    return out


def _punts(bboxes: Iterable[Any]):
    for b in bboxes:
        punts = [b] if len(b) == 3 and not hasattr(b[0], "__len__") else list(b)
        for p in punts:
            yield p


def proposar_shift(bboxes: Iterable[Any], pas: int = PAS_SHIFT) -> list[int]:
    """Shift proposat: centre de la unió dels bbox (o punts) arredonit avall a ``pas`` m; Z0 = 0.

    ``bboxes``: elements ``(min_xyz, max_xyz)`` o punts ``(x, y, z)`` en coordenades de mapa.
    """
    xs: list[float] = []
    ys: list[float] = []
    for p in _punts(bboxes):
        xs.append(float(p[0]))
        ys.append(float(p[1]))
    if not xs:
        raise ValueError("Cap chunk georeferenciat per proposar el shift.")
    cx = (min(xs) + max(xs)) / 2
    cy = (min(ys) + max(ys)) / 2
    return [int(math.floor(cx / pas) * pas), int(math.floor(cy / pas) * pas), 0]


def shift_plausible(shift: Sequence[Num] | None, bboxes_projectats: Iterable[Any]) -> bool:
    """El shift és un punt UTM vàlid (E 100 000–900 000, N 0–10 000 000) i tots els bbox (o punts) en el
    CRS de mapa queden a menys de ±50 km d'ell en X i Y."""
    try:
        x0, y0, _z0 = normalitzar_shift(shift)
    except ValueError:
        return False
    if not (E_UTM_MIN <= x0 <= E_UTM_MAX and N_UTM_MIN <= y0 <= N_UTM_MAX):
        return False
    for p in _punts(bboxes_projectats):
        if abs(float(p[0]) - x0) > RADI_PLAUSIBLE_M or abs(float(p[1]) - y0) > RADI_PLAUSIBLE_M:
            return False
    return True


def motiu_shift(shift: Sequence[Num] | None, bboxes_projectats: Iterable[Any]) -> str | None:
    """Per què el shift no és plausible: ``pendent`` (None: no és un error de l'usuari), ``fora del rang UTM``,
    ``lluny dels chunks`` (> 50 km) o ``no vàlid``; None si és plausible."""
    if shift is None:
        return TEXT_SHIFT_PENDENT
    try:
        x0, y0, _z0 = normalitzar_shift(shift)
    except ValueError:
        return "no vàlid"
    if not (E_UTM_MIN <= x0 <= E_UTM_MAX and N_UTM_MIN <= y0 <= N_UTM_MAX):
        return "fora del rang UTM"
    if not shift_plausible(shift, bboxes_projectats):
        return "lluny dels chunks"
    return None


def text_shift(shift: Sequence[Num] | None) -> str:
    """«(745000, 4316900, 0)», o «pendent» si el shift encara no està fixat."""
    if shift is None:
        return TEXT_SHIFT_PENDENT
    return "(" + ", ".join(_num(v) for v in shift) + ")"


# ── EPSG ─────────────────────────────────────────────────────────────────────


def epsg_de_text(text: str | None) -> int | None:
    """``EPSG::32718`` / ``EPSG:32718`` / ``32718`` → 32718; None si no és un codi. (Còpia de la d'``exportar_ortos``
    perquè el contracte no l'importe.)"""
    if not text:
        return None
    m = re.search(r"(\d{4,6})\s*$", str(text).strip())
    return int(m.group(1)) if m else None


def epsg_inicial(epsg_chunk: int | None, chunk_projectat: bool,
                 lonlat: tuple[float, float] | None) -> tuple[int | None, str]:
    """``(epsg, font)`` del selector quan no hi ha ``jaciment.yaml``: el CRS del chunk si és projectat **i** de la zona
    UTM dels chunks (font ``chunk``); si no, el primer candidat de la zona detectada (``candidats_zona``: ETRS89 a
    Europa, si no WGS 84; font ``detectat``); sense chunks georeferenciats, ``(None, "cap")``. Cap EPSG al codi."""
    if epsg_chunk and chunk_projectat and (lonlat is None or comprovar_epsg(int(epsg_chunk), *lonlat) is None):
        return int(epsg_chunk), "chunk"
    if lonlat is not None:
        _e, zona, hemi = zona_utm_des_de_lonlat(*lonlat)
        return candidats_zona(zona, hemi)[0][0], "detectat"
    return None, "cap"


def avis_epsg_carregat(epsg: int, lonlat: tuple[float, float] | None) -> str | None:
    """Avís si l'EPSG carregat de ``jaciment.yaml`` no és de la zona UTM dels chunks marcats (mateix criteri que
    ``comprovar_epsg``); None si coincideix, si no es pot comprovar o si no hi ha chunks."""
    if lonlat is None:
        return None
    avis = comprovar_epsg(int(epsg), *lonlat)
    if not avis:
        return None
    return (f"{FITXER_JACIMENT} carregat: {avis} Els chunks marcats no són d'este jaciment o el fitxer és erroni: "
            "no es pot exportar amb este marc.")


# ── Noms de fitxer i de campanya ─────────────────────────────────────────────


def nom_segur(label: str | None, fallback: str) -> str:
    """Text → nom de fitxer segur: els caràcters prohibits passen a ``_``. (Còpia de la d'``ortomosaics`` perquè el
    contracte no l'importe.)"""
    name = re.sub(r"[^\w\-. ]+", "_", (label or "").strip())
    name = re.sub(r"\s+", "_", name).strip("._")
    return name or fallback


def nom_fitxer(etiqueta: str | None, fallback: str = "Chunk") -> str:
    """Etiqueta de chunk → nom de fitxer segur: sense accents (transliterats), espais i caràcters prohibits → ``_``."""
    s = unicodedata.normalize("NFKD", etiqueta or "")
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    s = nom_segur(s, fallback)
    return re.sub(r"_+", "_", s).strip("_") or fallback


def nom_projecte(ruta_psx: str | Path | None) -> str:
    """Subcarpeta de campanya per defecte: el nom del ``.psx`` (sense extensió, normalitzat com els noms de
    fitxer); ``projecte`` si el document no està desat."""
    stem = Path(ruta_psx).stem if ruta_psx else ""
    return nom_fitxer(stem, NOM_PROJECTE_DEFECTE)


def nom_projecte_segur(text: str | None) -> str:
    """El nom escrit al diàleg com a nom de carpeta (sense accents ni caràcters prohibits); '' si no en queda res."""
    return nom_fitxer(text, "")


# ── jaciment.yaml → Jaciment ─────────────────────────────────────────────────


def text_jaciment_yaml(j: Jaciment, nomenclatura: dict | None = None, pontarch: dict | None = None) -> str:
    """Fitxer **nou** en ordre canònic (``text_nou``). Sobre un fitxer existent no s'usa mai: la capçalera es
    reescriu amb ``reemplacar_capcalera`` i la resta queda verbatim."""
    return text_nou(j, nomenclatura, pontarch)


def parse_yaml_simple(text: str) -> dict[str, Any]:
    """Lector sense PyYAML de ``jaciment.yaml`` (``escalars``: escalars, un nivell de mapes i llistes de mapes); el
    mateix que vendoritza Blender."""
    return escalars(text)


def jaciment_de_dict(d: dict[str, Any], nom_defecte: str = "") -> Jaciment:
    """``Jaciment`` del dict llegit: versió 2 (``campanyes``) o versió 1 (``projectes``: només carpetes)."""
    epsg = d.get("epsg")
    if not isinstance(epsg, int) or isinstance(epsg, bool) or epsg <= 0:
        raise ValueError(f"{FITXER_JACIMENT}: 'epsg' ha de ser un enter positiu (hi ha {epsg!r}).")
    shift = d.get("shift")
    if shift is None and d.get("origen") is not None:
        raise ValueError(f"{FITXER_JACIMENT}: la clau és 'shift' (no 'origen').")
    if shift is not None and not isinstance(shift, (list, tuple)):
        raise ValueError(f"{FITXER_JACIMENT}: 'shift' ha de ser una llista [X0, Y0, Z0] (o null: pendent).")
    campanyes = d.get("campanyes")
    if campanyes is None:
        campanyes = d.get("projectes")
    versio = d.get("versio")
    if not isinstance(versio, int) or isinstance(versio, bool):
        versio = 1 if "projectes" in d else VERSIO_FITXER
    return Jaciment(nom=str(d.get("nom") or nom_defecte), epsg=epsg, shift=_copia_shift(shift),
                    codi=d.get("codi"), versio=versio, campanyes=campanyes)


def amb_shift_resolt(j: Jaciment, a_disc: Sequence[str]) -> Jaciment:
    """Regla del ``[0, 0, 0]``: un shift zero explícit és **pendent** només si cap campanya té fitxers a disc
    (``a_disc`` = ``projectes_de(carpeta)`` buit); si en té, és un valor real i es respecta."""
    if j.shift is not None and [float(v) for v in j.shift] == [0.0, 0.0, 0.0] and not list(a_disc):
        return j.amb_shift(None)
    return j


def parse_jaciment_yaml(text: str, nom_defecte: str = "") -> Jaciment:
    return jaciment_de_dict(escalars(text), nom_defecte)


def nomenclatura_de(text: str) -> dict[str, Any]:
    """Bloc ``nomenclatura`` (``separador``, ``camp_data_format``) del text de ``jaciment.yaml``, o ``{}``."""
    v = escalars(text).get(CLAU_NOMENCLATURA)
    return dict(v) if isinstance(v, dict) else {}


def bloc_pontarch_de(text: str) -> dict[str, Any]:
    """Bloc ``pontarch`` (ara només ``plantilles``) del text de ``jaciment.yaml``, o ``{}``."""
    v = escalars(text).get(BLOC_PONTARCH)
    return dict(v) if isinstance(v, dict) else {}


def avis_migracio(existent: Jaciment | None) -> str | None:
    """Línia de consola si escriure la capçalera migra un fitxer de la versió 1 a la 2; None si no cal."""
    if existent is None or existent.versio >= VERSIO_FITXER:
        return None
    return (f"{FITXER_JACIMENT} versió {existent.versio} → {VERSIO_FITXER}: 'projectes' passa a 'campanyes' (només la "
            f"carpeta de cada una; any i dates es completen en vincular cada projecte)")


def _shift_f(shift: Sequence[Num] | None) -> list[float] | None:
    return None if shift is None else [float(v) for v in shift]


def mateix_marc(a: Jaciment | None, b: Jaciment | None) -> bool:
    """Mateix EPSG i mateix shift (dos pendents també són el mateix marc)."""
    if a is None or b is None:
        return False
    return a.epsg == b.epsg and _shift_f(a.shift) == _shift_f(b.shift)


def marc_compatible(existent: Jaciment | None, nou: Jaciment | None) -> bool:
    """El marc nou es pot escriure sense «Editar la identitat…»: mateix marc, o el shift de l'existent era pendent i
    el nou el fixa (mateix EPSG)."""
    if existent is None or nou is None:
        return False
    return mateix_marc(existent, nou) or (existent.shift is None and existent.epsg == nou.epsg)


# ── shift.txt (3D Survey Collection) ─────────────────────────────────────────


def text_shift_txt(j: Jaciment) -> str:
    if j.shift is None:
        raise ValueError(f"{FITXER_SHIFT}: el shift és pendent; no s'escriu fins que el fixe un vincle amb chunks.")
    return f"EPSG::{j.epsg} " + " ".join(_num(v) for v in j.shift) + "\n"


def parse_shift_txt(text: str) -> tuple[int, list[Num]]:
    m = re.match(r"\s*EPSG::?(\d{4,6})\s+(\S+)\s+(\S+)\s+(\S+)", text or "")
    if not m:
        raise ValueError(f"{FITXER_SHIFT}: s'espera 'EPSG::<epsg> X0 Y0 Z0'.")
    return int(m.group(1)), normalitzar_shift([float(m.group(i)) for i in (2, 3, 4)])


# ── Carpeta de jaciment ──────────────────────────────────────────────────────


def projectes_de(jaciment: str | Path) -> list[str]:
    """Subcarpetes de campanya trobades a disc: les que tenen ``3D/``, ``ORTO/`` o ``DEM/`` (``SUBCARPETES_PROJECTE``),
    pel nom i ordenades; [] si la carpeta no existeix. És la font «a disc» de ``vincle.campanyes_unides``."""
    j = Path(jaciment)
    if not j.is_dir():
        return []
    return sorted(d.name for d in j.iterdir()
                  if d.is_dir() and any((d / sub).is_dir() for sub in SUBCARPETES_PROJECTE))


def llegir_lliurament(carpeta: str | Path) -> Jaciment | None:
    """``jaciment.yaml`` de la carpeta (o ``shift.txt`` si només hi ha aquest); None si no n'hi ha."""
    carpeta = Path(carpeta)
    f = carpeta / FITXER_JACIMENT
    if f.is_file():
        return amb_shift_resolt(parse_jaciment_yaml(f.read_text(encoding="utf-8"), carpeta.name), projectes_de(carpeta))
    s = carpeta / FITXER_SHIFT
    if s.is_file():
        epsg, shift = parse_shift_txt(s.read_text(encoding="utf-8"))
        return Jaciment(nom=carpeta.name, epsg=epsg, shift=shift)
    return None


def desar_lliurament(carpeta: str | Path, j: Jaciment, campanya: Campanya | str | None = None,
                     forcar_marc: bool = False) -> tuple[list[Path], list[str]]:
    """Escriu la **capçalera** de ``jaciment.yaml`` (``reemplacar_capcalera``: ``oculat:``, ``pontarch:``,
    ``nomenclatura:`` i qualsevol altre bloc queden byte a byte; si el fitxer no existeix es crea en ordre canònic) i
    ``shift.txt`` a l'arrel de la carpeta de jaciment, **només si el contingut canvia**. ``campanya`` (``Campanya`` o
    nom de carpeta) s'afig o es completa a ``campanyes``. No crea cap subcarpeta (``lliurament.desar_lliurament`` hi
    afig ``<campanya>/3D/`` per a l'exportador de models).

    Si la carpeta ja té un ``jaciment.yaml`` amb un altre EPSG o shift, alça ``ValueError`` llevat que
    ``forcar_marc`` siga True (només després d'«Editar la identitat…» confirmat): EPSG i shift no es toquen mai
    d'una altra manera; fixar un shift **pendent** sí que es pot (``marc_compatible``). Amb el shift pendent no
    s'escriu ``shift.txt``. Retorna ``(rutes escrites de veritat, avisos)``; l'avís de migració v1 → v2 el mostra qui
    crida.
    """
    carpeta = Path(carpeta)
    carpeta.mkdir(parents=True, exist_ok=True)
    f_yaml = carpeta / FITXER_JACIMENT
    text_vell = f_yaml.read_text(encoding="utf-8") if f_yaml.is_file() else None
    existent = llegir_lliurament(carpeta)
    avisos: list[str] = []
    if existent is not None and not marc_compatible(existent, j):
        if not forcar_marc:
            raise ValueError(f"{FITXER_JACIMENT} de {carpeta} ja fixa EPSG:{existent.epsg} i shift "
                             f"{text_shift(existent.shift)}: no es canvien si no és amb «Editar la identitat…».")
        base: list[Campanya] = []
    else:
        base = existent.campanyes if existent is not None else []
    nou = Jaciment(j.nom, j.epsg, _copia_shift(j.shift), j.codi, VERSIO_FITXER, base + j.campanyes).amb_campanya(campanya)
    avis = avis_migracio(existent) if text_vell is not None else None
    if avis:
        avisos.append(avis)
    text_yaml = (reemplacar_capcalera(text_vell, text_capcalera(nou)) if text_vell is not None
                 else text_jaciment_yaml(nou))
    escrits = []
    fitxers = [(FITXER_JACIMENT, text_yaml)]
    if nou.shift is not None:                    # pendent: shift.txt espera el primer vincle amb chunks
        fitxers.append((FITXER_SHIFT, text_shift_txt(nou)))
    for nom, text in fitxers:
        p = carpeta / nom
        if p.is_file() and p.read_text(encoding="utf-8") == text:
            continue
        p.write_text(text, encoding="utf-8", newline="\n")
        escrits.append(p)
    return escrits, avisos


# ── doc.meta compartit: jaciment/carpeta i jaciment/campanya ─────────────────


def llegir_jaciment_meta(meta: Any) -> tuple[str | None, str | None, str]:
    """``(carpeta, campanya, font)`` del vincle del projecte: ``compartida`` si hi ha ``jaciment/carpeta``;
    ``heretada`` si només hi ha ``pontarch/lliurament`` (JSON de la 0.5–0.6; es migra a les compartides en desar el
    vincle i no s'escriu més); ``cap`` si res."""
    carpeta = _claus.llegir_compartida(meta, CLAU_CARPETA)
    if carpeta:
        campanya = _claus.llegir_compartida(meta, CLAU_CAMPANYA)
        return str(carpeta), (str(campanya) if campanya else None), "compartida"
    vell = parse_meta_lliurament(_claus.llegir_meta(meta, CLAU_META_LLIURAMENT))
    if vell:
        return vell[0], (vell[1] or None), "heretada"
    return None, None, "cap"


def escriure_jaciment_meta(meta: Any, carpeta: str | Path, campanya: str | None) -> list[str]:
    """Escriu les dues claus compartides (``jaciment/carpeta``, ``jaciment/campanya``). **No toca**
    ``pontarch/lliurament``: queda com a heretada fins que una eina la lleve explícitament. Retorna les claus."""
    return [_claus.escriure_compartida(meta, CLAU_CARPETA, str(carpeta)),
            _claus.escriure_compartida(meta, CLAU_CAMPANYA, str(campanya) if campanya else None)]


def text_font_vincle(font: str, carpeta: str | None = None, campanya: str | None = None) -> str:
    """Línia de consola sobre d'on ve el vincle del projecte."""
    if font == "compartida":
        return f"vincle del projecte (jaciment/carpeta): {carpeta} · campanya {campanya or '—'}"
    if font == "heretada":
        return (f"clau heretada pontarch/lliurament: {carpeta} · {campanya or '—'} (es migra a jaciment/carpeta en "
                "desar el vincle)")
    return "projecte sense vincle al jaciment (cap clau jaciment/carpeta ni pontarch/lliurament)"


def text_meta_lliurament(jaciment: str | Path, projecte: str) -> str:
    """Valor de ``doc.meta['pontarch/lliurament']`` (JSON amb la carpeta de jaciment i la subcarpeta). **Clau
    heretada des de la 0.7.0**: es llig (``llegir_jaciment_meta``) però ja no s'escriu; només per a tests."""
    import json
    return json.dumps({"jaciment": str(jaciment), "projecte": str(projecte)}, ensure_ascii=False)


def parse_meta_lliurament(text: Any) -> tuple[str, str] | None:
    """``(carpeta de jaciment, subcarpeta de projecte)`` del valor heretat, o None si no n'hi ha o no s'entén."""
    import json
    if not text:
        return None
    try:
        d = json.loads(str(text))
        jaciment, projecte = str(d.get("jaciment") or ""), str(d.get("projecte") or "")
    except (ValueError, TypeError, AttributeError):
        return None
    return (jaciment, projecte) if jaciment else None
