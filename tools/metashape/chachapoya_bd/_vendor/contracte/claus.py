"""
Claus de Pontarch als projectes de Metashape (lògica pura).

Pontarch deixa marques als fitxers ``.psx``: metadades (``camera.meta``, ``chunk.meta``, ``doc.meta``,
``shape.attributes``) i etiquetes de grups de shapes. Convenció:

- **Escriu** sempre ``pontarch/<nom>`` i grups ``pontarch_<base>_<sufix>``.
- **Llig** també les claus **heretades** d'Oculat (``oculat/<nom>``, ``oculat_<base>_<sufix>``), perquè els
  projectes ja marcats continuen funcionant; la clau pròpia té prioritat si hi ha les dues.
- **No migra** res: cap escriptura toca una clau heretada. L'única excepció és ``esborrar_meta``, que
  l'eina crida quan vol llevar explícitament una metadada (per exemple en restaurar una càmera): aleshores
  cauen les dues, perquè deixar-hi la heretada seria el mateix que no haver-la esborrat.
- Les eines imprimeixen al log quantes claus heretades han trobat (``comptar_heretades`` + ``text_heretades``).
- Namespace **compartit** ``jaciment/`` (``jaciment/carpeta``, ``jaciment/campanya``; contracte amb Oculat,
  ``docs/CONTRACTE_JACIMENT.md``): l'escriu qui vincula el projecte a la carpeta de jaciment, siga quin siga
  l'add-on. No és «pròpia» ni «heretada»: ``claus_meta`` no la inclou i ``comptar_heretades`` no la compta.

Els mòduls de ``metashape/`` només criden este mòdul (``metashape/claus_projecte.py`` arreplega les claus dels
objectes de Metashape i les passa ací). Les metadades de Metashape són cadenes: el valor es passa tal qual.
"""

from __future__ import annotations

from typing import Any, Iterable

PREFIX_META = "pontarch/"
PREFIX_GRUP = "pontarch_"
PREFIX_COMPARTIT = "jaciment/"
PREFIXOS_META_HERETATS = ("oculat/",)
PREFIXOS_GRUP_HERETATS = ("oculat_",)


# ── Noms ─────────────────────────────────────────────────────────────────────


def clau_meta(nom: str) -> str:
    """La clau que Pontarch escriu: ``pontarch/<nom>``."""
    return PREFIX_META + nom


def claus_meta(nom: str) -> list[str]:
    """Les claus que Pontarch llig per a ``nom``, en ordre de prioritat: la pròpia i després les heretades."""
    return [clau_meta(nom)] + [p + nom for p in PREFIXOS_META_HERETATS]


def nom_grup(base: str, sufix: str = "") -> str:
    """Etiqueta del grup de shapes que Pontarch crea: ``pontarch_<base>_<sufix>``."""
    return f"{PREFIX_GRUP}{base}_{sufix}"


def noms_grup(base: str, sufix: str = "") -> list[str]:
    """Etiquetes de grup que Pontarch reconeix per a ``base``/``sufix``: la pròpia i les heretades."""
    return [nom_grup(base, sufix)] + [f"{p}{base}_{sufix}" for p in PREFIXOS_GRUP_HERETATS]


def clau_compartida(nom: str) -> str:
    """La clau compartida amb Oculat: ``jaciment/<nom>``."""
    return PREFIX_COMPARTIT + nom


def es_meta_compartida(clau: Any) -> bool:
    return str(clau).startswith(PREFIX_COMPARTIT)


def es_meta_heretada(clau: Any) -> bool:
    return str(clau).startswith(PREFIXOS_META_HERETATS)


def es_grup_heretat(etiqueta: Any) -> bool:
    return str(etiqueta).startswith(PREFIXOS_GRUP_HERETATS)


def es_meta_pontarch(clau: Any) -> bool:
    """Clau pròpia o heretada (les de Metashape, com ``Image/Quality``, no ho són)."""
    return str(clau).startswith((PREFIX_META,) + PREFIXOS_META_HERETATS)


def nom_de_clau(clau: Any) -> str | None:
    """``<nom>`` d'una clau pròpia o heretada; None si no és cap de les dues."""
    c = str(clau)
    for p in (PREFIX_META,) + PREFIXOS_META_HERETATS:
        if c.startswith(p):
            return c[len(p):]
    return None


def noms_amb_prefix(claus: Iterable[Any], prefix_nom: str) -> set[str]:
    """Noms (sense cap prefix) de les claus pròpies o heretades que comencen per ``prefix_nom``
    (p. ex. ``"estat/"`` → ``{"textura", "orto"}`` de ``pontarch/estat/textura`` i ``oculat/estat/orto``)."""
    out: set[str] = set()
    for c in claus:
        nom = nom_de_clau(c)
        if nom is not None and nom.startswith(prefix_nom) and len(nom) > len(prefix_nom):
            out.add(nom[len(prefix_nom):])
    return out


# ── Lectura i escriptura ─────────────────────────────────────────────────────


def _valor(meta: Any, clau: str) -> Any:
    """``meta[clau]`` o None (``MetaData`` i ``attributes`` de Metashape alcen excepcions diverses)."""
    try:
        return meta[clau]
    except Exception:
        return None


def llegir_meta(meta: Any, nom: str, defecte: Any = None) -> Any:
    """Valor de ``pontarch/<nom>`` o, si no hi és, de les claus heretades; ``defecte`` si cap."""
    for clau in claus_meta(nom):
        v = _valor(meta, clau)
        if v is not None:
            return v
    return defecte


def llegir_compartida(meta: Any, nom: str, defecte: Any = None) -> Any:
    """Valor de ``jaciment/<nom>`` (cadena buida compta com a absent); ``defecte`` si no hi és."""
    v = _valor(meta, clau_compartida(nom))
    return defecte if v is None or str(v) == "" else v


def escriure_compartida(meta: Any, nom: str, valor: Any) -> str:
    """Escriu ``jaciment/<nom>``; None l'esborra. Retorna la clau."""
    clau = clau_compartida(nom)
    if valor is None:
        try:
            del meta[clau]
        except Exception:
            try:
                meta[clau] = ""
            except Exception:
                pass
        return clau
    meta[clau] = str(valor)
    return clau


def te_meta(meta: Any, nom: str) -> bool:
    return llegir_meta(meta, nom) is not None


def escriure_meta(meta: Any, nom: str, valor: Any) -> str:
    """Escriu ``pontarch/<nom>`` (mai la heretada). ``valor`` None equival a ``esborrar_meta``. Retorna la clau."""
    clau = clau_meta(nom)
    if valor is None:
        esborrar_meta(meta, nom)
        return clau
    meta[clau] = valor
    return clau


def esborrar_meta(meta: Any, nom: str) -> int:
    """Lleva ``pontarch/<nom>`` i les seues heretades. Retorna quantes claus hi havia."""
    n = 0
    for clau in claus_meta(nom):
        if _valor(meta, clau) is None:
            continue
        try:
            del meta[clau]
            n += 1
        except Exception:
            try:
                meta[clau] = ""
                n += 1
            except Exception:
                pass
    return n


def triar_grup(etiquetes: Iterable[str], base: str, sufix: str = "") -> str | None:
    """Entre les etiquetes de grup existents, la que correspon a ``base``/``sufix`` (la pròpia abans que la
    heretada), o None si no n'hi ha cap."""
    existents = set(str(e) for e in etiquetes)
    for nom in noms_grup(base, sufix):
        if nom in existents:
            return nom
    return None


# ── Recompte per al log ──────────────────────────────────────────────────────


def comptar_heretades(claus: Iterable[Any] = (), grups: Iterable[Any] = ()) -> tuple[int, int]:
    """``(metadades heretades, grups heretats)`` entre les claus i etiquetes donades."""
    n_meta = sum(1 for c in claus if es_meta_heretada(c))
    n_grups = sum(1 for g in grups if es_grup_heretat(g))
    return n_meta, n_grups


def text_heretades(n_meta: int, n_grups: int, context: str = "") -> str | None:
    """Línia de log si hi ha claus heretades; None si no n'hi ha cap."""
    if not n_meta and not n_grups:
        return None
    parts = []
    if n_meta:
        parts.append(f"{n_meta} metadad{'a' if n_meta == 1 else 'es'} oculat/…")
    if n_grups:
        parts.append(f"{n_grups} grup{'' if n_grups == 1 else 's'} de shapes oculat_…")
    on = f" a {context}" if context else ""
    return f"claus heretades d'Oculat{on}: " + " i ".join(parts) + " (es lligen com a pròpies; no es migren)"
