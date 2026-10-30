"""Exportació de dades de Metashape cap a la base de dades del jaciment.

Lògica pura (sense el mòdul ``Metashape``), testable amb pytest:

- ``codis``   gramàtica del codi d'estructura (config ``bd.yaml``)
- ``georef``  coordenades del marcador + orientació de façana → CSV
- ``shapes``  gramàtica d'etiquetes de shapes, mètriques i controls de qualitat
- ``vistes``  vistes del region (f/b/t/bo/r/l) i els seus eixos: contracte amb els ortomosaics de Pontarch
- ``csv_out`` escriptura dels CSV (UTF-8 amb BOM, per a Access)

Els diàlegs de Metashape viuen a ``metashape_addon/exportar_*_bd.py``.
"""
