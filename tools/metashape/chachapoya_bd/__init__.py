"""Menú *Extensions → Base de dades* de Metashape per a la base de dades del jaciment.

Traslladat d'Oculat (``52863de``) sense redissenyar-lo: vegeu ``README.md`` d'esta carpeta.

- ``db/``             lògica pura (sense ``Metashape``), amb tests a ``tests/``
- ``configuracio.py`` opcions de ``<jaciment>/bd.yaml`` i carpeta d'eixida ``<jaciment>/BD``
- ``_vendor/``        contracte del jaciment i peces Qt, vendoritzats (``_vendor/ORIGEN.md``)
- ``exportar_georef_bd.py``, ``exportar_metriques_bd.py``, ``auditar_shapes.py``: les tres eines
- ``install_menu.py`` registra les tres entrades del menú
"""
