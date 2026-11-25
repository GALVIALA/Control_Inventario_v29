# V29: migración y validación

## Estrategia IndexedDB V2 → V3

- La base conserva el nombre `vestas-inventario-offline` y sube de versión 2 a 3.
- La actualización crea `inventories`, `settings` y `backups`. No elimina ni vacía `sessions`.
- El reinicio único histórico conserva el umbral 2. Por lo tanto, una base que ya está en V2 no vuelve a limpiarse al pasar a V3.
- Después de abrir V3, una transacción atómica copia las sesiones V2 a inventarios V29 y registra `legacy-v2-to-v3`. Si falla, no queda una migración parcial y `sessions` permanece como recuperación.
- La sesión más reciente queda activa. Las restantes se conservan abiertas pero inactivas, porque V28 no registraba un cierre definitivo; pueden reactivarse explícitamente sin falsificar una fecha de cierre.
- Cada conexión instala `onversionchange` y se cierra para permitir futuras actualizaciones. Una apertura bloqueada pide cerrar otras pestañas.

## Origen del almacenamiento

IndexedDB pertenece al origen completo: esquema, host y puerto. `http://127.0.0.1:8028` y `http://127.0.0.1:8029` usan bases distintas. Una prueba local de V29 en 8029 no puede ver ni migrar los datos reales guardados por V28 en 8028. Para una migración real, V29 debe abrirse desde el mismo origen usado por V28.

## Recuperación

- Antes de actualizar desde Excel o cerrar definitivamente se guarda un snapshot en `backups` dentro de la misma transacción que modifica el inventario.
- Administración ofrece restaurar el respaldo más reciente del inventario activo y crea un respaldo adicional antes de restaurarlo.
- Cerrar no elimina: cambia `status` a `closed`, registra `closedAt` y quita el inventario activo.

## Prueba manual en Safari

1. Abrir V29 en su dirección local y confirmar el encabezado V29.
2. Crear `Prueba A` con una copia desechable de la plantilla.
3. Cambiar una cantidad, comentario y serie; recargar y continuar.
4. Crear `Prueba B`; confirmar que `Prueba A` permanece intacta y puede reactivarse.
5. Exportar `Prueba B`, editar COUNT QTY, borrar un comentario y cambiar una marca verde en Excel.
6. Usar **Actualizar desde Excel**, revisar el resumen y confirmar.
7. Exportar y abrir el resultado en Excel; verificar los 21/5 encabezados y fórmulas R:S:T:U.
8. Entrar en **Finalizar**, generar respaldo y realizar la segunda confirmación de cierre.
9. Consultar el inventario cerrado; comprobar que cantidades, comentarios y series no se pueden modificar.
10. Volver al inventario activo y confirmar que conserva búsqueda, datos y progreso.

No usar datos reales ni el origen de V28 durante estas pruebas.
