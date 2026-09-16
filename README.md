# Simulador del problema 5.5

## Origen y descripción del problema

Este proyecto resuelve el problema 5.5 del libro *Simulación: un enfoque práctico*, de Raúl Coss Bu.

Se estudia la gestión de inventario de un artículo con las siguientes condiciones:

- La demanda diaria sigue una distribución Binomial con `n = 6` y `p = 1/2`.
- El tiempo de entrega de un pedido, en días, sigue una distribución Poisson con `λ = 3`.
- Mantener una unidad en inventario cuesta `$1` por día.
- Cada unidad faltante cuesta `$10`.
- Emitir una orden cuesta `$50`, independientemente de la cantidad solicitada.

El objetivo es comparar cuál de estas políticas resulta más económica:

1. Revisión diaria: ordenar para llevar el inventario neto a 8 unidades.
2. Punto de pedido: si el inventario neto es menor o igual a 10, ordenar hasta 30.

En la simulación, el inventario negativo representa unidades pendientes de surtir. Cada jornada recibe pedidos programados, atiende la demanda, calcula costos y después emite las nuevas órdenes.

## Programa Go

Usando el Go indicado en el enunciado del trabajo:

```bash
go run ./cmd/simulador -dias 285 -uniformes ri.csv
```

`ri.csv` o `ri.txt` debe contener los valores externos `R_i` entre 0 y 1, separados por coma, punto y coma, espacios o saltos de línea. El programa Go no genera pseudoaleatorios.

## Interfaz web (TypeScript/HTML)

```bash
cd web
pnpm install
pnpm dev
```


Para abrirla con doble clic sin iniciar un servidor, ejecuta `pnpm build` y abre `web/index.html`. El proceso genera `web/standalone/simulador.js`, que funciona también con `file://`.

El Excel puede tener una columna `R_i` con los uniformes. Si sus valores son fórmulas no calculadas, la página puede reconstruir de forma determinística los generadores lineal mixto (`X0`, `a`, `c`, `m`, `N`) y multiplicativo (`X0`, `a`, `m`, `N`) usando solo parámetros incluidos en el propio archivo. Para otros métodos, exporta los valores calculados de `R_i`. No crea una nueva secuencia ni usa una semilla interna.
