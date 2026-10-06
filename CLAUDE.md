# tablero-atxlab

Tablero de proyectos de ATX Lab. HTML + JS plano, sin build ni dependencias: se abre `index.html` en el navegador.

- `index.html` — estructura, estilos y layout (incluye sidebar, tema claro/oscuro y barra de filtros)
- `app.js` — lógica del tablero: carga de datos, filtros, navegación entre proyectos
- `deck.js` — vista de presentación
- `plan.js` — datos del plan

## ⚠️ ANTES DE SUBIR CUALQUIER CAMBIO — LÉELO

**Este repo tiene dos personas trabajando y ya se perdió trabajo dos veces por la misma causa.**

El problema: subir archivos con **"Add files via upload"** desde la web de GitHub. Eso reemplaza el archivo completo con tu copia local. Si tu copia no tiene los cambios de la otra persona, los borras — y Git no avisa, porque técnicamente es un commit válido.

Ya pasó el **4 de septiembre** (commit `6de0923` = "restaurar sidebar, filtros y tema tras upload que los pisó") y otra vez el **23 de septiembre** (commits `c9e7cc9`..`9c9fec4`, que borraron el sidebar, el modo oscuro y el sistema de temas, y dejaron los filtros a medias).

### Reglas

1. **Nunca uses "Add files via upload" en github.com.** Ni para un archivo. Ni "es un cambio chiquito".
2. **Siempre `git pull` antes de empezar a trabajar**, y otra vez antes de subir.
3. **Nunca `git push --force` ni `--force-with-lease` sobre `main`.**
4. Si `git push` te rechaza, es que hay cambios que no tienes. **No lo fuerces**: haz `git pull --rebase`, resuelve, y vuelve a subir.

### El flujo correcto

```bash
git pull                      # antes de tocar nada
# ... tus cambios ...
git add -A
git commit -m "describe qué cambiaste"
git pull --rebase             # trae lo de la otra persona
git push
```

Si `git pull --rebase` marca conflictos, resuélvelos conservando **ambos** lados: los cambios de la otra persona y los tuyos. Si no sabes qué hacer, pregunta antes de descartar algo.

## Para Claude Code / agentes de IA

Si estás trabajando en este repo con un agente:

- **Ejecuta `git pull` antes de tu primera edición** de cualquier sesión. Si el repo está atrás del remoto, tráelo antes de escribir una sola línea.
- **Nunca ejecutes `git push --force`, `git push -f`, ni `git reset --hard` sobre trabajo no publicado** sin que la persona lo pida explícitamente.
- **Antes de hacer `push`, corre `git fetch && git log --oneline HEAD..origin/main`.** Si sale algo, hay cambios ajenos que no tienes: intégralos con `git pull --rebase` antes de subir.
- **Nunca reescribas un archivo completo** (`index.html`, `app.js`, `deck.js`, `plan.js`) desde cero. Edita solo las secciones que cambian. Reescribirlos enteros es exactamente lo que ha causado las dos pérdidas de trabajo.
- Si detectas que un archivo perdió funcionalidad que estaba en un commit anterior (sidebar, tema oscuro, filtros), **avísalo antes de commitear** en vez de asumir que fue intencional.

## Funcionalidad que hay que conservar

Estas features viven en `index.html` y `app.js` y ya se borraron accidentalmente una vez. Si vas a tocar esos archivos, verifica que sigan ahí antes de subir:

- **Sidebar** de navegación lateral
- **Tema claro / oscuro** con su toggle
- **Barra de filtros** del tablero
- **Navegación al abrir un proyecto** (se rompió una vez, arreglada en `d4af4c1`)

Comando rápido para verificar antes de commitear:

```bash
grep -c "sidebar\|theme\|oscuro" index.html app.js
```

Si el conteo baja mucho respecto a lo que había, algo se perdió.
