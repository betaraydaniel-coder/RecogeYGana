# Recoge y Gana — Minijuego 3D en Unity

Minijuego 3D desarrollado en **Unity 2022 LTS (o superior)** y **C#** en el que el
jugador controla una esfera azul que debe recolectar 10 gemas para ganar.
Incluye física, cámara con seguimiento suave, UI animada, audio con fades y
herramientas de editor para generar materiales, spawnear coleccionables y
construir el ejecutable.

---

## 1. Cómo abrir el proyecto en Unity Hub

1. Instala **Unity Hub** y dentro de él una versión **2022 LTS (o superior)**
   con el módulo **Windows Build Support (IL2CPP)**.
2. Descarga este repositorio (clónalo o descomprime el ZIP).
3. En Unity Hub → pestaña *Projects* → **Add → Add project from disk** y
   selecciona la carpeta raíz `RecogeYGana/`.
4. Ábrelo y espera a que termine la importación.
5. En el menú superior aparecerá una entrada llamada **Recoge y Gana** con
   las acciones del editor incluidas.

### Configuración inicial recomendada en la escena

> Estos pasos solo son necesarios la primera vez (la escena no se incluye
> serializada para evitar conflictos entre versiones de Unity).

1. Abre / crea `Assets/Scenes/Main.unity`.
2. Ejecuta `Recoge y Gana → Crear Materiales` (genera los 5 materiales).
3. Crea los GameObjects de la escena:
   * **Ground** — `3D Object → Plane`, escala `(2.5, 1, 2.5)`,
     material `MatGround`.
   * **Walls** — 4 cubos largos rodeando el plano, material `MatWall`.
   * **Decorations** — 6 cubos/cilindros pequeños, materiales
     `MatDeco_A` / `MatDeco_B`.
   * **Directional Light** — viene por defecto al crear la escena.
   * **Player** — `3D Object → Sphere`, tag `Player`, agrega `Rigidbody`
     (Freeze Rotation X/Z opcional) y el script `PlayerController`,
     material `MatPlayer`.
   * **Main Camera** — agrega el script `CameraFollow` y arrastra el
     Player al campo `target`.
   * **GameManager (vacío)** — agrega el script `GameManager`.
   * **AudioManager (vacío)** — agrega el script `AudioManager` y
     asigna los clips (ver sección 3).
   * **Canvas + EventSystem** — crea `ScorePanel` (esquina sup. izq.) y
     `WinPanel` (centrado, desactivado). Agrega el script `UIManager`
     al Canvas y asigna sus referencias.
4. Crea un Prefab llamado `Collectible` en `Assets/Prefabs/` con el
   script `Collectible` y un Collider con *Is Trigger* activado.
5. Ejecuta `Recoge y Gana → Spawnear Coleccionables (x10)`.
6. Abre **File → Build Settings** y añade `Main.unity` como escena 0.
7. ¡Pulsa Play!

---

## 2. Lista de scripts y funciones

| Script | Ruta | Resumen |
| --- | --- | --- |
| `PlayerController.cs` | `Assets/Scripts/` | Movimiento físico con `AddForce` en `FixedUpdate` (`moveSpeed=12`, `maxSpeed=8`). |
| `CameraFollow.cs` | `Assets/Scripts/` | Cámara que sigue al jugador con `Vector3.Lerp` (`smoothSpeed=0.125`, `offset=(0,10,-8)`). |
| `Collectible.cs` | `Assets/Scripts/` | Rotación doble (X:45°/s, Y:90°/s), *bob* con `Mathf.Sin` (amp=0.3, freq=2), trigger que notifica a `GameManager` y `AudioManager` y se autodestruye. |
| `GameManager.cs` | `Assets/Scripts/` | Singleton. Cuenta gemas recogidas, notifica UI, detecta victoria y pide *fade-out* a `AudioManager`. |
| `UIManager.cs` | `Assets/Scripts/` | Singleton. Muestra `✦ {c} / {t}` y anima el `WinPanel` (fade + scale 0.8→1.0 en 0.6s). |
| `AudioManager.cs` | `Assets/Scripts/` | Singleton + `DontDestroyOnLoad`. Música en loop con *fade-in* 2s, *fade-out* 1s al ganar y SFX `PlayCollect()`. |
| `MaterialSetup.cs` | `Assets/Scripts/Editor/` | `[MenuItem]` que crea los 5 materiales con sus colores HEX. |
| `CollectibleSpawner.cs` | `Assets/Scripts/Editor/` | `[MenuItem]` que instancia 10 coleccionables en x/z ∈ [-11, 11] excluyendo un radio de 2u en el origen. |
| `BuildScript.cs` | `Assets/Scripts/Editor/` | `[MenuItem]` que ejecuta `BuildPipeline.BuildPlayer` para `StandaloneWindows64`. |

### Atajos del menú "Recoge y Gana"

* **Crear Materiales** → genera los 5 materiales en `Assets/Materials/`.
* **Spawnear Coleccionables (x10)** → coloca 10 prefabs aleatorios.
* **Build Windows64** → genera `../Build/RecogeYGana/RecogeYGana.exe`.

---

## 3. Cómo agregar audios desde freesound.org

1. Entra a <https://freesound.org/> y crea una cuenta gratuita.
2. En el filtro de búsqueda activa **License → Creative Commons 0**
   (uso libre, sin atribución obligatoria).
3. Sugerencias de búsqueda:
   * **Música de fondo**: `ambient loop`, `chiptune loop`,
     `calm background music`.
   * **Recoger objeto**: `pickup coin`, `collect item`, `ding short`.
4. Descarga los `.wav` o `.ogg` y guárdalos en `Assets/Audio/`.
5. Renómbralos para mayor claridad, por ejemplo:
   * `bgm_loop.ogg`
   * `sfx_collect.wav`
6. Selecciona el GameObject **AudioManager** en la escena y arrastra
   los clips a los campos:
   * `Background Music` → `bgm_loop`
   * `Collect Sfx` → `sfx_collect`
7. Pulsa Play. La música hará *fade-in* hasta 0.4 y al ganar hará
   *fade-out* hasta 0.

> **Atribución (opcional)** — aunque CC0 no la exige, es buena práctica
> listar los autores y URLs en `Assets/Audio/CREDITS.txt`.

---

## 4. Cómo hacer el Build (Standalone Windows 64)

### Opción A — Desde el menú personalizado

1. Asegúrate de tener `Main.unity` como escena 0 en **Build Settings**.
2. Menú: **Recoge y Gana → Build Windows64**.
3. El ejecutable se genera en `../Build/RecogeYGana/RecogeYGana.exe`
   (junto a la carpeta raíz del proyecto, NO dentro de ella).
4. Comprueba en *Console* el mensaje `✔ Build OK`.

### Opción B — Manual

1. **File → Build Settings…**
2. *Platform*: **Windows, Mac, Linux** → **Target Platform: Windows** →
   **Architecture: x86_64**.
3. Pulsa **Switch Platform** si es necesario.
4. **Build** → selecciona `Build/RecogeYGana/` y nombre `RecogeYGana.exe`.

---

## 5. Qué incluir / excluir en el ZIP de entrega

### ✅ Incluir

* Carpeta `Assets/` completa (Scripts, Materials, Audio, Scenes, Prefabs).
* Carpeta `Packages/` (manifest de paquetes).
* Carpeta `ProjectSettings/`.
* `README.md`.
* Carpeta `Build/RecogeYGana/` con el ejecutable y sus archivos `_Data`,
  `UnityPlayer.dll`, `MonoBleedingEdge/`, etc. (zip aparte si es muy grande).

### ❌ Excluir

* `Library/` (cache de Unity, se regenera al abrir).
* `Temp/` y `Obj/` (artefactos temporales de compilación).
* `Logs/` (logs del editor).
* `UserSettings/` (preferencias locales del editor).
* `.vs/`, `.idea/`, `*.csproj`, `*.sln` (se regeneran).
* Cualquier audio que **no** sea CC0 / con licencia restrictiva.

> Tip: añade un `.gitignore` estándar de Unity (ya disponible en
> <https://github.com/github/gitignore/blob/main/Unity.gitignore>) y haz
> el ZIP a partir del repositorio limpio.

---

## Checklist de verificación final

* [x] **Paso 1** — Estructura de carpetas `Assets/{Scripts,Materials,Audio,Scenes,Prefabs}` creada.
* [x] **Paso 2** — `MaterialSetup.cs` con `[MenuItem]` que genera los 5 materiales (`MatGround #2D6A4F`, `MatWall #4A4E69`, `MatDeco_A #C1666B`, `MatDeco_B #E8C547`, `MatPlayer #4361EE`) y los guarda en `Assets/Materials/` con `AssetDatabase`.
* [x] **Paso 3** — `PlayerController.cs` (Sphere + Rigidbody + SphereCollider, tag *Player*, `AddForce` en `FixedUpdate`, `moveSpeed=12`, `maxSpeed=8`, `Input.GetAxis`) y `CameraFollow.cs` (`Vector3.Lerp`, `smoothSpeed=0.125`, offset por defecto `(0,10,-8)` y editable en Inspector). Comentarios en español ✔.
* [x] **Paso 4** — `Collectible.cs` (rotación doble eje Y:90°/s X:45°/s, *bob* `Mathf.Sin` amp=0.3 freq=2, trigger que verifica tag *Player*, llama `GameManager.Instance.CollectItem()` y `AudioManager.Instance.PlayCollect()` y se destruye). `CollectibleSpawner.cs` con `[MenuItem]` que instancia 10 coleccionables en x/z ∈ [-11, 11] excluyendo radio 2u del origen.
* [x] **Paso 5** — `GameManager.cs` Singleton con `CollectItem()`, condición de victoria y llamadas a `UIManager.ShowWinScreen()` + `AudioManager` fade-out. `UIManager.cs` con `UpdateScore` (`✦ {c} / {t}`), `ShowWinScreen()` con Coroutine fade+scale 0.8→1.0 en 0.6s y botón *Jugar de nuevo* → `SceneManager.LoadScene(0)`. Jerarquía Canvas (`ScorePanel` + `WinPanel` inactivo) documentada.
* [x] **Paso 6** — `AudioManager.cs` Singleton + `DontDestroyOnLoad`, dos `AudioSource` (música loop vol 0.4 con fade-in 0→0.4 en 2s y fade-out al ganar →0 en 1s, SFX vol 0.8), `PlayCollect()` público y comentario con instrucciones para descargar audios CC0 de freesound.org.
* [x] **Paso 7** — `BuildScript.cs` con `[MenuItem]` que ejecuta `BuildPipeline.BuildPlayer` → `StandaloneWindows64` → `../Build/RecogeYGana/RecogeYGana.exe`, y este `README.md` con las 5 secciones requeridas.

¡Listo para jugar y entregar! 🎮
