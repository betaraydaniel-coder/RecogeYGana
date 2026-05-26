# Recoge y Gana — Minijuego 3D en Unity

Minijuego 3D desarrollado en **Unity 2022 LTS** y **C#** en el que el jugador
controla una esfera azul que debe recolectar 10 gemas para ganar. Incluye
física, cámara con seguimiento suave, UI animada, audio con fades y
herramientas de editor para generar la escena, los materiales, el prefab
del coleccionable y el ejecutable.

> **El proyecto está 100 % auto-construido.** No necesitas crear nada a
> mano en Unity: al abrir el proyecto por primera vez, el script
> `SceneSetup.cs` genera la escena `Main.unity`, el prefab
> `Collectible.prefab`, los 5 materiales y los 10 coleccionables.

---

## 1. Cómo abrir el proyecto en Unity Hub

1. Instala **Unity Hub** y dentro de él la versión **2022.3.20f1 (LTS)** con
   el módulo **Windows Build Support (IL2CPP)**.
2. Descarga este repositorio (clónalo o descomprime el ZIP).
3. En Unity Hub → pestaña *Projects* → **Add → Add project from disk** y
   selecciona la carpeta raíz del proyecto.
4. Ábrelo. La primera vez Unity importará todos los assets (puede tardar
   1–3 min). En cuanto termine, el script `SceneSetup.cs` se ejecuta
   automáticamente y deja:
   * 5 materiales en `Assets/Materials/`.
   * `Assets/Prefabs/Collectible.prefab`.
   * `Assets/Scenes/Main.unity` con la escena completa (Ground, 4 paredes,
     6 decoraciones, Directional Light, Player, Main Camera, GameManager,
     AudioManager, Canvas, EventSystem y 10 coleccionables).
   * La escena registrada en *Build Settings* como índice 0.
5. Pulsa **Play** y juega.

### Menús personalizados (barra superior "Recoge y Gana")

* **Construir Escena Completa** — recrea la escena desde cero (por si la
  borraste o quieres regenerarla).
* **Crear Materiales** — regenera los 5 materiales.
* **Spawnear Coleccionables (x10)** — añade 10 coleccionables aleatorios
  a la escena abierta.
* **Build Windows64** — compila el `.exe` localmente en
  `../Build/RecogeYGana/RecogeYGana.exe`.

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
| `MaterialSetup.cs` | `Assets/Scripts/Editor/` | `[MenuItem]` que crea los 5 materiales con sus colores HEX usando `AssetDatabase`. |
| `CollectibleSpawner.cs` | `Assets/Scripts/Editor/` | `[MenuItem]` que instancia 10 coleccionables en x/z ∈ [-11, 11] excluyendo un radio de 2u en el origen. |
| `SceneSetup.cs` | `Assets/Scripts/Editor/` | Genera la escena completa (Ground, paredes, decoraciones, luz, jugador, cámara, GameManager, AudioManager, Canvas, EventSystem y 10 coleccionables) la primera vez que se abre el proyecto. |
| `BuildScript.cs` | `Assets/Scripts/Editor/` | `[MenuItem]` que compila a `StandaloneWindows64`. Incluye un punto de entrada `BuildGameCI()` usado por el workflow de GitHub Actions. |

---

## 3. Cómo agregar audios desde freesound.org

1. Entra a <https://freesound.org/> y crea una cuenta gratuita.
2. En el filtro de búsqueda activa **License → Creative Commons 0**.
3. Sugerencias de búsqueda:
   * **Música de fondo**: `ambient loop`, `chiptune loop`, `calm background music`.
   * **Recoger objeto**: `pickup coin`, `collect item`, `ding short`.
4. Descarga los `.wav` o `.ogg` y guárdalos en `Assets/Audio/`.
5. Renómbralos por claridad:
   * `bgm_loop.ogg`
   * `sfx_collect.wav`
6. Selecciona el GameObject **AudioManager** en la escena y arrastra los
   clips a los campos `Background Music` y `Collect Sfx`.
7. Pulsa Play. La música hará *fade-in* hasta 0.4 y al ganar hará
   *fade-out* hasta 0.

---

## 4. Cómo generar el ejecutable

### Opción A — En tu PC con Unity (rápido si tienes Unity instalado)

1. Asegúrate de que `Main.unity` esté en *Build Settings* como índice 0
   (lo hace automáticamente `SceneSetup`).
2. Menú: **Recoge y Gana → Build Windows64**.
3. El `.exe` se genera en `../Build/RecogeYGana/RecogeYGana.exe`.

### Opción B — Build en la nube con GitHub Actions (PC sin Unity)

Si tu PC no puede correr Unity (i3, poca RAM, sin GPU), usa la build de la
nube. **Tú no abres Unity en ningún momento**: GitHub compila el `.exe` y
te lo deja como artefacto descargable.

#### Configuración inicial (una sola vez, ~5 minutos)

1. Crea una cuenta gratuita en Unity:
   <https://id.unity.com/account/new>.
2. En tu repo de GitHub, ve a **Actions → "Request Unity License" → Run workflow**.
   Espera ~1 minuto y descarga el artefacto `Unity_v2022.x.alf`.
3. Sube ese `.alf` a <https://license.unity3d.com/manual>, elige
   **Unity Personal**, acepta los términos y descarga el archivo `.ulf`
   que te devuelve la web.
4. En tu repo: **Settings → Secrets and variables → Actions → New
   repository secret** y crea **tres** secretos:

   | Nombre | Valor |
   | --- | --- |
   | `UNITY_LICENSE` | El contenido COMPLETO del archivo `.ulf` (ábrelo con bloc de notas y pega todo) |
   | `UNITY_EMAIL` | Tu email de Unity |
   | `UNITY_PASSWORD` | Tu contraseña de Unity |

#### Compilar el .exe

1. Cada `git push` dispara el workflow **"Build Windows64"** automáticamente.
   También puedes lanzarlo manualmente en **Actions → Build Windows64 → Run workflow**.
2. Espera ~8–12 minutos (la primera vez tarda más; las siguientes son
   más rápidas por el caché).
3. Cuando el workflow termine en verde, abre la corrida y descarga el
   artefacto **`RecogeYGana-Windows64`**.
4. Descomprime el ZIP descargado: dentro encontrarás `RecogeYGana.exe`
   junto a `RecogeYGana_Data/`, `UnityPlayer.dll`, `MonoBleedingEdge/`,
   etc. Esa carpeta es exactamente lo que tienes que entregar.

---

## 5. Qué incluir / excluir en el ZIP de entrega

### ✅ Incluir

* `Assets/` (Scripts, Materials, Audio, Scenes, Prefabs).
* `Packages/` (manifest de paquetes).
* `ProjectSettings/` (configuración del proyecto).
* `README.md`.
* La carpeta de build descargada del artefacto de GitHub Actions
  (`RecogeYGana.exe` + `RecogeYGana_Data/` + `UnityPlayer.dll` + …).

Estructura sugerida del ZIP final:

```
Entrega_RecogeYGana.zip
├── RecogeYGana/                  ← carpeta del proyecto Unity
│   ├── Assets/
│   ├── Packages/
│   ├── ProjectSettings/
│   └── README.md
└── Build/
    └── RecogeYGana/
        ├── RecogeYGana.exe
        ├── RecogeYGana_Data/
        ├── UnityPlayer.dll
        └── MonoBleedingEdge/
```

### ❌ Excluir

* `Library/` (cache de Unity, se regenera al abrir).
* `Temp/`, `Obj/`, `Logs/`, `UserSettings/`.
* `.vs/`, `.idea/`, `*.csproj`, `*.sln` (se regeneran).
* `.github/` no es obligatorio incluirla, pero no estorba.
* Cualquier audio que NO sea CC0 / con licencia restrictiva.

> El `.gitignore` incluido ya excluye todo lo anterior, así que si haces
> el ZIP a partir de un `git archive` o de la carpeta limpia obtendrás
> exactamente lo que se debe entregar.

---

## Checklist de verificación final

* [x] **Paso 1** — Estructura de carpetas `Assets/{Scripts,Materials,Audio,Scenes,Prefabs}` creada.
* [x] **Paso 2** — `MaterialSetup.cs` con `[MenuItem]` genera los 5 materiales (`MatGround #2D6A4F`, `MatWall #4A4E69`, `MatDeco_A #C1666B`, `MatDeco_B #E8C547`, `MatPlayer #4361EE`) y los guarda en `Assets/Materials/` con `AssetDatabase`.
* [x] **Paso 3** — `PlayerController.cs` (Sphere + Rigidbody + SphereCollider, tag *Player*, `AddForce` en `FixedUpdate`, `moveSpeed=12`, `maxSpeed=8`, `Input.GetAxis`) y `CameraFollow.cs` (`Vector3.Lerp`, `smoothSpeed=0.125`, offset por defecto `(0,10,-8)` y editable en Inspector). Comentarios en español ✔.
* [x] **Paso 4** — `Collectible.cs` (rotación doble eje Y:90°/s X:45°/s, *bob* `Mathf.Sin` amp=0.3 freq=2, trigger que verifica tag *Player*, llama `GameManager.Instance.CollectItem()` y `AudioManager.Instance.PlayCollect()` y se destruye). `CollectibleSpawner.cs` con `[MenuItem]` que instancia 10 coleccionables en x/z ∈ [-11, 11] excluyendo radio 2u del origen.
* [x] **Paso 5** — `GameManager.cs` Singleton con `CollectItem()`, condición de victoria y llamadas a `UIManager.ShowWinScreen()` + `AudioManager` fade-out. `UIManager.cs` con `UpdateScore` (`✦ {c} / {t}`), `ShowWinScreen()` con Coroutine fade+scale 0.8→1.0 en 0.6s y botón *Jugar de nuevo* → `SceneManager.LoadScene(0)`. Jerarquía Canvas (`ScorePanel` + `WinPanel` inactivo) construida automáticamente por `SceneSetup`.
* [x] **Paso 6** — `AudioManager.cs` Singleton + `DontDestroyOnLoad`, dos `AudioSource` (música loop vol 0.4 con fade-in 0→0.4 en 2s y fade-out al ganar →0 en 1s, SFX vol 0.8), `PlayCollect()` público y comentario con instrucciones para descargar audios CC0 de freesound.org.
* [x] **Paso 7** — `BuildScript.cs` con `[MenuItem]` que ejecuta `BuildPipeline.BuildPlayer` → `StandaloneWindows64` → `../Build/RecogeYGana/RecogeYGana.exe`, y este `README.md` con las 5 secciones requeridas.
* [x] **Extras** — `SceneSetup.cs` que auto-genera escena y prefab al abrir; `Packages/manifest.json` y `ProjectSettings/ProjectVersion.txt` para que el proyecto esté listo "out of the box"; workflows de GitHub Actions (`build.yml` y `request-license.yml`) para compilar el `.exe` en la nube sin necesidad de instalar Unity.

¡Listo para jugar y entregar!
