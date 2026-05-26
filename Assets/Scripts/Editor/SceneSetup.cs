// SceneSetup.cs
// ─────────────────────────────────────────────────────────────────────────────
// Genera programáticamente la escena Main.unity, el prefab Collectible.prefab
// y los materiales — todo sin que el usuario tenga que hacer clic en nada.
//
// Se ejecuta automáticamente:
//   • La PRIMERA vez que Unity abre el proyecto (vía [InitializeOnLoadMethod]),
//     siempre que la escena Main.unity todavía no exista.
//   • Manualmente desde el menú "Recoge y Gana → Construir Escena Completa".
//
// También se usa en la build de la nube (GitHub Actions / game-ci): cuando
// Unity arranca en modo batch, este script crea la escena antes de que
// el builder lea EditorBuildSettings.
// ─────────────────────────────────────────────────────────────────────────────

using System.IO;
using UnityEditor;
using UnityEditor.SceneManagement;
using UnityEngine;
using UnityEngine.EventSystems;
using UnityEngine.SceneManagement;
using UnityEngine.UI;

public static class SceneSetup
{
    private const string ScenesFolder   = "Assets/Scenes";
    private const string ScenePath      = "Assets/Scenes/Main.unity";
    private const string PrefabsFolder  = "Assets/Prefabs";
    private const string PrefabPath     = "Assets/Prefabs/Collectible.prefab";

    // ─── Auto-ejecución al abrir el proyecto ─────────────────────────────────
    [InitializeOnLoadMethod]
    private static void AutoCreateOnLoad()
    {
        // Se retrasa un frame para asegurar que AssetDatabase esté listo.
        EditorApplication.delayCall += () =>
        {
            if (Application.isBatchMode || !File.Exists(ScenePath))
            {
                if (!File.Exists(ScenePath))
                {
                    BuildEverything();
                }
                else
                {
                    EnsureSceneInBuildSettings();
                }
            }
        };
    }

    [MenuItem("Recoge y Gana/Construir Escena Completa")]
    public static void BuildEverything()
    {
        Debug.Log("[SceneSetup] Construyendo escena completa...");

        // 1. Materiales.
        MaterialSetup.CrearMateriales();

        // 2. Prefab del coleccionable.
        CreateCollectiblePrefab();

        // 3. Escena vacía.
        Directory.CreateDirectory(ScenesFolder);
        var scene = EditorSceneManager.NewScene(NewSceneSetup.EmptyScene, NewSceneSetup.Single);

        // 4. Cargar materiales (ya creados en el paso 1).
        Material matGround = Load<Material>("Assets/Materials/MatGround.mat");
        Material matWall   = Load<Material>("Assets/Materials/MatWall.mat");
        Material matDecoA  = Load<Material>("Assets/Materials/MatDeco_A.mat");
        Material matDecoB  = Load<Material>("Assets/Materials/MatDeco_B.mat");
        Material matPlayer = Load<Material>("Assets/Materials/MatPlayer.mat");

        // 5. Ground (Plane escalado).
        var ground = GameObject.CreatePrimitive(PrimitiveType.Plane);
        ground.name = "Ground";
        ground.transform.localScale = new Vector3(2.5f, 1f, 2.5f);
        AssignMaterial(ground, matGround);

        // 6. Cuatro paredes.
        CreateWall("WallNorth", new Vector3(0f, 1f, 12.5f),  new Vector3(25f, 2f, 1f),  matWall);
        CreateWall("WallSouth", new Vector3(0f, 1f, -12.5f), new Vector3(25f, 2f, 1f),  matWall);
        CreateWall("WallEast",  new Vector3(12.5f, 1f, 0f),  new Vector3(1f, 2f, 25f),  matWall);
        CreateWall("WallWest",  new Vector3(-12.5f, 1f, 0f), new Vector3(1f, 2f, 25f),  matWall);

        // 7. Seis decoraciones (mezcla de cubos y cilindros, colores A y B).
        CreateDeco("Deco_01", PrimitiveType.Cube,     new Vector3( 8f, 0.5f,  8f), matDecoA);
        CreateDeco("Deco_02", PrimitiveType.Cube,     new Vector3(-8f, 0.5f,  8f), matDecoA);
        CreateDeco("Deco_03", PrimitiveType.Cube,     new Vector3( 8f, 0.5f, -8f), matDecoA);
        CreateDeco("Deco_04", PrimitiveType.Cylinder, new Vector3(-8f, 0.5f, -8f), matDecoB);
        CreateDeco("Deco_05", PrimitiveType.Cylinder, new Vector3( 5f, 0.5f,  0f), matDecoB);
        CreateDeco("Deco_06", PrimitiveType.Cube,     new Vector3(-5f, 0.5f,  0f), matDecoB);

        // 8. Directional Light.
        var light = new GameObject("Directional Light");
        var lightComp = light.AddComponent<Light>();
        lightComp.type = LightType.Directional;
        lightComp.intensity = 1f;
        lightComp.shadows = LightShadows.Soft;
        light.transform.rotation = Quaternion.Euler(50f, -30f, 0f);

        // 9. Player (Sphere con Rigidbody + tag "Player").
        var player = GameObject.CreatePrimitive(PrimitiveType.Sphere);
        player.name = "Player";
        player.tag = "Player";
        player.transform.position = new Vector3(0f, 0.5f, 0f);
        AssignMaterial(player, matPlayer);
        var rb = player.AddComponent<Rigidbody>();
        rb.constraints = RigidbodyConstraints.FreezeRotationX | RigidbodyConstraints.FreezeRotationZ;
        player.AddComponent<PlayerController>();

        // 10. Main Camera (con CameraFollow apuntando al jugador).
        var cameraGo = new GameObject("Main Camera");
        cameraGo.tag = "MainCamera";
        var cam = cameraGo.AddComponent<Camera>();
        cam.clearFlags = CameraClearFlags.SolidColor;
        cam.backgroundColor = new Color(0.15f, 0.18f, 0.25f);
        cameraGo.AddComponent<AudioListener>();
        cameraGo.transform.position = new Vector3(0f, 10f, -8f);
        cameraGo.transform.LookAt(player.transform);
        var follow = cameraGo.AddComponent<CameraFollow>();
        follow.target = player.transform;
        follow.offset = new Vector3(0f, 10f, -8f);
        follow.smoothSpeed = 0.125f;

        // 11. GameManager (vacío + script).
        var gm = new GameObject("GameManager");
        var gmComp = gm.AddComponent<GameManager>();
        gmComp.totalCollectibles = 10;

        // 12. AudioManager (vacío + 2 AudioSources + script).
        var am = new GameObject("AudioManager");
        var musicSrc = am.AddComponent<AudioSource>();
        var sfxSrc   = am.AddComponent<AudioSource>();
        musicSrc.playOnAwake = false;
        musicSrc.loop = true;
        musicSrc.volume = 0f;
        sfxSrc.playOnAwake = false;
        sfxSrc.loop = false;
        sfxSrc.volume = 0.8f;
        var amComp = am.AddComponent<AudioManager>();
        amComp.musicSource = musicSrc;
        amComp.sfxSource   = sfxSrc;
        amComp.musicVolume = 0.4f;
        amComp.sfxVolume   = 0.8f;
        amComp.fadeInDuration  = 2f;
        amComp.fadeOutDuration = 1f;

        // 13. Canvas + EventSystem + UIManager + ScorePanel + WinPanel.
        BuildCanvas(out UIManager uiManager);

        // 14. Spawnear 10 coleccionables en posiciones predefinidas.
        SpawnTenCollectibles();

        // 15. Guardar la escena.
        EditorSceneManager.SaveScene(scene, ScenePath);

        // 16. Registrar la escena en Build Settings (índice 0).
        EnsureSceneInBuildSettings();

        AssetDatabase.SaveAssets();
        AssetDatabase.Refresh();

        Debug.Log("[SceneSetup] ✔ Escena lista en " + ScenePath +
                  " (10 coleccionables, UI, audio y materiales).");
    }

    // ─── Helpers ─────────────────────────────────────────────────────────────

    private static T Load<T>(string path) where T : Object
    {
        return AssetDatabase.LoadAssetAtPath<T>(path);
    }

    private static void AssignMaterial(GameObject go, Material mat)
    {
        if (mat == null) return;
        var renderer = go.GetComponent<MeshRenderer>();
        if (renderer != null) renderer.sharedMaterial = mat;
    }

    private static void CreateWall(string name, Vector3 pos, Vector3 scale, Material mat)
    {
        var wall = GameObject.CreatePrimitive(PrimitiveType.Cube);
        wall.name = name;
        wall.transform.position = pos;
        wall.transform.localScale = scale;
        AssignMaterial(wall, mat);
    }

    private static void CreateDeco(string name, PrimitiveType type, Vector3 pos, Material mat)
    {
        var deco = GameObject.CreatePrimitive(type);
        deco.name = name;
        deco.transform.position = pos;
        AssignMaterial(deco, mat);
    }

    private static void CreateCollectiblePrefab()
    {
        Directory.CreateDirectory(PrefabsFolder);

        if (File.Exists(PrefabPath))
        {
            return; // Ya existe.
        }

        Material matDecoB = Load<Material>("Assets/Materials/MatDeco_B.mat");

        var temp = GameObject.CreatePrimitive(PrimitiveType.Capsule);
        temp.name = "Collectible";
        temp.transform.localScale = new Vector3(0.45f, 0.45f, 0.45f);
        AssignMaterial(temp, matDecoB);

        // Forzar trigger.
        var col = temp.GetComponent<CapsuleCollider>();
        if (col != null) col.isTrigger = true;

        // Script de comportamiento.
        temp.AddComponent<Collectible>();

        PrefabUtility.SaveAsPrefabAsset(temp, PrefabPath);
        Object.DestroyImmediate(temp);

        Debug.Log("[SceneSetup] Prefab creado en " + PrefabPath);
    }

    private static void SpawnTenCollectibles()
    {
        var prefab = Load<GameObject>(PrefabPath);
        if (prefab == null)
        {
            Debug.LogWarning("[SceneSetup] No se encontró el prefab del coleccionable.");
            return;
        }

        // Posiciones determinísticas (evitan radio 2u alrededor del origen).
        Vector3[] positions =
        {
            new Vector3( 3f, 1f,  3f),
            new Vector3(-4f, 1f,  5f),
            new Vector3( 6f, 1f, -3f),
            new Vector3(-6f, 1f, -5f),
            new Vector3( 9f, 1f,  0f),
            new Vector3(-9f, 1f,  0f),
            new Vector3( 0f, 1f,  9f),
            new Vector3( 0f, 1f, -9f),
            new Vector3( 4f, 1f, -7f),
            new Vector3(-7f, 1f,  4f),
        };

        var container = new GameObject("Collectibles");
        for (int i = 0; i < positions.Length; i++)
        {
            var instance = (GameObject)PrefabUtility.InstantiatePrefab(prefab);
            instance.transform.position = positions[i];
            instance.transform.SetParent(container.transform);
            instance.name = $"Collectible_{i + 1:00}";
        }
    }

    private static void BuildCanvas(out UIManager uiManagerComp)
    {
        // Canvas raíz.
        var canvasGo = new GameObject("Canvas",
            typeof(Canvas), typeof(CanvasScaler), typeof(GraphicRaycaster));
        var canvas = canvasGo.GetComponent<Canvas>();
        canvas.renderMode = RenderMode.ScreenSpaceOverlay;

        var scaler = canvasGo.GetComponent<CanvasScaler>();
        scaler.uiScaleMode = CanvasScaler.ScaleMode.ScaleWithScreenSize;
        scaler.referenceResolution = new Vector2(1920f, 1080f);

        uiManagerComp = canvasGo.AddComponent<UIManager>();

        Font defaultFont = Resources.GetBuiltinResource<Font>("LegacyRuntime.ttf");

        // ───── ScorePanel (esquina superior izquierda) ────────────────────────
        var scorePanel = new GameObject("ScorePanel", typeof(RectTransform));
        scorePanel.transform.SetParent(canvasGo.transform, false);
        var spRT = scorePanel.GetComponent<RectTransform>();
        spRT.anchorMin = new Vector2(0f, 1f);
        spRT.anchorMax = new Vector2(0f, 1f);
        spRT.pivot     = new Vector2(0f, 1f);
        spRT.anchoredPosition = new Vector2(30f, -30f);
        spRT.sizeDelta = new Vector2(400f, 80f);

        var scoreTextGo = new GameObject("ScoreText", typeof(RectTransform), typeof(Text));
        scoreTextGo.transform.SetParent(scorePanel.transform, false);
        var stRT = scoreTextGo.GetComponent<RectTransform>();
        stRT.anchorMin = Vector2.zero;
        stRT.anchorMax = Vector2.one;
        stRT.offsetMin = Vector2.zero;
        stRT.offsetMax = Vector2.zero;
        var stText = scoreTextGo.GetComponent<Text>();
        stText.text = "\u2726 0 / 10";
        stText.fontSize = 40;
        stText.alignment = TextAnchor.MiddleLeft;
        stText.color = Color.white;
        stText.font = defaultFont;

        uiManagerComp.scoreText = stText;

        // ───── WinPanel (centrado, inactivo al inicio) ────────────────────────
        var winPanel = new GameObject("WinPanel", typeof(RectTransform), typeof(CanvasGroup));
        winPanel.transform.SetParent(canvasGo.transform, false);
        var wpRT = winPanel.GetComponent<RectTransform>();
        wpRT.anchorMin = new Vector2(0.5f, 0.5f);
        wpRT.anchorMax = new Vector2(0.5f, 0.5f);
        wpRT.pivot     = new Vector2(0.5f, 0.5f);
        wpRT.sizeDelta = new Vector2(800f, 500f);
        wpRT.anchoredPosition = Vector2.zero;
        var winCG = winPanel.GetComponent<CanvasGroup>();
        winCG.alpha = 0f;

        // Fondo semitransparente.
        var bg = new GameObject("Background", typeof(RectTransform), typeof(Image));
        bg.transform.SetParent(winPanel.transform, false);
        var bgRT = bg.GetComponent<RectTransform>();
        bgRT.anchorMin = Vector2.zero;
        bgRT.anchorMax = Vector2.one;
        bgRT.offsetMin = Vector2.zero;
        bgRT.offsetMax = Vector2.zero;
        bg.GetComponent<Image>().color = new Color(0f, 0f, 0f, 0.75f);

        // Título.
        var title = new GameObject("Title", typeof(RectTransform), typeof(Text));
        title.transform.SetParent(winPanel.transform, false);
        var titleRT = title.GetComponent<RectTransform>();
        titleRT.anchorMin = new Vector2(0.5f, 0.7f);
        titleRT.anchorMax = new Vector2(0.5f, 0.7f);
        titleRT.pivot     = new Vector2(0.5f, 0.5f);
        titleRT.sizeDelta = new Vector2(700f, 120f);
        titleRT.anchoredPosition = Vector2.zero;
        var titleText = title.GetComponent<Text>();
        titleText.text = "\u00a1Ganaste!";
        titleText.fontSize = 80;
        titleText.fontStyle = FontStyle.Bold;
        titleText.alignment = TextAnchor.MiddleCenter;
        titleText.color = Color.white;
        titleText.font = defaultFont;

        // Subtítulo (✦ con conteo).
        var subtitle = new GameObject("Subtitle", typeof(RectTransform), typeof(Text));
        subtitle.transform.SetParent(winPanel.transform, false);
        var subRT = subtitle.GetComponent<RectTransform>();
        subRT.anchorMin = new Vector2(0.5f, 0.5f);
        subRT.anchorMax = new Vector2(0.5f, 0.5f);
        subRT.pivot     = new Vector2(0.5f, 0.5f);
        subRT.sizeDelta = new Vector2(700f, 80f);
        subRT.anchoredPosition = Vector2.zero;
        var subText = subtitle.GetComponent<Text>();
        subText.text = "Recogiste todos los coleccionables";
        subText.fontSize = 32;
        subText.alignment = TextAnchor.MiddleCenter;
        subText.color = new Color(1f, 1f, 1f, 0.85f);
        subText.font = defaultFont;

        // Botón "Jugar de nuevo".
        var btn = new GameObject("ReplayButton",
            typeof(RectTransform), typeof(Image), typeof(Button));
        btn.transform.SetParent(winPanel.transform, false);
        var btnRT = btn.GetComponent<RectTransform>();
        btnRT.anchorMin = new Vector2(0.5f, 0.25f);
        btnRT.anchorMax = new Vector2(0.5f, 0.25f);
        btnRT.pivot     = new Vector2(0.5f, 0.5f);
        btnRT.sizeDelta = new Vector2(360f, 90f);
        btnRT.anchoredPosition = Vector2.zero;
        var btnImg = btn.GetComponent<Image>();
        btnImg.color = new Color(0.263f, 0.380f, 0.933f, 1f); // #4361EE

        var btnTextGo = new GameObject("Text", typeof(RectTransform), typeof(Text));
        btnTextGo.transform.SetParent(btn.transform, false);
        var btnTextRT = btnTextGo.GetComponent<RectTransform>();
        btnTextRT.anchorMin = Vector2.zero;
        btnTextRT.anchorMax = Vector2.one;
        btnTextRT.offsetMin = Vector2.zero;
        btnTextRT.offsetMax = Vector2.zero;
        var btnText = btnTextGo.GetComponent<Text>();
        btnText.text = "Jugar de nuevo";
        btnText.fontSize = 36;
        btnText.alignment = TextAnchor.MiddleCenter;
        btnText.color = Color.white;
        btnText.font = defaultFont;

        // Referencias en UIManager.
        uiManagerComp.winPanel       = winPanel;
        uiManagerComp.winCanvasGroup = winCG;
        uiManagerComp.winContent     = winPanel.transform;
        uiManagerComp.replayButton   = btn.GetComponent<Button>();
        uiManagerComp.showDuration   = 0.6f;
        winPanel.SetActive(false);

        // ───── EventSystem (necesario para que los botones reciban clic) ─────
        if (Object.FindObjectOfType<EventSystem>() == null)
        {
            new GameObject("EventSystem",
                typeof(EventSystem), typeof(StandaloneInputModule));
        }
    }

    private static void EnsureSceneInBuildSettings()
    {
        var existing = EditorBuildSettings.scenes;
        foreach (var s in existing)
        {
            if (s.path == ScenePath && s.enabled) return;
        }

        var list = new System.Collections.Generic.List<EditorBuildSettingsScene>(existing)
        {
            new EditorBuildSettingsScene(ScenePath, true)
        };
        EditorBuildSettings.scenes = list.ToArray();
        Debug.Log("[SceneSetup] Main.unity agregada a Build Settings.");
    }
}
