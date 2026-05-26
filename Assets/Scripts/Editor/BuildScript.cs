// BuildScript.cs
// ─────────────────────────────────────────────────────────────────────────────
// Automatiza el Build del minijuego desde el menú de Unity O desde la nube
// (GitHub Actions / game-ci).
//
// Uso local:
//   • Menú "Recoge y Gana → Build Windows64"  →  ../Build/RecogeYGana/RecogeYGana.exe
//
// Uso en la nube (game-ci):
//   • El workflow llama a BuildScript.BuildGameCI mediante el parámetro
//     "buildMethod". La salida queda en build/StandaloneWindows64/RecogeYGana.exe
//     y se sube como artefacto descargable.
//
// En ambos casos, si la escena Main.unity no existe se construye al vuelo
// llamando a SceneSetup.BuildEverything().
// ─────────────────────────────────────────────────────────────────────────────

using System.IO;
using UnityEditor;
using UnityEditor.Build.Reporting;
using UnityEngine;

public static class BuildScript
{
    private const string LocalOutputDir   = "../Build/RecogeYGana";
    private const string CloudOutputDir   = "build/StandaloneWindows64";
    private const string ExecutableName   = "RecogeYGana.exe";
    private const string ScenePath        = "Assets/Scenes/Main.unity";

    [MenuItem("Recoge y Gana/Build Windows64")]
    public static void BuildGame()
    {
        BuildToFolder(LocalOutputDir);
    }

    /// <summary>
    /// Punto de entrada que utiliza GitHub Actions / game-ci.
    /// </summary>
    public static void BuildGameCI()
    {
        BuildToFolder(CloudOutputDir);
    }

    private static void BuildToFolder(string folder)
    {
        // Garantizar que la escena exista (si el usuario no ha abierto Unity nunca).
        if (!File.Exists(ScenePath))
        {
            Debug.Log("[BuildScript] Main.unity no existe → llamando a SceneSetup.BuildEverything().");
            SceneSetup.BuildEverything();
        }

        string outputDir  = Path.GetFullPath(folder);
        string outputPath = Path.Combine(outputDir, ExecutableName);

        if (!Directory.Exists(outputDir))
        {
            Directory.CreateDirectory(outputDir);
        }

        // Recoger escenas habilitadas en Build Settings.
        EditorBuildSettingsScene[] active = EditorBuildSettings.scenes;
        var enabled = new System.Collections.Generic.List<string>();
        foreach (var s in active)
        {
            if (s.enabled) enabled.Add(s.path);
        }

        if (enabled.Count == 0)
        {
            // Fallback: si Build Settings está vacío, usar Main.unity directamente.
            enabled.Add(ScenePath);
        }

        BuildPlayerOptions opts = new BuildPlayerOptions
        {
            scenes           = enabled.ToArray(),
            locationPathName = outputPath,
            target           = BuildTarget.StandaloneWindows64,
            options          = BuildOptions.None,
        };

        BuildReport report = BuildPipeline.BuildPlayer(opts);
        BuildSummary summary = report.summary;

        if (summary.result == BuildResult.Succeeded)
        {
            Debug.Log($"[BuildScript] ✔ Build OK ({summary.totalSize / 1024 / 1024} MB) → {outputPath}");
        }
        else
        {
            Debug.LogError($"[BuildScript] ✘ Build falló: {summary.result}");
            EditorApplication.Exit(1);
        }
    }
}
