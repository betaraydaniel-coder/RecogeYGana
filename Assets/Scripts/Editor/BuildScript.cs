// BuildScript.cs
// ─────────────────────────────────────────────────────────────────────────────
// Automatiza el Build del minijuego desde el menú de Unity.
//
// Uso: Recoge y Gana → Build Windows64
// Salida: ../Build/RecogeYGana/RecogeYGana.exe (junto a la carpeta del proyecto).
// ─────────────────────────────────────────────────────────────────────────────

using System.IO;
using UnityEditor;
using UnityEditor.Build.Reporting;
using UnityEngine;

public static class BuildScript
{
    private const string BuildFolder   = "../Build/RecogeYGana";
    private const string ExecutableName = "RecogeYGana.exe";

    [MenuItem("Recoge y Gana/Build Windows64")]
    public static void BuildGame()
    {
        string outputDir  = Path.GetFullPath(BuildFolder);
        string outputPath = Path.Combine(outputDir, ExecutableName);

        if (!Directory.Exists(outputDir))
        {
            Directory.CreateDirectory(outputDir);
        }

        // Recoge todas las escenas activas en Build Settings.
        EditorBuildSettingsScene[] activeScenes = EditorBuildSettings.scenes;
        string[] scenes = new string[activeScenes.Length];
        int idx = 0;
        foreach (var s in activeScenes)
        {
            if (s.enabled) scenes[idx++] = s.path;
        }
        System.Array.Resize(ref scenes, idx);

        if (scenes.Length == 0)
        {
            Debug.LogError("[BuildScript] No hay escenas habilitadas en Build Settings. " +
                           "Abre File → Build Settings y agrega la escena principal.");
            return;
        }

        BuildPlayerOptions opts = new BuildPlayerOptions
        {
            scenes           = scenes,
            locationPathName = outputPath,
            target           = BuildTarget.StandaloneWindows64,
            options          = BuildOptions.None
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
        }
    }
}
