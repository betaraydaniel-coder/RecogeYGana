// MaterialSetup.cs
// ─────────────────────────────────────────────────────────────────────────────
// Script de editor que genera automáticamente los 5 materiales del proyecto
// "Recoge y Gana" (Temas 1, 2 y 4 del curso).
//
// Se ubica en la carpeta Editor porque utiliza UnityEditor / AssetDatabase
// (esto evita que el build de Standalone falle al compilar).
//
// Uso: en la barra superior de Unity → Recoge y Gana → Crear Materiales.
// ─────────────────────────────────────────────────────────────────────────────

using UnityEditor;
using UnityEngine;
using System.IO;

public static class MaterialSetup
{
    private const string MaterialsFolder = "Assets/Materials";

    [MenuItem("Recoge y Gana/Crear Materiales")]
    public static void CrearMateriales()
    {
        if (!Directory.Exists(MaterialsFolder))
        {
            Directory.CreateDirectory(MaterialsFolder);
            AssetDatabase.Refresh();
        }

        CrearMaterial("MatGround",   "#2D6A4F");
        CrearMaterial("MatWall",     "#4A4E69");
        CrearMaterial("MatDeco_A",   "#C1666B");
        CrearMaterial("MatDeco_B",   "#E8C547");
        CrearMaterial("MatPlayer",   "#4361EE");

        AssetDatabase.SaveAssets();
        AssetDatabase.Refresh();

        Debug.Log("[MaterialSetup] 5 materiales creados en " + MaterialsFolder);
    }

    private static void CrearMaterial(string nombre, string hexColor)
    {
        string ruta = $"{MaterialsFolder}/{nombre}.mat";

        Material mat = AssetDatabase.LoadAssetAtPath<Material>(ruta);
        if (mat == null)
        {
            mat = new Material(Shader.Find("Standard"));
            AssetDatabase.CreateAsset(mat, ruta);
        }

        if (ColorUtility.TryParseHtmlString(hexColor, out Color color))
        {
            mat.color = color;
        }
        else
        {
            Debug.LogWarning($"[MaterialSetup] Color HEX inválido para {nombre}: {hexColor}");
        }

        EditorUtility.SetDirty(mat);
    }
}
