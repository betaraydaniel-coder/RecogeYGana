// CollectibleSpawner.cs
// ─────────────────────────────────────────────────────────────────────────────
// Herramienta de editor que instancia 10 coleccionables en posiciones
// aleatorias dentro de la zona jugable (Tema 3).
//
// Reglas:
//   • Rango X/Z: -11 a 11 (dentro de las paredes).
//   • Se excluye un radio de 2 unidades alrededor del origen
//     para que el jugador no aparezca encima de un coleccionable.
//
// Uso: Recoge y Gana → Spawnear Coleccionables (x10)
// Antes de ejecutarlo, asigna el campo "Collectible Prefab" en una instancia
// del menú (o coloca el prefab en Assets/Prefabs/Collectible.prefab y se
// detectará automáticamente).
// ─────────────────────────────────────────────────────────────────────────────

using UnityEditor;
using UnityEngine;

public static class CollectibleSpawner
{
    private const int   TotalCollectibles   = 10;
    private const float MinX                = -11f;
    private const float MaxX                =  11f;
    private const float MinZ                = -11f;
    private const float MaxZ                =  11f;
    private const float ExclusionRadius     =  2f;
    private const float DefaultY            =  1f;
    private const string PrefabPath         = "Assets/Prefabs/Collectible.prefab";

    [MenuItem("Recoge y Gana/Spawnear Coleccionables (x10)")]
    public static void SpawnCollectibles()
    {
        GameObject prefab = AssetDatabase.LoadAssetAtPath<GameObject>(PrefabPath);

        if (prefab == null)
        {
            Debug.LogError($"[CollectibleSpawner] No se encontró el prefab en {PrefabPath}.\n" +
                           "Crea un prefab llamado 'Collectible' en Assets/Prefabs/ con el script Collectible.cs.");
            return;
        }

        // Contenedor opcional para mantener ordenada la jerarquía.
        GameObject contenedor = GameObject.Find("Collectibles");
        if (contenedor == null)
        {
            contenedor = new GameObject("Collectibles");
            Undo.RegisterCreatedObjectUndo(contenedor, "Crear contenedor Collectibles");
        }

        int spawneados = 0;
        int intentos   = 0;
        const int maxIntentos = 500;

        while (spawneados < TotalCollectibles && intentos < maxIntentos)
        {
            intentos++;

            float x = Random.Range(MinX, MaxX);
            float z = Random.Range(MinZ, MaxZ);

            // Excluir radio alrededor del origen (posición inicial del jugador).
            if (Mathf.Sqrt(x * x + z * z) < ExclusionRadius) continue;

            Vector3 pos = new Vector3(x, DefaultY, z);

            GameObject instancia = (GameObject)PrefabUtility.InstantiatePrefab(prefab);
            instancia.transform.position = pos;
            instancia.transform.SetParent(contenedor.transform);
            instancia.name = $"Collectible_{spawneados + 1:00}";

            Undo.RegisterCreatedObjectUndo(instancia, "Spawn Collectible");
            spawneados++;
        }

        Debug.Log($"[CollectibleSpawner] Spawneados {spawneados}/{TotalCollectibles} coleccionables (intentos: {intentos}).");
    }
}
