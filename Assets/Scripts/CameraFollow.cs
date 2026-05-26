// CameraFollow.cs
// ─────────────────────────────────────────────────────────────────────────────
// Cámara que sigue suavemente al jugador usando Vector3.Lerp (Tema 5).
//
// Cómo usar:
//   1. Asigna este script a la Main Camera.
//   2. Arrastra el GameObject del jugador al campo "target" en el Inspector.
//   3. El offset por defecto (0, 10, -8) ofrece una vista isométrica clásica.
// ─────────────────────────────────────────────────────────────────────────────

using UnityEngine;

public class CameraFollow : MonoBehaviour
{
    [Header("Objetivo a seguir")]
    [Tooltip("Transform del jugador (esfera).")]
    public Transform target;

    [Header("Configuración de Cámara")]
    [Tooltip("Desplazamiento relativo respecto al jugador.")]
    public Vector3 offset = new Vector3(0f, 10f, -8f);

    [Tooltip("Factor de suavizado para Vector3.Lerp (0 = no se mueve, 1 = sigue instantáneo).")]
    [Range(0.01f, 1f)]
    public float smoothSpeed = 0.125f;

    [Tooltip("Si está activo, la cámara siempre mira hacia el jugador.")]
    public bool lookAtTarget = true;

    private void LateUpdate()
    {
        if (target == null) return;

        // Posición deseada = posición del jugador + offset configurable.
        Vector3 desiredPosition = target.position + offset;

        // Interpolación lineal suave entre la posición actual y la deseada.
        Vector3 smoothed = Vector3.Lerp(transform.position, desiredPosition, smoothSpeed);
        transform.position = smoothed;

        // Apuntar hacia el jugador (efecto "follow cam" típico).
        if (lookAtTarget)
        {
            transform.LookAt(target);
        }
    }
}
