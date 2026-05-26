// Collectible.cs
// ─────────────────────────────────────────────────────────────────────────────
// Objeto coleccionable que rota, flota y se destruye al ser recogido por el
// jugador (Temas 3 y 5).
//
// Requisitos del Prefab "Collectible":
//   • Cualquier Mesh (cápsula, cubo, gema personalizada, etc.)
//   • Collider con IsTrigger = true
//   • Material decorativo (MatDeco_A o MatDeco_B)
// ─────────────────────────────────────────────────────────────────────────────

using UnityEngine;

[RequireComponent(typeof(Collider))]
public class Collectible : MonoBehaviour
{
    [Header("Rotación")]
    [Tooltip("Velocidad de rotación en el eje Y (grados/segundo).")]
    public float rotSpeedY = 90f;

    [Tooltip("Velocidad de rotación en el eje X (grados/segundo).")]
    public float rotSpeedX = 45f;

    [Header("Movimiento Flotante (Bob)")]
    [Tooltip("Amplitud del desplazamiento vertical en unidades.")]
    public float bobAmplitude = 0.3f;

    [Tooltip("Frecuencia del movimiento (ciclos por segundo).")]
    public float bobFrequency = 2f;

    private Vector3 startPos;

    private void Awake()
    {
        // Forzar que el Collider sea Trigger por seguridad.
        Collider col = GetComponent<Collider>();
        col.isTrigger = true;

        startPos = transform.position;
    }

    private void Update()
    {
        // Rotación doble eje (Y: 90°/s, X: 45°/s).
        transform.Rotate(rotSpeedX * Time.deltaTime,
                         rotSpeedY * Time.deltaTime,
                         0f,
                         Space.World);

        // Movimiento flotante usando Mathf.Sin → genera un bob suave.
        float yOffset = Mathf.Sin(Time.time * bobFrequency) * bobAmplitude;
        transform.position = new Vector3(startPos.x,
                                         startPos.y + yOffset,
                                         startPos.z);
    }

    private void OnTriggerEnter(Collider other)
    {
        if (!other.CompareTag("Player")) return;

        if (GameManager.Instance != null)
        {
            GameManager.Instance.CollectItem();
        }

        if (AudioManager.Instance != null)
        {
            AudioManager.Instance.PlayCollect();
        }

        Destroy(gameObject);
    }
}
