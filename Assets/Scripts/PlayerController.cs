// PlayerController.cs
// ─────────────────────────────────────────────────────────────────────────────
// Controla el movimiento físico de la esfera del jugador (Temas 1 y 5).
//
// Requisitos del GameObject "Player":
//   • Mesh: Sphere
//   • Componentes: Rigidbody + SphereCollider
//   • Tag: "Player"
//   • Material: MatPlayer (#4361EE)
//
// Movimiento basado en AddForce dentro de FixedUpdate (física estable),
// con limitación de velocidad máxima para evitar deslizamientos infinitos.
// ─────────────────────────────────────────────────────────────────────────────

using UnityEngine;

[RequireComponent(typeof(Rigidbody))]
[RequireComponent(typeof(SphereCollider))]
public class PlayerController : MonoBehaviour
{
    [Header("Parámetros de Movimiento")]
    [Tooltip("Fuerza aplicada por AddForce en cada FixedUpdate.")]
    public float moveSpeed = 12f;

    [Tooltip("Velocidad horizontal máxima permitida (m/s).")]
    public float maxSpeed = 8f;

    private Rigidbody rb;

    private void Awake()
    {
        rb = GetComponent<Rigidbody>();

        if (!CompareTag("Player"))
        {
            Debug.LogWarning("[PlayerController] El objeto no tiene el tag 'Player'. Asígnalo en el Inspector.");
        }
    }

    private void FixedUpdate()
    {
        // Lee los ejes virtuales (teclas WASD / flechas configuradas por defecto en Unity).
        float h = Input.GetAxis("Horizontal");
        float v = Input.GetAxis("Vertical");

        Vector3 input = new Vector3(h, 0f, v);

        // Aplica fuerza al Rigidbody → respuesta física natural.
        rb.AddForce(input * moveSpeed, ForceMode.Force);

        // Limitar la velocidad horizontal para no excederse de maxSpeed.
        // Nota: en Unity 6+ la propiedad se llama `linearVelocity`;
        // en Unity 2022 LTS la propiedad correcta es `velocity`.
        Vector3 horizontalVel = new Vector3(rb.velocity.x, 0f, rb.velocity.z);
        if (horizontalVel.magnitude > maxSpeed)
        {
            Vector3 limited = horizontalVel.normalized * maxSpeed;
            rb.velocity = new Vector3(limited.x, rb.velocity.y, limited.z);
        }
    }
}
