// GameManager.cs
// ─────────────────────────────────────────────────────────────────────────────
// Singleton que controla el estado global del minijuego (Tema 3).
//
// Responsabilidades:
//   • Llevar la cuenta de coleccionables recogidos.
//   • Notificar a UIManager para actualizar el HUD.
//   • Detectar la condición de victoria y dispararla.
//   • Solicitar fade-out de la música a AudioManager al ganar.
// ─────────────────────────────────────────────────────────────────────────────

using UnityEngine;

public class GameManager : MonoBehaviour
{
    public static GameManager Instance { get; private set; }

    [Header("Configuración")]
    [Tooltip("Total de coleccionables que deben recogerse para ganar.")]
    public int totalCollectibles = 10;

    [Header("Estado (solo lectura en runtime)")]
    [SerializeField] private int collected = 0;
    [SerializeField] private bool gameWon  = false;

    public int Collected => collected;
    public int Total     => totalCollectibles;
    public bool GameWon  => gameWon;

    private void Awake()
    {
        if (Instance != null && Instance != this)
        {
            Destroy(gameObject);
            return;
        }
        Instance = this;
    }

    private void Start()
    {
        if (UIManager.Instance != null)
        {
            UIManager.Instance.UpdateScore(collected, totalCollectibles);
        }
    }

    /// <summary>
    /// Llamado por cada Collectible al ser recogido por el jugador.
    /// </summary>
    public void CollectItem()
    {
        if (gameWon) return;

        collected++;

        if (UIManager.Instance != null)
        {
            UIManager.Instance.UpdateScore(collected, totalCollectibles);
        }

        if (collected >= totalCollectibles)
        {
            WinGame();
        }
    }

    private void WinGame()
    {
        gameWon = true;
        Debug.Log("[GameManager] ¡Victoria! Todos los coleccionables fueron recogidos.");

        if (UIManager.Instance != null)
        {
            UIManager.Instance.ShowWinScreen();
        }

        if (AudioManager.Instance != null)
        {
            AudioManager.Instance.FadeOutMusic(1f);
        }
    }
}
