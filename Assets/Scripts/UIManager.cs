// UIManager.cs
// ─────────────────────────────────────────────────────────────────────────────
// Maneja la interfaz de usuario del minijuego (Temas 3 y 6).
//
// Jerarquía esperada en el Canvas (la genera SceneSetup automáticamente):
//   Canvas (Screen Space - Overlay)
//   ├── ScorePanel        (esquina superior izquierda)
//   │   └── ScoreText     → UI.Text → asignado al campo "scoreText"
//   └── WinPanel          (centrado, inactivo al inicio)
//       ├── Background    (Image semitransparente)
//       ├── TitleText     → "¡Ganaste!"
//       └── ReplayButton  → texto "Jugar de nuevo" → OnClick: UIManager.Replay()
//
// Se usa UI.Text clásico (no TextMeshPro) para mantener el proyecto ligero
// y compatible con cualquier instalación de Unity 2022.3+ sin paquetes extra.
// ─────────────────────────────────────────────────────────────────────────────

using System.Collections;
using UnityEngine;
using UnityEngine.SceneManagement;
using UnityEngine.UI;

public class UIManager : MonoBehaviour
{
    public static UIManager Instance { get; private set; }

    [Header("Referencias de UI - Score")]
    [Tooltip("Texto del marcador (UI.Text clásico).")]
    public Text scoreText;

    [Header("Referencias de UI - WinScreen")]
    [Tooltip("Panel raíz que se muestra al ganar (inactivo al iniciar).")]
    public GameObject winPanel;
    [Tooltip("CanvasGroup del WinPanel para controlar el alpha durante el fade.")]
    public CanvasGroup winCanvasGroup;
    [Tooltip("Transform del contenido del WinPanel para el efecto de escala.")]
    public Transform   winContent;
    [Tooltip("Botón 'Jugar de nuevo'.")]
    public Button      replayButton;

    [Header("Animación")]
    [Tooltip("Duración de la animación de aparición (segundos).")]
    public float showDuration = 0.6f;

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
        if (winPanel != null) winPanel.SetActive(false);
        if (winCanvasGroup != null) winCanvasGroup.alpha = 0f;
        if (replayButton  != null) replayButton.onClick.AddListener(Replay);
    }

    /// <summary>
    /// Actualiza el contador "✦ {c} / {t}".
    /// </summary>
    public void UpdateScore(int collected, int total)
    {
        string texto = $"\u2726 {collected} / {total}";
        if (scoreText != null) scoreText.text = texto;
    }

    /// <summary>
    /// Muestra la pantalla de victoria con fade-in y un efecto de escala 0.8 → 1.0.
    /// </summary>
    public void ShowWinScreen()
    {
        if (winPanel == null)
        {
            Debug.LogWarning("[UIManager] WinPanel no está asignado en el Inspector.");
            return;
        }

        winPanel.SetActive(true);
        StartCoroutine(AnimateWinScreen());
    }

    private IEnumerator AnimateWinScreen()
    {
        float t = 0f;

        Vector3 startScale = Vector3.one * 0.8f;
        Vector3 endScale   = Vector3.one;

        if (winContent != null)      winContent.localScale = startScale;
        if (winCanvasGroup != null)  winCanvasGroup.alpha  = 0f;

        while (t < showDuration)
        {
            t += Time.unscaledDeltaTime;
            float k = Mathf.Clamp01(t / showDuration);
            float eased = 1f - Mathf.Pow(1f - k, 3f); // ease-out cubic

            if (winCanvasGroup != null) winCanvasGroup.alpha = eased;
            if (winContent != null)     winContent.localScale = Vector3.Lerp(startScale, endScale, eased);

            yield return null;
        }

        if (winCanvasGroup != null) winCanvasGroup.alpha = 1f;
        if (winContent != null)     winContent.localScale = endScale;
    }

    /// <summary>
    /// Recarga la escena 0 (build index) reiniciando la partida.
    /// </summary>
    public void Replay()
    {
        SceneManager.LoadScene(0);
    }
}
