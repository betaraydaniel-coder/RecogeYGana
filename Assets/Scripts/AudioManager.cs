// AudioManager.cs
// ─────────────────────────────────────────────────────────────────────────────
// Singleton de audio del minijuego (Tema 7).
//
// Funcionalidades:
//   • Música de fondo en loop a volumen 0.4 con fade-in de 2s al iniciar.
//   • Fade-out de la música a 0 en 1s cuando el jugador gana.
//   • Reproducción de SFX (volumen 0.8) — incluye PlayCollect() para los
//     coleccionables.
//   • DontDestroyOnLoad para sobrevivir al recargar la escena.
//
// ─── CÓMO OBTENER LOS AUDIOS (CC0 / dominio público) ─────────────────────────
//   1. Visita https://freesound.org/ y crea una cuenta gratuita.
//   2. Filtra por Licencia → "Creative Commons 0" para usar libremente.
//   3. Sugerencias de búsqueda:
//        • Música de fondo: "ambient loop", "chiptune loop", "calm background".
//        • Recoger: "pickup coin", "collect item", "ding short".
//   4. Descarga los .wav / .ogg y colócalos en Assets/Audio/.
//   5. Renómbralos a:
//        • bgm_loop.ogg
//        • sfx_collect.wav
//   6. Arrástralos a los campos "Background Music" y "Collect Sfx" del
//      GameObject "AudioManager" en el Inspector.
// ─────────────────────────────────────────────────────────────────────────────

using System.Collections;
using UnityEngine;

[RequireComponent(typeof(AudioSource))]
public class AudioManager : MonoBehaviour
{
    public static AudioManager Instance { get; private set; }

    [Header("AudioSources")]
    [Tooltip("AudioSource para la música de fondo (loop).")]
    public AudioSource musicSource;
    [Tooltip("AudioSource para efectos de sonido (one-shot).")]
    public AudioSource sfxSource;

    [Header("Clips")]
    [Tooltip("Música de fondo en loop.")]
    public AudioClip backgroundMusic;
    [Tooltip("Sonido al recoger un coleccionable.")]
    public AudioClip collectSfx;

    [Header("Volúmenes objetivo")]
    [Range(0f, 1f)] public float musicVolume = 0.4f;
    [Range(0f, 1f)] public float sfxVolume   = 0.8f;

    [Header("Fades")]
    public float fadeInDuration  = 2f;
    public float fadeOutDuration = 1f;

    private Coroutine fadeRoutine;

    private void Awake()
    {
        if (Instance != null && Instance != this)
        {
            Destroy(gameObject);
            return;
        }
        Instance = this;
        DontDestroyOnLoad(gameObject);

        // Configuración por defecto de los AudioSources (por si no se asignaron).
        if (musicSource == null)
        {
            musicSource = gameObject.AddComponent<AudioSource>();
        }
        if (sfxSource == null)
        {
            sfxSource = gameObject.AddComponent<AudioSource>();
        }

        musicSource.loop        = true;
        musicSource.playOnAwake = false;
        musicSource.volume      = 0f;

        sfxSource.loop          = false;
        sfxSource.playOnAwake   = false;
        sfxSource.volume        = sfxVolume;
    }

    private void Start()
    {
        if (backgroundMusic != null)
        {
            musicSource.clip = backgroundMusic;
            musicSource.Play();
            fadeRoutine = StartCoroutine(FadeMusic(0f, musicVolume, fadeInDuration));
        }
        else
        {
            Debug.LogWarning("[AudioManager] No se asignó música de fondo. Coloca un clip CC0 en el campo 'Background Music'.");
        }
    }

    /// <summary>
    /// Reproduce el SFX de "recoger un coleccionable".
    /// </summary>
    public void PlayCollect()
    {
        if (collectSfx == null)
        {
            Debug.LogWarning("[AudioManager] No se asignó el clip 'collectSfx'.");
            return;
        }
        sfxSource.PlayOneShot(collectSfx, sfxVolume);
    }

    /// <summary>
    /// Inicia un fade-out hasta volumen 0 en `duration` segundos.
    /// </summary>
    public void FadeOutMusic(float duration)
    {
        if (fadeRoutine != null) StopCoroutine(fadeRoutine);
        fadeRoutine = StartCoroutine(FadeMusic(musicSource.volume, 0f, duration));
    }

    private IEnumerator FadeMusic(float from, float to, float duration)
    {
        float t = 0f;
        musicSource.volume = from;

        while (t < duration)
        {
            t += Time.deltaTime;
            float k = Mathf.Clamp01(t / duration);
            musicSource.volume = Mathf.Lerp(from, to, k);
            yield return null;
        }

        musicSource.volume = to;

        if (Mathf.Approximately(to, 0f)) musicSource.Stop();
    }
}
