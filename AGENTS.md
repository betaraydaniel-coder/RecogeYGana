# AGENTS.md

## Cursor Cloud specific instructions

- This repository is a Unity project for the Recoge y Gana minigame. The only required runtime service is Unity Editor / Play Mode; the standard open, play, and Windows build flow is documented in `README.md`.
- Use Unity `2022.3.20f1` for local checks. In this headless VM, run Unity batch-mode commands through `xvfb-run -a` and locate the editor with Unity Hub if `$HOME/Unity/Hub/Editor/2022.3.20f1/Editor/Unity` is not present.
- Unity package dependencies are resolved by the editor from `Packages/manifest.json` during project import, so the startup update script should stay minimal and should not start Unity services.
- Unity import, test, build, and Play Mode checks require an activated Unity license. Without one, batch mode exits with `No valid Unity Editor license found`; provide a manual Unity `.ulf` as `UNITY_LICENSE` before rerunning those checks.
