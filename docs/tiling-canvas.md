# Tiling canvas (Omarchy-style layout) in `rapid_main`

Status (2026-10-08): **implemented and tested in software.** It still needs
operator review on the lab displays.

The main window no longer swaps one page at a time. Every panel, and every
non-modal tool window, is a **tile** on a canvas. The model is a tiling window
manager such as Omarchy/Hyprland:

* tiles always fill the space, so nothing overlaps and nothing floats over
  your work;
* opening a panel splits the focused tile along its longer side ("dwindle");
* closing a tile gives its space back to its neighbour;
* tiles are grouped into **workspaces** 1–5, shown at the top of the sidebar.

The VB6-style one-panel-at-a-time view is still available as **Classic pages**.

## Look

The main app uses a macOS-style glass theme:

* translucent sidebar and toolbar over a soft colour field;
* macOS-style controls and menus;
* tiles drawn as macOS windows.

Each tile has **traffic lights**:

* 🔴 close;
* 🟡 float (tool windows only; grey for panels);
* 🟢 monocle.

The glyphs appear when the pointer is over the title bar. As on macOS, only
the focused tile shows coloured lights.

## Default workspaces (View → Reset Layout restores them)

| Workspace | Tiles |
|---|---|
| 1 Run | Dashboard \| Live Measure |
| 2 Queue | Sample Queue |
| 3 Program | Sequence |
| 4 Setup | Settings \| Calibration |
| 5 | empty: drop anything here |

Sidebar buttons (Dashboard, Sample Queue, …) still work. They **focus** the
panel's tile and jump to its workspace. If the panel is closed, they reopen it
in the current workspace.

## Mouse

* **Resize:** drag the gap between two tiles.
* **Move / dock:** drag a tile by its title bar onto another tile. The
  highlighted half shows where it lands. The left, right, top or bottom band
  docks it on that side; the centre swaps the two tiles.
* **Monocle:** double-click a title bar (or the green light) to fill the
  canvas with that tile, and again to restore.
* **Close:** the red light. A panel can be reopened from the sidebar or the
  launcher.
* **Float:** the yellow light on a tool-window tile (Step Monitor, Debug Console, Webcam,
  Vacuum, DC Motors) pops it out as a normal window. Opening the tool again
  tiles it back.

## Keyboard

Windows reserves the Super key, so the modifier is **Ctrl+Alt**.

| Keys | Action |
|---|---|
| Ctrl+Alt+Arrows | Focus the tile in that direction |
| Ctrl+Alt+Shift+Arrows | Swap the focused tile with its neighbour |
| Ctrl+Alt+= / − | Grow / shrink the focused tile |
| Ctrl+Alt+J | Toggle split direction (side by side ↔ stacked) |
| Ctrl+Alt+B | Balance the splits on this workspace |
| Ctrl+Alt+F | Monocle (focused tile fills the canvas) |
| Ctrl+Alt+W | Close the focused tile |
| Ctrl+Alt+T | Float / tile a tool window |
| Ctrl+Alt+1 … 5 | Switch workspace |
| Ctrl+Alt+Shift+1 … 5 | Send the focused tile to a workspace |
| Ctrl+Alt+Space | Launcher: type to open a panel or tool |
| Ctrl+Alt+C | Classic pages on/off |
| Ctrl+Alt+K | Show the key bindings |

The same commands are under **View → Tiling**.

## Behaviour that matters for safety

* The header keeps **Pause**, **Halt**, the run state, sample and step visible
  in every arrangement. On narrow windows it hides the step text first, then
  the sample text and title, then reduces No-Comm and Exit to icons. Pause,
  Halt and the run state are never truncated.
* Tool windows that own hardware (Vacuum, DC Motors) still hold their device
  lease while tiled. Closing the tile asks the window to close in the normal
  way. A window that refuses because work is active keeps its tile.
* Modal dialogs (SQUID settings, Susceptibility, IRM/ARM, login, About) remain
  modal windows. They are never tiled.
* A tile never squeezes a panel's height below what the panel needs; the tile
  scrolls vertically instead. Panels already adapt their width, so tiles
  hand them exactly the tile width.

## Persistence

* The panel arrangement (workspaces, splits, ratios, focus, monocle, classic
  mode) is saved in the RAPID QSettings key `ui/tiling_state` on exit, and
  restored at the next launch.
* Tool windows are not reopened automatically.
* **View → Tiling → Tile tool windows into the canvas** switches tool-window
  docking off for operators who prefer separate windows.

## Implementation

| Piece | Location |
|---|---|
| Tiling engine (tree, dwindle, focus/swap, workspaces, monocle, classic, drag-dock, launcher, key bindings, workspace bar, theme) | `RapidPy/rapidpy_common/tiling.py` |
| Main-window integration (panel tiles, default workspaces, tool docking, persistence, menu, adaptive header) | `RapidPy/rapid_main/rapid_main/app.py` |
| Tests | `RapidPy/rapid_main/tests/test_tiling_canvas.py` |

`TilingCanvas` also implements the former page-stack API (`setCurrentIndex`,
`currentWidget`, `currentChanged`, …). Code that "switched pages" therefore now
opens or focuses tiles without other changes. The engine is in `rapidpy_common`
so standalone apps can adopt it later.
