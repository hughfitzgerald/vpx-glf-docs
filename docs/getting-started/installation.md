# Installation

GLF is a single VBScript file that you include in your table script.

1. Download [`vpx-glf.vbs`](https://github.com/mpcarr/vpx-glf/raw/main/scripts/vpx-glf.vbs).
2. Paste its contents into your table script, or keep it as a separate file and load it with
   `ExecuteGlobal GetTextFile("vpx-glf.vbs")`. (For a multi-file setup that concatenates sources
   into the table at build time, see [Project layout](project-layout.md).)

The rest of this page covers what GLF expects the table itself to provide.

---

## Table objects

### Game timer

Add a **Timer** object named `Glf_GameTimer`. Set it **Enabled**, with an **Interval of -1 ms**
(fire every frame).

This timer drives the whole framework: event dispatch, delays, show steps and light updates all
run from it. Without it, nothing happens.

### Collections

Create these collections in the VPX Collections Manager (F8):

| Collection | Contents |
|---|---|
| `glf_lights` | Every light GLF should control |
| `glf_switches` | Rollover switches, kicker/scoop switches, bumper switches |
| `glf_slingshots` | Slingshot triggers |
| `glf_spinners` | Spinners |

GLF walks these collections at init and generates the `_Hit` / `_UnHit` / `_Slingshot` / `_Spin`
handlers for you. **Do not write your own handlers for objects in these collections** — GLF's
generated ones will replace them.

What goes where is easy to get wrong, so:

- **Do** add to `glf_switches`: plunger lane switch, rollovers, scoop/kicker switches, bumper
  switches, orbit and ramp switches.
- **Do not** add: flippers, drop targets, stand-up targets. Those are declared as devices instead
  (see [Flipper](../devices/flipper.md), [Targets](../devices/targets.md)), and GLF wires their
  switches itself.

Object names are case-sensitive and become event names, so a switch named `s_LeftOrb1` posts
`s_LeftOrb1_active`. A consistent `s_` prefix is conventional but not required.

### Trough

GLF expects a physical (non-destructive) trough — kickers named `swTrough1` … `swTrough7` plus a
`Drain` kicker. See [Trough](../devices/trough.md) for the full setup.

---

## Required globals

```vbscript
Const cGameName  = "MyAwesomeGame"   ' unique; used for the settings .ini and log filenames
Const BallSize   = 50                ' must be 50
Const BallMass   = 1                 ' must be 1
Const tnob       = 5                 ' total number of balls the trough holds
Const lob        = 0                 ' locked balls (captive balls not in the trough)
Dim gBot                             ' ball array, populated by GLF at init
Dim tablewidth  : tablewidth  = Table1.width
Dim tableheight : tableheight = Table1.height
```

`tnob` sets the trough size, which determines how many trough kickers GLF creates balls in.

---

## Script hooks

GLF needs five hooks. Add them to your table's event subs.

```vbscript
Sub Table1_Init
    ConfigureGlfDevices()      ' your configuration, see below
    Glf_Init(Table1)           ' must come after ConfigureGlfDevices
End Sub

Sub Table1_Exit
    Glf_Exit()                 ' persists machine variables and high scores
End Sub

Sub Table1_KeyDown(ByVal keycode)
    Glf_KeyDown(keycode)
End Sub

Sub Table1_KeyUp(ByVal keycode)
    Glf_KeyUp(keycode)
End Sub

Sub Table1_OptionEvent(ByVal eventId)
    Glf_Options(eventId)
End Sub
```

Order matters in `Table1_Init`: `ConfigureGlfDevices` must run **before** `Glf_Init`, because
`Glf_Init` pre-compiles and caches every show, shot profile and dynamic value you declared.

If you use `DisableStaticPreRendering`, wrap the options call as in the stock example:

```vbscript
Sub Table1_OptionEvent(ByVal eventId)
    If eventId = 1 Then DisableStaticPreRendering = True
    Glf_Options(eventId)
    If eventId = 3 Then DisableStaticPreRendering = False
End Sub
```

---

## The configuration entry point

All your game logic is declared in one sub:

```vbscript
Sub ConfigureGlfDevices()

    ' 1. Shows          — must come before anything that references them
    CreateGeneralShows()

    ' 2. Sounds and sound buses
    CreateSounds()

    ' 3. Shared shot profiles
    CreateSharedShotProfiles()

    ' 4. High scores and machine variables
    ' 5. Initial player variables
    ' 6. Modes           — must come after shows, sounds and profiles
    ' 7. Machine devices — ball devices, flippers, targets, magnets…
    ' 8. Segment displays

End Sub
```

The ordering constraint that actually bites: **a mode that references a show by name will silently
do nothing if the show has not been created yet.** Create shows first.

---

## Keys GLF handles

`Glf_KeyDown` / `Glf_KeyUp` post events for the standard VPX keys. You get these for free:

| Key | Events posted |
|---|---|
| Start | `s_start_active` / `s_start_inactive`, and `request_to_start_game` on release |
| Left / right flipper | `s_left_flipper_active` / `s_left_flipper_inactive`, same for right |
| Staged flippers | `s_left_staged_flipper_key_active` / `_inactive`, same for right |
| Plunger | `s_plunger_key_active` / `s_plunger_key_inactive` |
| Lockbar | `s_lockbar_key_active` / `s_lockbar_key_inactive` |
| Left / right MagnaSave | `s_left_magna_key_active` / `_inactive`, same for right |
| Add credit ×2 | `s_add_credit_key_active` / `_inactive`, `s_add_credit_key2_*` |
| Mechanical tilt | `s_tilt_warning_active` (debounced 300 ms) |
| Nudge keys | Nudges the table and accumulates virtual tilt |

Nudge keys feed a decaying tilt accumulator; when it crosses the threshold GLF posts
`s_tilt_warning_active` for you, so digital nudging and a real tilt bob behave the same way.
Sensitivity is a table option.

---

## Table options GLF adds

`Glf_Options` registers these in the VPX options dialog (F6):

| Option | Default | Purpose |
|---|---|---|
| Balls Per Game | 3 Balls | 3 or 5 |
| Tilt Sensitivity (digital nudge) | 5 | 1–10 |
| Glf Debug Log | Off | Writes to `glf_logs/` |
| Glf Debug Log Level | Info | Info or Debug |
| Glf Frame Dispatch | 5 | Max events dispatched per frame (5–50) |
| Glf Min Lightmap Update Rate | Disabled | Throttles lightmap colour updates |
| Glf Monitor | Off | Connects the external GLF monitor |
| Glf Backbox Control Protocol | Off | Connects a BCP media controller |
| Glf Virtual Segment DMD | Off | Mirrors segment displays to a FlexDMD window |

See [Debugging](../appendix/debugging.md) for what to reach for when something misbehaves.

---

## Next

[Your first mode →](first-mode.md)
