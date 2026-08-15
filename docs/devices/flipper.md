# Flipper

Wraps a flipper so GLF can enable and disable it with game state.

```vbscript
With CreateGlfFlipper("left")
    .Switch = "s_left_flipper"
    .ActionCallback = "LeftFlipperAction"
    .EnableEvents  = Array("ball_started", "enable_flippers")
    .DisableEvents = Array("kill_flippers")
End With
```

**Flippers must not be in `glf_switches`** or any other GLF collection — the device wires the
switch itself.

## Configuration

| Property | Default | Meaning |
|---|---|---|
| `Switch` | — | The flipper button switch |
| `ActionCallback` | — | Function that drives the flipper |
| `EnableEvents` | `ball_started` | Events that enable it |
| `DisableEvents` | `ball_will_end`, `service_mode_entered` | Events that disable it |

The defaults are usually right: flippers live from ball start to ball end. Adding
`"enable_flippers"` / `"kill_flippers"` gives you a hook for mode-driven flipper control.

## The action callback

Receives 1 on press and 0 on release:

```vbscript
Sub LeftFlipperAction(enabled)
    If enabled = 1 Then
        LeftFlipper.RotateToEnd
        RandomSoundFlipperUpLeft LeftFlipper
    Else
        LeftFlipper.RotateToStart
        RandomSoundFlipperDownLeft LeftFlipper
    End If
End Sub
```

This is where flipper physics, sounds and any polarity correction live — GLF only decides *whether*
the flipper responds.

## Switch names

Use the switch names GLF's key handling posts:

| Flipper | Switch |
|---|---|
| Left | `s_left_flipper` |
| Right | `s_right_flipper` |
| Upper left (staged) | `s_left_staged_flipper_key` |
| Upper right (staged) | `s_right_staged_flipper_key` |

```vbscript
With CreateGlfFlipper("upper_left")
    .Switch = "s_left_staged_flipper_key"
    .ActionCallback = "UpperLeftFlipperAction"
    .EnableEvents  = Array("ball_started", "enable_flippers")
    .DisableEvents = Array("kill_flippers")
End With
```

## Events posted

| Event | When |
|---|---|
| `flipper_<name>_activate` | Button pressed while enabled |
| `flipper_<name>_deactivate` | Button released |

For game logic that reacts to flipper presses regardless of whether flippers are live — menus,
skill shot selection — listen for `s_left_flipper_active` instead. Those post even when flippers
are disabled.

## Killing flippers

```vbscript
With .EventPlayer()
    .Add "mode_cutscene_started",  Array("kill_flippers")
    .Add "mode_cutscene_stopping", Array("enable_flippers")
End With
```

GLF also disables flippers automatically on tilt and at ball end.

## See also

[Auto-fire device](auto-fire.md) · [Installation](../getting-started/installation.md)
