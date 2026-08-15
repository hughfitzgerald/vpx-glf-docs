# Tilt

Handles tilt warnings and the tilt itself. One tilt device per table, in a high-priority mode of
its own.

```vbscript
Sub CreateTiltMode()

    With CreateGlfMode("tilt", 10000)
        .StartEvents = Array("ball_started")
        .StopEvents  = Array("ball_will_end")

        With .Tilt()
            .MultipleHitWindow = 3000
            .SettleTime = 5000
            .WarningsToTilt = 3
            .ResetWarningEvents = Array("mode_tilt_started")
        End With
    End With

End Sub
```

## Configuration

| Property | Default | Meaning |
|---|---|---|
| `WarningsToTilt` | 0 | Warnings before the ball tilts |
| `MultipleHitWindow` | 0 | Ms after a warning during which further hits are ignored |
| `SettleTime` | 0 | Ms of quiet required before the tilt clears |
| `ResetWarningEvents` | — | Events that reset the warning count |

`WarningsToTilt` defaults to 0, which tilts on the first hit. Always set it.

All accept [placeholders](../concepts/placeholders.md), so sensitivity can follow a difficulty
setting.

`ResetWarningEvents = Array("mode_tilt_started")` gives the usual behaviour of a fresh set of
warnings each ball. Omit it and warnings accumulate across the whole game.

## Input

The tilt device listens for `s_tilt_warning_active`. GLF posts that from two sources
automatically: the mechanical tilt key (debounced), and the digital nudge keys once accumulated
nudging crosses the threshold set by the *Tilt Sensitivity* table option. A real tilt bob switch in
`glf_switches` named `s_tilt_warning` feeds the same path.

`MultipleHitWindow` stops one violent nudge registering as three warnings — a single shake usually
bounces the bob several times.

## Events posted

| Event | Kwargs | When |
|---|---|---|
| `tilt_warning` | `warnings`, `warnings_remaining` | A warning was issued |
| `tilt_warning_<n>` | | Warning number *n* — `tilt_warning_1`, `tilt_warning_2`… |
| `tilt` | | The ball has tilted |
| `tilt_mode_<modename>_clear` | | Tilt cleared after the settle time |

The clear event includes the mode's full internal name, so a tilt device in a mode named `tilt`
posts `tilt_mode_tilt_clear`.

Numbered warning events make escalating feedback easy:

```vbscript
With .EventPlayer()
    .Add "tilt_warning_1", Array("play_sfx_tilt_warning", "show_warning_1")
    .Add "tilt_warning_2", Array("play_sfx_tilt_warning", "show_warning_2")
    .Add "tilt",           Array("play_sfx_tilt", "show_tilt_slide")
End With
```

## What tilting does

When the ball tilts, GLF sets the game to tilted, disables flippers and auto-fire devices, and
ends the ball once all balls have drained. While tilted, switch hits post no events at all — so
scoring stops dead without your rules needing to check.

Modes should stop themselves on `tilt`:

```vbscript
With CreateGlfMode("base", 110)
    .StopEvents = Array("ball_ended", "tilt")
End With
```

## Settle time

`SettleTime` requires a period with no further tilt hits before the tilt clears and the next ball
starts, so a player who keeps shaking the machine waits longer.

The warning count is kept in the player variable `tilt_warnings`, so it can be read like any other:

```vbscript
.Add "mode_base_started{current_player.tilt_warnings > 0}", Array("show_warning_reminder")
```

## Notes

- Give the tilt mode a very high priority so its handlers run before anything else.
- Stop it on `ball_will_end` rather than `ball_ended`, so it isn't live while the ball is ending.
- The `game.tilted` placeholder is true while tilted — useful to skip a bonus count.

## See also

[Modes](../concepts/modes.md) · [Installation](../getting-started/installation.md)
