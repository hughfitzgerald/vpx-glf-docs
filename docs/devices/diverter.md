# Diverter

A coil that switches a ball's path, or fires it back into play — gates, magnets on ramps, and
kickbacks.

```vbscript
With CreateGlfDiverter("diverter1")
    .EnableEvents     = Array("ball_started", "reset_complete", "enable_diverter")
    .ActivateEvents   = Array("open_diverter", "game_ending")
    .DeactivateEvents = Array("close_diverter", "ball_ended")
    .ActionCallback   = "OpenDiverter"
End With
```

## Configuration

| Property | Default | Meaning |
|---|---|---|
| `ActionCallback` | — | Function that drives the coil |
| `EnableEvents` | — | Events that enable the device |
| `DisableEvents` | — | Events that disable it |
| `ActivateEvents` | — | Events that activate the diverter |
| `DeactivateEvents` | — | Events that deactivate it |
| `ActivationSwitches` | — | Switches that activate it directly |
| `ActivationTime` | 0 | Ms to stay active before auto-deactivating; 0 stays until told |
| `BallSearchHoldTime` | 1000 | Ms to hold during ball search |
| `ExcludeFromBallSearch` | `False` | Skip during ball search |

Enable and activate are separate: **enabled** means the device will respond at all, **active**
means the coil is currently on. A diverter is typically enabled for the whole ball and activated
only when needed.

## The action callback

Receives 1 to activate, 0 to deactivate:

```vbscript
Sub OpenDiverter(enabled)
    If enabled = 1 Then
        DiverterWall.IsDropped = True
        SolDiverter 1
    Else
        DiverterWall.IsDropped = False
        SolDiverter 0
    End If
End Sub
```

## Kickbacks

A kickback is a diverter that fires momentarily rather than holding a path open:

```vbscript
With CreateGlfDiverter("kickback1")
    .EnableEvents     = Array("ball_started", "reset_complete", "enable_kickback")
    .ActivateEvents   = Array("fire_kickback")
    .DeactivateEvents = Array("reset_kickback", "ball_ended")
    .ActionCallback   = "SolKickback"
End With
```

With the outlane switch posting the fire event, conditioned on the kickback being lit:

```vbscript
With .EventPlayer()
    .Add "s_LeftOutlane_active{current_player.kickback_lit == 1}", Array("fire_kickback")
End With
```

`ActivationSwitches` can wire the switch straight to the coil instead, which is faster but skips
any conditions — use it only when the diverter should always fire.

## Events posted

| Event | When |
|---|---|
| `diverter_<name>_activating` | Activating |
| `diverter_<name>_deactivating` | Deactivating |

State is readable as `device.diverters.<name>.enabled` and `.active`.

## Timed activation

```vbscript
With CreateGlfDiverter("ramp_gate")
    .EnableEvents   = Array("ball_started")
    .ActivateEvents = Array("open_ramp")
    .ActivationTime = 3000
    .ActionCallback = "RampGateAction"
End With
```

The gate closes itself after three seconds, so no deactivate event is needed.

## See also

[Auto-fire device](auto-fire.md) · [Magnet](magnet.md)
