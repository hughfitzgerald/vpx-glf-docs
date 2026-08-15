# Auto-fire device

A coil that fires automatically when its switch closes: slingshots and pop bumpers. Same shape as a
[flipper](flipper.md), but the response is automatic rather than player-driven.

```vbscript
With CreateGlfAutoFireDevice("left_sling")
    .Switch = "s_LeftSlingshot"
    .ActionCallback = "LeftSlingshotAction"
    .EnabledCallback = "LeftSlingshotEnabled"
    .DisabledCallback = "LeftSlingshotDisabled"
    .EnableEvents  = Array("ball_started", "enable_flippers")
    .DisableEvents = Array("kill_flippers")
End With
```

## Configuration

| Property | Default | Meaning |
|---|---|---|
| `Switch` | — | Switch that triggers the coil |
| `ActionCallback` | — | Function that fires the coil |
| `EnabledCallback` | — | Called when the device is enabled |
| `DisabledCallback` | — | Called when disabled |
| `EnableEvents` | `ball_started` | Events that enable it |
| `DisableEvents` | `ball_will_end`, `service_mode_entered` | Events that disable it |
| `ExcludeFromBallSearch` | `False` | Skip during ball search |

## Collections

This is the fiddly part, and it differs by device type:

- **Slingshots** go in the `glf_slingshots` collection.
- **Pop bumper switches** go in the `glf_switches` collection.

Both are then also declared as auto-fire devices. The collection is what generates the switch
event; the device is what fires the coil and gates it on game state.

## The action callback

Receives `Array(state, activeBall)` — 1 on activation, 0 on release:

```vbscript
Sub LeftSlingshotAction(args)
    Dim enabled : enabled = args(0)
    If enabled = 1 Then
        LSling.VelocityCorrect args(1)
        LeftSlingShot.TransZ = -20
        RandomSoundSlingshotLeft SLING1
        DOF 101, DOFPulse
    Else
        LeftSlingShot.TransZ = 0
    End If
End Sub
```

The enabled/disabled callbacks are for visual state — dimming a bumper light when the device is
off:

```vbscript
Sub Bumper1Enabled(args)
    Bumper1Light.State = 1
End Sub
```

## Events posted

| Event | When |
|---|---|
| `auto_fire_coil_<name>_activate` | Coil fired |
| `auto_fire_coil_<name>_deactivate` | Released |

For scoring, listen to the switch event — it's shorter and reads better:

```vbscript
With .EventPlayer()
    .Add "s_Bumper1_active", Array("score_5000", "play_bumper1_show")
End With
```

## Ball search

Auto-fire devices are pulsed during [ball search](ball-search.md) to dislodge a stuck ball. Set
`ExcludeFromBallSearch = True` for any that would be unhelpful or noisy.

## See also

[Flipper](flipper.md) · [Ball search](ball-search.md)
