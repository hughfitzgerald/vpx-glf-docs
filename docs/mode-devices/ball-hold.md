# Ball hold

Holds balls in a device and releases them on cue. Unlike [multiball locks](multiball-locks.md),
which count toward a goal, a ball hold is about physically parking balls — a captive ball trap, a
holdover during a mode intro, or staging balls for a wizard mode.

```vbscript
With .BallHolds("mode_hold")
    .HoldDevices = Array("kicker1")
    .BallsToHold = 2
    .ReleaseAllEvents = Array("mode_wizard_started")
End With
```

## Configuration

| Property | Default | Meaning |
|---|---|---|
| `HoldDevices` | — | Ball devices that hold balls |
| `BallsToHold` | 0 | Maximum balls to hold |
| `ReleaseAllEvents` | `Array("tilt")` | Release everything |
| `ReleaseOneEvents` | — | Release a single ball |
| `ReleaseOneIfFullEvents` | — | Release one, but only when the hold is full |
| `EnableEvents` / `DisableEvents` | — | Enable/disable; auto-enables if no enable events |

Note the default on `ReleaseAllEvents`: a hold releases on `tilt` unless you override it. Setting
your own release events replaces that, so include `"tilt"` yourself if you want to keep the
behaviour:

```vbscript
.ReleaseAllEvents = Array("mode_wizard_started", "tilt", "ball_will_end")
```

## Events posted

| Event | Kwargs | When |
|---|---|---|
| `ball_hold_<name>_held_ball` | `balls_held` | A ball was captured |
| `ball_hold_<name>_full` | `balls_held` | `BallsToHold` reached |
| `ball_hold_<name>_balls_released` | | Balls released |

Readable as `device.ball_holds.<name>.balls_held`.

## Staging a wizard mode

Hold balls as the player qualifies, then release them all at once:

```vbscript
With .BallHolds("wizard_stage")
    .HoldDevices = Array("lock_saucer")
    .BallsToHold = 3
    .ReleaseAllEvents = Array("wizard_mode_started", "tilt")
End With

With .EventPlayer()
    .Add "ball_hold_wizard_stage_full", Array("wizard_mode_ready")
End With
```

## Release one at a time

`ReleaseOneIfFullEvents` gives a rotating hold — a new ball goes in, an old one comes out, keeping
one ball in play:

```vbscript
With .BallHolds("rotating_lock")
    .HoldDevices = Array("kicker1")
    .BallsToHold = 1
    .ReleaseOneIfFullEvents = Array("balldevice_kicker1_ball_entered")
End With
```

## Notes

- Disabling a hold releases everything it holds.
- Held balls are excluded from ball search's "where did the ball go" logic.
- Holding all balls in play with nothing to release them stalls the game — always give a hold a
  release path, including on `ball_will_end`.

## See also

[Multiball locks](multiball-locks.md) · [Ball device](../devices/ball-device.md)
