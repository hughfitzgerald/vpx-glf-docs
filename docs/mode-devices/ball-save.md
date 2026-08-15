# Ball save

Returns a ball that drains within a grace window.

```vbscript
With .BallSaves("new_ball")
    .ActiveTime = 6000
    .HurryUpTime = 3000
    .GracePeriod = 2000
    .BallsToSave = 1
    .AutoLaunch = True
    .EnableEvents = Array("new_ball_active")
End With
```

## Configuration

| Property | Default | Meaning |
|---|---|---|
| `ActiveTime` | — | Milliseconds the save stays active |
| `GracePeriod` | 0 | Extra milliseconds after `ActiveTime` that still save |
| `HurryUpTime` | 0 | Post a warning this many ms *before* the save expires |
| `BallsToSave` | 1 | How many balls may be saved before it disables itself |
| `AutoLaunch` | `False` | Re-launch the saved ball automatically |
| `EnableEvents` | — | Events that arm the save |
| `DisableEvents` | — | Events that disarm it |
| `TimerStartEvents` | — | Events that start the countdown |

`ActiveTime`, `GracePeriod` and `HurryUpTime` accept
[placeholders](../concepts/placeholders.md), so the save can be longer on ball 1 or for a new
player.

## Arming versus timing

By default the countdown starts the moment the save is enabled. `TimerStartEvents` separates the
two:

```vbscript
With .BallSaves("new_ball")
    .ActiveTime = 8000
    .EnableEvents = Array("ball_started")
    .TimerStartEvents = Array("s_Plunger1_inactive")
End With
```

Now the save is armed at ball start but doesn't begin counting until the ball actually leaves the
plunger lane — so a player who takes their time isn't punished for it.

## Events posted

| Event | When |
|---|---|
| `ball_save_<name>_enabled` | Armed |
| `ball_save_<name>_timer_start` | Countdown started |
| `ball_save_<name>_hurry_up` | `HurryUpTime` ms before expiry |
| `ball_save_<name>_grace_period` | `ActiveTime` elapsed; grace period running |
| `ball_save_<name>_saving_ball` | A ball was saved |

## Driving the shoot-again light

The standard pattern — a shot whose states track the save's phases:

```vbscript
With .Shots("base_shoot_again")
    .Profile = "shoot_again"
    With .Tokens()
        .Add "color", ShootAgainColor
    End With
    With .ControlEvents()
        .Events = Array("ball_save_new_ball_enabled")
        .State = 1
    End With
    With .ControlEvents()
        .Events = Array("ball_save_new_ball_hurry_up")
        .State = 2
    End With
    .RestartEvents = Array("ball_save_new_ball_grace_period", "ball_save_new_ball_saving_ball")
End With
```

State 1 is a steady flash, state 2 a faster one, and the restart events return it to unlit when the
save ends. See [Shot profile](shot-profile.md) for the matching `shoot_again` profile.

Announce a save with a widget or sound off `ball_save_<name>_saving_ball`:

```vbscript
With .WidgetPlayer()
    With .EventName("ball_save_new_ball_saving_ball")
        .Widget = "ball_save"
        .Action = "play"
        .Expire = 2
    End With
End With
```

## Grace period

`GracePeriod` extends the save past its visible expiry. The shoot-again light goes out at
`ActiveTime`, but a drain during the grace window is still saved — which feels fair rather than
generous, because the player believes the save had ended.

## Multiple saves

A mode can hold several ball saves with different lifetimes:

```vbscript
With .BallSaves("new_ball")
    .ActiveTime = 6000
    .EnableEvents = Array("ball_started")
End With

With .BallSaves("multiball_start")
    .ActiveTime = 15000
    .BallsToSave = 3
    .EnableEvents = Array("multiball_main_started")
End With
```

Multiball has its own built-in save via `ShootAgain` — see [Multiball](multiball.md) — so a
separate device is only needed for finer control.

## Notes

- `AutoLaunch = True` ejects the saved ball from the plunger without the player plunging.
- The save disables itself once `BallsToSave` balls have been returned.
- Saves belong to their mode and are torn down when it stops.

## See also

[Multiball](multiball.md) · [Shot](shot.md)
