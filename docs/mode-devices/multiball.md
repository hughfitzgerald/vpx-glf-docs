# Multiball

Puts additional balls into play and manages the shoot-again period, ending when the count returns
to one.

```vbscript
With .Multiballs("main")
    .BallCount = 3
    .BallCountType = "total"
    .BallLocks = Array("kicker1", "kicker2")
    .StartEvents = Array("start_main_multiball")
    .ShootAgain = 15000
    .HurryUp = 5000
End With
```

## Configuration

| Property | Default | Meaning |
|---|---|---|
| `BallCount` | 0 | Balls for the multiball |
| `BallCountType` | `"total"` | `"total"` = end with this many in play; `"add"` = add this many |
| `BallLocks` | — | Ball devices to eject balls from first |
| `StartEvents` | — | Events that start it |
| `StopEvents` | — | Events that stop it early |
| `ResetEvents` | — | Events that reset it so it can run again |
| `EnableEvents` / `DisableEvents` | — | Enable/disable; auto-enables if no enable events |
| `ShootAgain` | 10000 | Ball save duration in ms; 0 disables the save |
| `GracePeriod` | 0 | Extra ms after the save that still count |
| `HurryUp` | 0 | Warning this many ms before the save ends |
| `ReplaceBallsInPlay` | `False` | Replace balls already in play rather than adding |
| `AddABallEvents` | — | Events that add one ball mid-multiball |
| `AddABallShootAgain` | 5000 | Ball save for an added ball |
| `AddABallGracePeriod` | 0 | Grace period for an added ball |
| `AddABallHurryUpTime` | 0 | Hurry-up warning for an added ball |

`BallCount` and the timing values accept [placeholders](../concepts/placeholders.md).

## Events posted

| Event | Kwargs | When |
|---|---|---|
| `multiball_<name>_started` | `balls` | Multiball began |
| `multiball_<name>_ended` | | Back to one ball |
| `multiball_<name>_ball_lost` | | A ball drained, others remain |
| `multiball_<name>_shoot_again` | | Shoot-again period began |
| `multiball_<name>_hurry_up` | | Save about to expire |
| `multiball_<name>_grace_period` | | Save in its grace window |
| `multiball_<name>_shoot_again_ended` | | Save finished |
| `multiball_<name>_reset_event` | | Reset for another run |

Also `ball_save_<name>_timer_start` and `ball_save_<name>_add_a_ball_timer_start`, matching the
[ball save](ball-save.md) naming so the same shoot-again light logic works for both.

## Driving a mode from it

Multiball is usually paired with a mode carrying its rules:

```vbscript
With CreateGlfMode("multiball", 500)
    .StartEvents = Array("multiball_main_started")
    .StopEvents  = Array("multiball_main_ended", "ball_will_end")
    ' jackpots, lighting, music…
End With
```

## Ball counts

`BallCountType = "total"` with `BallCount = 3` means three balls in play regardless of how many
were live — starting from one ball adds two. `"add"` always adds `BallCount` more, which is what
you want for a multiball that can stack on top of another.

## Locks

`BallLocks` names [ball devices](../devices/ball-device.md) to empty when multiball starts, so
physically locked balls are the ones released. Any shortfall is served from the trough.

```vbscript
With .Multiballs("main")
    .BallCount = 3
    .BallLocks = Array("kicker1", "kicker2")
    .StartEvents = Array("lock_complete")
End With
```

See [Multiball locks](multiball-locks.md) for tracking the locks that lead here.

## Add-a-ball

```vbscript
With .Multiballs("main")
    .BallCount = 3
    .StartEvents = Array("start_main_multiball")
    .AddABallEvents = Array("add_a_ball_awarded")
    .AddABallShootAgain = 5000
End With
```

## Resetting

A multiball won't start again until reset:

```vbscript
.ResetEvents = Array("ball_started")
```

Without a reset event, it runs once per game. Resetting on `ball_started` gives once per ball.

## Notes

- Starting is refused while balls from a previous run are still in play.
- `ShootAgain = 0` disables the save, and the multiball ends the moment a ball drains.
- The multiball ends when balls in play return to one — this posts `_ended`, which is what should
  stop the associated mode.

## See also

[Multiball locks](multiball-locks.md) · [Ball save](ball-save.md) ·
[Ball device](../devices/ball-device.md)
