# Multiball locks

Counts balls locked toward a multiball, physically or virtually.

```vbscript
With .MultiballLocks("main_locks")
    .LockDevices = Array("kicker1", "kicker2")
    .BallsToLock = 3
    .ResetEvents = Array("multiball_main_started")
End With
```

## Configuration

| Property | Default | Meaning |
|---|---|---|
| `LockDevices` | — | Ball devices whose entering balls count as locks |
| `BallsToLock` | 0 | Locks needed before the lock is full |
| `LockEvents` | — | Events that count as a virtual lock, without a device |
| `ResetEvents` | — | Events that clear the count |
| `BallsToReplace` | -1 | How many locked balls to replace from the trough; -1 for all |
| `LockedBallCountingStrategy` | `"virtual_only"` | How locks are counted |
| `EnableEvents` / `DisableEvents` | — | Enable/disable; auto-enables if no enable events |

## Events posted

| Event | Kwargs | When |
|---|---|---|
| `multiball_lock_<name>_locked_ball` | `balls_locked` | A ball was locked |
| `multiball_lock_<name>_full` | `balls_locked` | `BallsToLock` reached |

`multiball_lock_<name>_full` is the natural trigger for starting a multiball:

```vbscript
With .Multiballs("main")
    .BallCount = 3
    .BallLocks = Array("kicker1", "kicker2")
    .StartEvents = Array("multiball_lock_main_locks_full")
End With
```

Current state is readable as `device.multiball_locks.<name>.locked_balls` and `.enabled`.

## Physical versus virtual locks

**Physical** — balls sit in a device until multiball:

```vbscript
With .MultiballLocks("main_locks")
    .LockDevices = Array("lock_saucer")
    .BallsToLock = 3
End With
```

Each ball entering `lock_saucer` counts, and a replacement is served from the trough so play
continues.

**Virtual** — the ball is returned immediately and only the count is kept:

```vbscript
With .MultiballLocks("main_locks")
    .LockEvents = Array("lock_shot_made")
    .BallsToLock = 3
End With
```

Virtual locks avoid needing a saucer that holds several balls, and are far easier to get right.
Most tables use them unless the physical lock is part of the appeal.

## Per-player locks

Lock counts are per player, but a physical saucer is not — a ball sitting in it belongs to
whoever put it there. The usual approach is a machine variable tracking occupancy alongside the
per-player lock count:

```vbscript
With CreateMachineVar("bottom_ball_locked")
    .InitialValue = 0
    .ValueType = "int"
    .Persist = False
End With
```

```vbscript
.Add "s_Lock_active{machine.bottom_ball_locked == 0}", Array("lock_ball")
```

See [Variables](../concepts/variables.md).

## Resetting

```vbscript
.ResetEvents = Array("multiball_main_started")
```

Clearing the count when multiball starts lets the player begin qualifying the next one immediately.

## See also

[Multiball](multiball.md) · [Ball device](../devices/ball-device.md) · [Ball hold](ball-hold.md)
