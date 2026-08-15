# Counter

Counts occurrences of an event and posts events when a target is reached. Simpler than a
[shot](shot.md) when there's no insert to drive and no per-state behaviour.

```vbscript
With .Counters("ramps_for_award")
    .CountEvents = Array("ramp_made")
    .CountCompleteValue = 3
    .EventsWhenComplete = Array("ramp_award_ready")
    .ResetOnComplete = True
End With
```

## Configuration

| Property | Default | Meaning |
|---|---|---|
| `CountEvents` | — | Events that increment the count |
| `CountCompleteValue` | — | Count at which completion fires |
| `EventsWhenComplete` | — | Plain array of event names posted on completion |
| `EnableEvents` | — | Events that enable counting |
| `DisableOnComplete` | `False` | Stop counting after completion |
| `ResetOnComplete` | `False` | Return to zero after completion |
| `PersistState` | `False` | Keep the count across balls, per player |

With no `EnableEvents`, the counter starts counting when the mode starts.

`EventsWhenComplete` takes bare event names — unlike most event lists, these are not parsed for
conditions.

## Reading the count

With `PersistState = True` the count lives in the player variable `counter_<name>_state`:

```vbscript
.Add "s_Scoop_active{current_player.counter_ramps_for_award_state >= 2}", Array("almost_there")
```

Without persistence the count resets when the mode stops, and isn't exposed as a variable.

## Repeating awards

`ResetOnComplete` plus a counter that stays enabled gives an award every N hits:

```vbscript
With .Counters("every_ten_pops")
    .CountEvents = Array("s_Bumper1_active", "s_Bumper2_active", "s_Bumper3_active")
    .CountCompleteValue = 10
    .EventsWhenComplete = Array("pop_bumper_award")
    .ResetOnComplete = True
End With
```

Several `CountEvents` all increment the same counter, which is the tidy way to count "any of these
targets".

## One-shot progress

```vbscript
With .Counters("locks_needed")
    .CountEvents = Array("ball_locked")
    .CountCompleteValue = 3
    .EventsWhenComplete = Array("multiball_ready")
    .DisableOnComplete = True
    .PersistState = True
End With
```

`PersistState` matters here — multiball progress should survive the ball ending.

## Counter or shot?

| Use a counter | Use a shot |
|---|---|
| No insert to light | An insert reflects progress |
| Only the total matters | Each state behaves differently |
| Counting several unrelated events | One target with a profile |

If you find yourself adding conditions on a counter's value to change behaviour per step, that's a
shot with a multi-state profile.

## See also

[Shot](shot.md) · [Timer](timer.md) · [Variables](../concepts/variables.md)
