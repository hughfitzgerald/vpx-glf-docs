# Random event player

Posts one event chosen at random from a weighted list. For mystery awards, rotating callouts, and
anything that should vary between plays.

```vbscript
With .RandomEventPlayer()
    With .EventName("mystery_awarded")
        .Add "award_points", 3
        .Add "award_extra_ball", 1
        .Add "award_light_lock", 2
    End With
End With
```

`.Add event, weight` — higher weights are picked more often. The example above picks points half
the time, lock a third, extra ball a sixth.

## Properties

| Property | Default | Meaning |
|---|---|---|
| `ForceAll` | `True` | Don't repeat an event until every option has been used |
| `ForceDifferent` | `True` | Never pick the same event twice in a row |
| `DisableRandom` | `False` | Pick in order instead of randomly — useful for testing |
| `FallbackEvent` | — | Posted when no option is available |
| `Scope` | `"player"` | Tracking scope for the above |

`ForceAll` is the one that makes this feel designed rather than random: a player cycles through
every award before any repeats, so ten mystery hits give ten different outcomes rather than the
same one three times.

## Conditional options

Each option is an event string, so it can carry a condition. Options whose condition is false are
excluded from the draw:

```vbscript
With .EventName("mystery_awarded")
    .Add "award_extra_ball{current_player.extra_balls_awarded < 1}", 1
    .Add "award_light_lock{current_player.locks_lit == 0}", 2
    .Add "award_points", 3
End With
```

If every option is excluded, `FallbackEvent` fires — worth setting so the player always gets
*something*:

```vbscript
With .EventName("mystery_awarded")
    .FallbackEvent = "award_points_consolation"
    .Add "award_extra_ball{current_player.extra_balls_awarded < 1}", 1
End With
```

## Notes

- Selection state is stored in player variables, so each player has their own cycle.
- Weights are relative, not percentages — they needn't sum to anything in particular.
- With `ForceDifferent = True` and only two options, the result alternates strictly.

## See also

[Event player](event-player.md)
