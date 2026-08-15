# Shot

A shot tracks progress at one target and drives the insert that shows it. It is the workhorse of
GLF rule-writing: most "hit this three times to light it" logic is a shot plus a
[shot profile](shot-profile.md).

```vbscript
With .Shots("left_ramp")
    .Profile = "off_on_color"
    .Switch = "s_LeftRamp"
    With .Tokens()
        .Add "lights", "L12"
        .Add "color", RampshotColor
    End With
End With
```

A shot has a **state** — an index into its profile's list of states. Hitting it advances the state
and plays that state's show. That's the whole idea; everything else is refinement.

## Configuration

| Property | Default | Meaning |
|---|---|---|
| `Profile` | `"default"` | Shot profile name |
| `Switch` | — | One switch that registers a hit |
| `Switches` | — | Array of switches, any of which registers a hit |
| `HitEvents` | — | Events that register a hit, as well as (or instead of) switches |
| `Tokens` | — | Token values passed to the profile's shows |
| `EnableEvents` | — | Events that enable the shot |
| `DisableEvents` | — | Events that disable it |
| `StartEnabled` | auto | Force enabled/disabled at mode start |
| `AdvanceEvents` | — | Advance one state without registering a hit |
| `ControlEvents` | — | Jump to a specific state |
| `ResetEvents` | — | Return to state 0 |
| `RestartEvents` | — | Reset and re-enable |
| `Persist` | `True` | Remember state across balls, per player |

If you set no `EnableEvents`, the shot enables itself when the mode starts. Set them when a shot
should only be live for part of the mode.

## Events posted

Hitting a shot named `left_ramp` with profile `off_on_color`, currently in state `unlit`, posts:

```
left_ramp_hit
left_ramp_off_on_color_hit
left_ramp_off_on_color_unlit_hit
left_ramp_unlit_hit
```

All four carry kwargs `profile`, `state` (the state *before* advancing) and `advancing`.

That fan-out is deliberate: score generically off `left_ramp_hit`, or differently per state off
`left_ramp_unlit_hit`.

```vbscript
With .EventPlayer()
    .Add "left_ramp_hit",     Array("score_25000")
    .Add "left_ramp_lit_hit", Array("collect_ramp_award")
End With
```

## Control events

Jump to a specific state, rather than advancing:

```vbscript
With .Shots("base_shoot_again")
    .Profile = "shoot_again"
    With .ControlEvents()
        .Events = Array("ball_save_new_ball_enabled")
        .State = 1
    End With
    With .ControlEvents()
        .Events = Array("ball_save_new_ball_hurry_up")
        .State = 2
    End With
    .RestartEvents = Array("ball_save_new_ball_grace_period")
End With
```

| Property | Default | Meaning |
|---|---|---|
| `Events` | — | Events that trigger the jump |
| `State` | 0 | State index to jump to |
| `Force` | `True` | Jump even while disabled |
| `ForceShow` | `False` | Replay the show even if already in that state |

Each `With .ControlEvents()` creates a new entry, so repeat the block for each state you want to
be able to jump to.

## State persistence

With `Persist = True` (the default), state is stored in the player variable `shot_<name>` and
survives across balls. Set `Persist = False` for shots that should start fresh each ball.

Because it's an ordinary player variable, you can read it:

```vbscript
.Add "s_Scoop_active{current_player.shot_left_ramp == 2}", Array("collect_super")
```

## Shots without switches

A shot can be driven purely by events, which is how you build progress indicators that aren't tied
to one target:

```vbscript
With .Shots("progress_indicator")
    .Profile = "three_stage"
    .HitEvents = Array("any_ramp_completed")
    With .Tokens()
        .Add "lights", "L30"
    End With
End With
```

## Blocking

When a shot fires on a switch it registers a block, so a second shot listening to the same switch
in the same mode won't also fire. A profile with `.Block = True` blocks by shot name rather than
shot-plus-profile. If two shots on one switch both need to fire, put them in different modes.

## See also

[Shot profile](shot-profile.md) · [Shot group](shot-group.md) ·
[Sequence shot](sequence-shot.md)
