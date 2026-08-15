# Variable player

Writes player and machine variables in response to events. Scoring lives here.

```vbscript
With .VariablePlayer()
    With .EventName("score_5000")
        With .Variable("score")
            .Action = "add"
            .Int = 5000
        End With
    End With
End With
```

Structure: an `EventName` block per triggering event, containing a `Variable` block per variable
to write.

## Actions

| Action | Effect |
|---|---|
| `add` | Add to a player variable |
| `set` | Set a player variable |
| `add_machine` | Add to a machine variable |
| `set_machine` | Set a machine variable |

Default is `add`. Player and machine variables use different actions — mixing them up writes to
the wrong store silently, which is a recurring source of "my condition never fires".

## Value types

Set exactly one of these; it also declares the type:

```vbscript
.Int = 5000
.Float = 1.5
.String = "READY"
```

Any of them accepts a [placeholder](../concepts/placeholders.md):

```vbscript
.Int = "{current_player.jackpot_value}"
.Int = "{current_player.base_value * current_player.multiplier}"
.Int = "{50000 if current_player.mode_level > 2 else 10000}"
.Int = "{device.timers.hurry_up.ticks * 1000}"
```

That last form — scoring from a running timer — is how hurry-up awards are built.

## Several variables per event

```vbscript
With .EventName("ramp_completed")
    With .Variable("score")
        .Action = "add"
        .Int = "{25000 * current_player.scoring_multiplier}"
    End With
    With .Variable("ramps_made")
        .Action = "add"
        .Int = 1
    End With
    With .Variable("last_shot")
        .Action = "set"
        .String = "ramp"
    End With
End With
```

## Conditions

```vbscript
With .EventName("s_Spinner_active{modes.super_spinner.active == True}")
    With .Variable("score")
        .Action = "add"
        .Int = 25000
    End With
End With
```

## Machine variables

```vbscript
With .EventName("lock_ball")
    With .Variable("bottom_ball_locked")
        .Action = "set_machine"
        .Int = 1
    End With
End With
```

The variable must have been declared with `CreateMachineVar` — see
[Variables](../concepts/variables.md).

## A scoring pattern worth copying

Rather than scattering point values through the ruleset, define one event per award and let
everything else post it:

```vbscript
With .VariablePlayer()
    With .EventName("score_5000")
        With .Variable("score")
            .Action = "add"
            .Int = "{5000 * current_player.scoring_multiplier}"
        End With
    End With
    With .EventName("score_25000")
        With .Variable("score")
            .Action = "add"
            .Int = "{25000 * current_player.scoring_multiplier}"
        End With
    End With
End With
```

Now every rule posts `score_5000`, the multiplier applies everywhere automatically, and rebalancing
is one edit.

## Notes

- Writing the value a variable already holds is a no-op — it won't re-fire change listeners.
- Undeclared player variables read as `False`, which behaves as `0` in arithmetic.

## See also

[Variables](../concepts/variables.md) · [Placeholders](../concepts/placeholders.md)
