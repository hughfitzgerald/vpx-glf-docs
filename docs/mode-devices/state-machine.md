# State machine

Models progression through named states with explicit transitions. Where a [shot](shot.md)
advances linearly on hits, a state machine lets any state move to any other, driven by different
events.

```vbscript
With .StateMachines("mode_progress")
    .PersistState = True
    .StartingState = "unqualified"

    With .States("unqualified")
        With .ShowWhenActive()
            .Show = "off"
            .Key = "key_progress_unqualified"
            With .Tokens()
                .Add "lights", "L20"
            End With
        End With
    End With

    With .States("qualified")
        .EventsWhenStarted = Array("play_voc_mode_ready")
        With .ShowWhenActive()
            .Show = "flash_color"
            .Key = "key_progress_qualified"
            .Speed = 4
            With .Tokens()
                .Add "lights", "L20"
                .Add "color", "00ff00"
            End With
        End With
    End With

    With .States("running")
        With .ShowWhenActive()
            .Show = "led_color"
            .Key = "key_progress_running"
            With .Tokens()
                .Add "lights", "L20"
                .Add "color", "ffffff"
            End With
        End With
    End With

    With .Transitions()
        .Source = Array("unqualified")
        .Target = "qualified"
        .Events = Array("targets_complete")
    End With

    With .Transitions()
        .Source = Array("qualified")
        .Target = "running"
        .Events = Array("s_Scoop_active")
        .EventsWhenTransitioning = Array("start_mode")
    End With

    With .Transitions()
        .Source = Array("running")
        .Target = "unqualified"
        .Events = Array("mode_ended")
    End With
End With
```

## Machine properties

| Property | Default | Meaning |
|---|---|---|
| `StartingState` | `"start"` | State entered when the machine starts |
| `PersistState` | `False` | Keep state across balls, per player |
| `EnableEvents` / `DisableEvents` | — | Enable/disable the machine |

Default `StartingState` is `"start"`, so either name your first state `start` or set this — a
machine whose starting state doesn't exist simply won't start.

## States

| Property | Meaning |
|---|---|
| `ShowWhenActive` | A show played while in this state |
| `EventsWhenStarted` | Events posted on entering |
| `EventsWhenStopped` | Events posted on leaving |
| `Label` | Display label |

`ShowWhenActive` is configured like a [show player](../players/show-player.md) entry — `Show`,
`Key`, `Speed`, `Loops`, `Priority`, `Tokens`. The show starts on entry and stops on exit, so the
insert always matches the state without any extra wiring.

## Transitions

| Property | Meaning |
|---|---|
| `Source` | Array of states this transition can fire from |
| `Target` | State to move to |
| `Events` | Events that trigger it |
| `EventsWhenTransitioning` | Events posted during the transition |

Each `With .Transitions()` creates one. A transition only fires when the machine is in one of its
`Source` states, so the same event can mean different things depending on where you are — the whole
point of using a state machine rather than a pile of conditions.

Several sources let one transition serve as a catch-all:

```vbscript
With .Transitions()
    .Source = Array("qualified", "running", "hurry_up")
    .Target = "unqualified"
    .Events = Array("ball_will_end")
End With
```

## Reading the state

```vbscript
.Add "s_Scoop_active{device.state_machines.mode_progress.state == ""qualified""}", Array("start_mode")
```

Note the doubled quotes — VBScript string escaping inside an already-quoted event string.

With `PersistState = True` it's also in the player variable `state_machine_<name>`.

## When to reach for one

Use a state machine when the *same input means different things depending on progress* — a scoop
that qualifies, then starts, then collects. Use a [shot](shot.md) when progress is linear and each
hit advances one step; use a [counter](counter.md) when only a total matters.

State machines also give you `EventsWhenStarted` per state, which is a clean place to hang callouts
and display changes that belong to a phase rather than to an action.

## See also

[Shot](shot.md) · [Counter](counter.md) · [Modes](../concepts/modes.md)
