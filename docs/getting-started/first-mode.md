# Your first mode

This walks through a complete, working rule set: a spinner that scores, ramps up while a timer
runs, lights an insert when you complete three of them, and awards a bonus. Every concept used
here has a reference page; the point of this page is to show how they fit together.

---

## 1. A mode is a container

```vbscript
Sub CreateSpinnerMode()

    With CreateGlfMode("spinner", 200)
        .StartEvents = Array("ball_started")
        .StopEvents  = Array("ball_will_end")
    End With

End Sub
```

`CreateGlfMode(name, priority)` returns a mode. It does nothing until one of its `StartEvents`
fires, and it tears itself down when a `StopEvent` fires.

Priority does two jobs: it orders event handlers (higher runs first), and it becomes the base
priority for light shows the mode plays, so a higher-priority mode's lighting wins. Pick numbers
with gaps — 100, 200, 500 — so you can slot things in later.

Starting the mode posts `mode_spinner_started`; stopping posts `mode_spinner_stopped`. You'll use
those constantly.

## 2. React to a switch

Every switch in `glf_switches` posts `<name>_active` when hit. The **event player** turns one
event into others:

```vbscript
With .EventPlayer()
    .Add "s_Spinner_active", Array("spinner_hit")
End With
```

Why bother with an intermediate `spinner_hit` event rather than acting on `s_Spinner_active`
directly? Because it names the *rule*, not the *hardware*. Later you can post `spinner_hit` from a
second spinner, or from a test key, and everything downstream still works.

## 3. Score it

The **variable player** writes player variables. `score` is the built-in one:

```vbscript
With .VariablePlayer()
    With .EventName("spinner_hit")
        With .Variable("score")
            .Action = "add"
            .Int = 1000
        End With
    End With
End With
```

## 4. Make the value dynamic

Hard-coding 1000 is fine until you want the spinner to be worth more during a mode. Any value in
GLF can be a **placeholder** — a `{}` expression evaluated when the event fires:

```vbscript
Glf_SetInitialPlayerVar "spinner_value", 1000     ' in ConfigureGlfDevices

With .Variable("score")
    .Action = "add"
    .Int = "{current_player.spinner_value}"
End With
```

Now raising the spinner's worth is just another variable write. See
[Placeholders](../concepts/placeholders.md) for the full syntax, including arithmetic and
conditionals.

## 5. Add a timer

```vbscript
With .Timers("spinner_boost")
    .StartValue = 10
    .EndValue = 0
    .Direction = "down"
    .TickInterval = 1000
    With .ControlEvents()
        .EventName = "spinner_hit{current_player.spinner_value == 1000}"
        .Action = "start"
    End With
End With
```

Two things to notice.

The condition in braces is attached to the *listening* event: this control event only fires when
the player variable is still at its base value, so hitting the spinner again mid-boost won't
restart the timer.

The timer posts `timer_spinner_boost_started`, `timer_spinner_boost_tick` and
`timer_spinner_boost_complete`. Wire those up like any other event:

```vbscript
With .VariablePlayer()
    With .EventName("timer_spinner_boost_started")
        With .Variable("spinner_value")
            .Action = "set"
            .Int = 5000
        End With
    End With
    With .EventName("timer_spinner_boost_complete")
        With .Variable("spinner_value")
            .Action = "set"
            .Int = 1000
        End With
    End With
End With
```

## 6. Light an insert

Shots are the standard way to track "has the player done this thing yet" *and* drive the insert
that shows it. A shot has a **profile**: a list of named states, each with a light show.

```vbscript
' In ConfigureGlfDevices, before any mode that uses it:
With GlfShotProfiles("off_on")
    With .States("unlit")
        .Show = "off"
        .Key = "key_off_on_unlit"
    End With
    With .States("lit")
        .Show = "led_color"
        .Key = "key_off_on_lit"
    End With
End With
```

Then in the mode:

```vbscript
With .Shots("spinner_target")
    .Profile = "off_on"
    .Switch = "s_Spinner"
    With .Tokens()
        .Add "lights", "L12"
        .Add "color", "ff8800"
    End With
End With
```

The shot starts in state 0 (`unlit`) and advances one state per hit, playing that state's show.
`lights` and `color` are **tokens** — placeholders inside the show definition that the shot fills
in, which is what makes one profile reusable across a dozen inserts.

Hitting it posts `spinner_target_hit`, plus `spinner_target_unlit_hit` (named for the state it was
in *before* advancing), so you can score differently depending on state.

## 7. Group the shots

```vbscript
With .ShotGroups("all_spinners")
    .Shots = Array("spinner_target", "left_orbit_target", "right_orbit_target")
End With
```

When every shot in the group reaches the same state, the group posts `all_spinners_complete` with
a `state` kwarg, and `all_spinners_lit_complete`. That's your "completed the set" hook:

```vbscript
With .EventPlayer()
    .Add "all_spinners_lit_complete", Array("award_spinner_bonus", "play_sfx_jackpot")
End With
```

## 8. Put it together

```vbscript
Sub CreateSpinnerMode()

    With CreateGlfMode("spinner", 200)
        .StartEvents = Array("ball_started")
        .StopEvents  = Array("ball_will_end")

        With .EventPlayer()
            .Add "s_Spinner_active", Array("spinner_hit")
            .Add "all_spinners_lit_complete", Array("award_spinner_bonus")
        End With

        With .VariablePlayer()
            With .EventName("spinner_hit")
                With .Variable("score")
                    .Action = "add"
                    .Int = "{current_player.spinner_value}"
                End With
            End With
            With .EventName("timer_spinner_boost_started")
                With .Variable("spinner_value")
                    .Action = "set"
                    .Int = 5000
                End With
            End With
            With .EventName("timer_spinner_boost_complete")
                With .Variable("spinner_value")
                    .Action = "set"
                    .Int = 1000
                End With
            End With
            With .EventName("award_spinner_bonus")
                With .Variable("score")
                    .Action = "add"
                    .Int = 250000
                End With
            End With
        End With

        With .Timers("spinner_boost")
            .StartValue = 10
            .EndValue = 0
            .Direction = "down"
            .TickInterval = 1000
            With .ControlEvents()
                .EventName = "spinner_hit{current_player.spinner_value == 1000}"
                .Action = "start"
            End With
        End With

        With .Shots("spinner_target")
            .Profile = "off_on"
            .Switch = "s_Spinner"
            With .Tokens()
                .Add "lights", "L12"
                .Add "color", "ff8800"
            End With
        End With

        With .ShotGroups("all_spinners")
            .Shots = Array("spinner_target")
        End With

    End With

End Sub
```

Call `CreateSpinnerMode()` from `ConfigureGlfDevices`, after your shows and shot profiles.

---

## The mental model

Nothing in that config calls anything else directly. Each piece announces what happened and
listens for what it cares about. That indirection is the whole point: to change when the boost
triggers, you edit one event string, and nothing downstream needs to know.

Two habits worth forming early:

- **Name events after intent** (`spinner_hit`, `award_spinner_bonus`), not after mechanism.
- **Let modes own lifetime.** Anything declared inside a mode is registered on start and removed
  on stop. If a rule should only apply during multiball, put it in the multiball mode rather than
  guarding it with conditions in the base mode.

---

## Next

[Project layout →](project-layout.md) · [The event system →](../concepts/events.md)
