# VPX Game Logic Framework (GLF) — Usage Documentation

GLF is a VBScript framework for building **original** Visual Pinball X tables. Instead of writing
game logic by hand, you declare it: modes, shots, timers, multiballs, light shows and displays are
all configured as objects that listen for and post events.

The design borrows heavily from the [Mission Pinball Framework](https://missionpinball.org) (MPF).
If you know MPF's YAML config, most concepts map across directly — but GLF runs entirely inside
VPX with no external process required.

> **Version.** This documentation describes the `main` branch of
> [mpcarr/vpx-glf](https://github.com/mpcarr/vpx-glf) as of **April 14, 2026**. Where the published
> example tables use older patterns, the differences are called out in
> [Migrating older tables](appendix/migrating.md).

**CAUTION:** This documentation project is AI-generated and unrelated to the development of GLF. Huge
thanks to the creators of GLF and all developers and creators in the VPX community. There may be errors
in this documentation. If there are, please let me know and I can fix them.

---

## Start here

| | |
|---|---|
| [Installation](getting-started/installation.md) | Table objects, collections and script hooks GLF needs |
| [Your first mode](getting-started/first-mode.md) | A complete worked example, from switch hit to score |
| [Project layout](getting-started/project-layout.md) | Splitting your script into files, build workflow |

## Core concepts

Read these before the reference sections — nearly every device is configured in terms of them.

| | |
|---|---|
| [The event system](concepts/events.md) | How events are posted, ordered, conditioned and consumed |
| [Placeholders & dynamic values](concepts/placeholders.md) | `current_player.score`, `device.timers.x.ticks`, formatting |
| [Modes](concepts/modes.md) | Lifecycle, priority, start/stop events, mode devices |
| [Variables](concepts/variables.md) | Player variables vs machine variables, persistence |
| [Lights & shows](concepts/lights-and-shows.md) | Light naming, tags, show authoring, tokens, the light stack |

## Reference

**[Machine devices](devices/)** — created once at startup, outside any mode:

[Ball device](devices/ball-device.md) ·
[Trough](devices/trough.md) ·
[Flipper](devices/flipper.md) ·
[Auto-fire device](devices/auto-fire.md) ·
[Diverter](devices/diverter.md) ·
[Targets](devices/targets.md) ·
[Magnet](devices/magnet.md) ·
[Segment display](devices/segment-display.md) ·
[Sound system](devices/sound-system.md) ·
[Ball search](devices/ball-search.md)

**[Mode devices](mode-devices/)** — created inside a mode, live only while it runs:

[Shot](mode-devices/shot.md) ·
[Shot profile](mode-devices/shot-profile.md) ·
[Shot group](mode-devices/shot-group.md) ·
[Sequence shot](mode-devices/sequence-shot.md) ·
[Timer](mode-devices/timer.md) ·
[Counter](mode-devices/counter.md) ·
[Ball save](mode-devices/ball-save.md) ·
[Ball hold](mode-devices/ball-hold.md) ·
[Multiball](mode-devices/multiball.md) ·
[Multiball locks](mode-devices/multiball-locks.md) ·
[Extra ball](mode-devices/extra-ball.md) ·
[Combo switches](mode-devices/combo-switches.md) ·
[Timed switches](mode-devices/timed-switches.md) ·
[State machine](mode-devices/state-machine.md) ·
[Tilt](mode-devices/tilt.md) ·
[High score](mode-devices/high-score.md)

**[Players](players/)** — a mode's outputs; each maps events to an action:

[Event player](players/event-player.md) ·
[Queue event player](players/queue-event-player.md) ·
[Queue relay player](players/queue-relay-player.md) ·
[Random event player](players/random-event-player.md) ·
[Variable player](players/variable-player.md) ·
[Light player](players/light-player.md) ·
[Show player](players/show-player.md) ·
[Segment display player](players/segment-display-player.md) ·
[Sound player](players/sound-player.md) ·
[DOF player](players/dof-player.md) ·
[Slide & widget player](players/slide-and-widget-player.md)

**Appendix**

| | |
|---|---|
| [Event reference](appendix/event-reference.md) | Every event GLF posts, in one table |
| [Debugging](appendix/debugging.md) | Logs, the monitor, table options |
| [Migrating older tables](appendix/migrating.md) | Changes since the published example tables |

---

## The shape of a GLF table

Everything hangs off two ideas: **events** and **modes**.

A switch hit posts an event. Modes that are running listen for events and react — by scoring,
starting a show, advancing a shot, posting further events of their own. When a mode stops, every
listener it registered is torn down automatically, so logic can't leak between modes.

```vbscript
Sub ConfigureGlfDevices()

    ' Machine devices: exist for the whole session
    With CreateGlfBallDevice("plunger")
        .BallSwitches = Array("s_Plunger1")
        .MechanicalEject = True
        .DefaultDevice = True
        .EjectCallback = "PlungerEjectCallback"
    End With

    ' Modes: created here, started and stopped by events
    With CreateGlfMode("base", 110)
        .StartEvents = Array("ball_started")
        .StopEvents  = Array("ball_ended")

        With .EventPlayer()
            .Add "s_LeftSlingshot_active", Array("score_5000")
        End With

        With .VariablePlayer()
            With .EventName("score_5000")
                With .Variable("score")
                    .Action = "add"
                    .Int = 5000
                End With
            End With
        End With
    End With

End Sub
```

That is a complete, working scoring rule. The rest of this documentation is that idea, elaborated.
