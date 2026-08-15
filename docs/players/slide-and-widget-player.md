# Slide & widget player

Drive an external media controller over BCP (Backbox Control Protocol) — typically a Godot Media
Controller project displaying on an LCD backglass or DMD.

Both players are inert unless BCP is connected (the *Glf Backbox Control Protocol* table option).
A table with no media controller can configure them harmlessly; nothing will happen.

## Slide player

A slide is a full-screen display state.

```vbscript
With .SlidePlayer()
    With .EventName("mode_base_started")
        .Slide = "base"
        .Action = "play"
    End With
End With
```

| Property | Default | Meaning |
|---|---|---|
| `Slide` | — | Slide name in the media controller project |
| `Action` | `"play"` | `play` or `remove` |
| `Expire` | 0 | Auto-remove after N seconds; 0 means stay |
| `Priority` | 0 | Added to the mode's priority |

## Widget player

A widget is an overlay element on top of the current slide.

```vbscript
With .WidgetPlayer()
    With .EventName("ball_save_new_ball_saving_ball")
        .Widget = "ball_save"
        .Action = "play"
        .Expire = 2
    End With
End With
```

Same properties, with `Widget` in place of `Slide`. `Expire` is the common one — most widgets are
transient notifications ("BALL SAVED", "JACKPOT") that should clear themselves.

## Priority and context

Slides and widgets are priority-stacked on the media controller side much as lights are locally,
using the mode's priority plus any offset. The mode name is sent as the display context, so the
controller can clear everything belonging to a mode when it stops.

## From a show step

Both can be triggered from inside a show step, keeping display changes in sync with lighting:

```vbscript
With .AddStep(Null, Null, 1)
    .Lights = Array("(lights)|100|ffffff")
    With .Slides("jackpot_award")
        .Expire = 2
    End With
End With
```

## See also

[Modes](../concepts/modes.md) · [High score](../mode-devices/high-score.md)
