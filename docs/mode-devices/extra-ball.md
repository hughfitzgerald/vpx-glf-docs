# Extra ball

Awards an extra ball, with a per-game cap.

```vbscript
With .ExtraBalls("main")
    .AwardEvents = Array("extra_ball_lit_collected")
    .MaxPerGame = 3
End With
```

## Configuration

| Property | Default | Meaning |
|---|---|---|
| `AwardEvents` | — | Events that award an extra ball |
| `MaxPerGame` | 0 | Cap per player per game |

`MaxPerGame` defaults to **0**, meaning no extra balls will ever be awarded. Set it explicitly.

It accepts a [placeholder](../concepts/placeholders.md), so the cap can vary with difficulty
settings.

## Events posted

| Event | When |
|---|---|
| `extra_ball_<name>_awarded` | This device awarded one |
| `extra_ball_awarded` | Any extra ball was awarded |

Hook the generic one for the callout and display, so it works no matter which device awarded:

```vbscript
With .EventPlayer()
    .Add "extra_ball_awarded", Array("play_voc_extra_ball", "show_extra_ball_slide")
End With
```

## How the award is consumed

Awarding increments the player variable `extra_balls`. At end of ball, if that count is above zero,
GLF decrements it and gives the player another ball instead of advancing to the next player. You
don't need to handle that yourself.

The device also tracks its own awards in `extra_ball_<name>_awarded`, which is what `MaxPerGame`
tests against. That's per device — two devices each capped at 1 can award two extra balls between
them.

## Conditional awards

`AwardEvents` accepts conditions, so gating is done in the event string:

```vbscript
With .ExtraBalls("main")
    .AwardEvents = Array("mystery_award{current_player.extra_balls == 0}")
    .MaxPerGame = 2
End With
```

## Lighting the extra ball

Typically a shot lit when the extra ball qualifies:

```vbscript
With .Shots("extra_ball_light")
    .Profile = "off_on_color"
    .Switch = "s_Scoop"
    With .Tokens()
        .Add "lights", "L05"
        .Add "color", ExtraBallColor
    End With
    .EnableEvents = Array("extra_ball_lit")
End With

With .EventPlayer()
    .Add "extra_ball_light_hit", Array("extra_ball_lit_collected")
End With
```

## Notes

- Once the cap is reached, award events are ignored silently.
- Extra balls carry across balls but not across games.
- Standard practice is a "shoot again" indication at the start of the extra ball; hook
  `ball_started` with a condition on `current_player.extra_balls`.

## See also

[Variables](../concepts/variables.md) · [Shot](shot.md)
