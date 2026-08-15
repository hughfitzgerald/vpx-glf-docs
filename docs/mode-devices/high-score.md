# High score

Compares end-of-game scores against stored high scores, collects initials, and persists the table.

```vbscript
With EnableGlfHighScores()
    With .Categories()
        .Add "score", Array("GRAND CHAMPION", "HIGH SCORE 1", "HIGH SCORE 2", "HIGH SCORE 3")
    End With
    With .Defaults("score")
        .Add "DAN", 5000000
        .Add "MPC", 2000000
        .Add "LFS", 1000000
        .Add "PUP",  500000
    End With
End With
```

Call `EnableGlfHighScores()` from `ConfigureGlfDevices`. It creates its own internal mode, started
on `game_ending`, which holds the game-end queue open until high score entry finishes.

## Categories and defaults

`Categories` maps a player variable to the award labels for each position. The array length sets
how many places are tracked — four labels means a four-deep table.

`Defaults` seeds the table the first time the machine runs. After that, values come from
`<cGameName>_glf.ini`; delete it to reset.

Categories aren't limited to `score`. Any player variable can be ranked:

```vbscript
With .Categories()
    .Add "score", Array("GRAND CHAMPION", "HIGH SCORE 1", "HIGH SCORE 2")
    .Add "ramps_made", Array("RAMP CHAMPION")
End With
With .Defaults("ramps_made")
    .Add "ABC", 25
End With
```

## Other properties

| Property | Default | Meaning |
|---|---|---|
| `AwardSlideDisplayTime` | 4000 | Ms an award is shown |
| `EnterInitialsTimeout` | 20000 | Ms before initials entry gives up |
| `ResetHighScoreEvents` | — | Events that reset the table to defaults |

## Events

| Event | When |
|---|---|
| `high_score_enter_initials` | Initials are needed |
| `high_score_award_display` | An award is being shown |
| `<category>_award_display` | Award for a specific category |
| `<award>_award_display` | Award for a specific label, e.g. `GRAND CHAMPION_award_display` |
| `high_score_complete` | Entry finished; the mode stops |

## Initials entry

The built-in mode drives entry over BCP, mapping the flippers and start button to left / right /
select. If you have a media controller with a `high_score` slide, that works out of the box.

Without BCP — segment displays only — you build entry yourself in a table-side mode and tell GLF
the result by posting `text_input_high_score_complete` with a `text` kwarg:

```vbscript
With CreateGlfMode("high_score", 120)
    .StartEvents = Array("game_will_end")
    .StopEvents  = Array("high_score_complete")

    With .EventPlayer()
        ' each press commits the currently selected character
        .Add "s_start_active{current_player.hs_input_ready == 1}", Array("text_inputted")

        ' when three characters are in, hand the string back to GLF
        .Add "text_inputted.1{machine.high_score_initials_chars == 3}", _
             Array("text_input_high_score_complete:{text: machine.high_score_initials}")

        ' and give up if they walk away
        .Add "timer_high_score_timeout_complete", _
             Array("text_input_high_score_complete:{text: machine.high_score_initials}")
    End With
End With
```

The pieces: machine variables hold the in-progress string and the character index, flipper events
move the selection, and the `.1` priority offset makes the completion check run before the handler
that would append a fourth character.

## Notes

- The high score mode uses `UseWaitQueue`, so the game won't finish until entry completes or times
  out.
- Always set `EnterInitialsTimeout` to something sane — an abandoned game otherwise hangs at
  entry until the timeout.
- Scores are written to `<cGameName>_glf.ini` at table exit, not immediately.
- Tables also get `player1_score` … `player4_score` as persisted machine variables, updated at
  game end, for attract displays.

## See also

[Variables](../concepts/variables.md) ·
[Slide & widget player](../players/slide-and-widget-player.md)
