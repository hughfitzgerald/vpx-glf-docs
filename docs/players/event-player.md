# Event player

Posts events in response to events. The most-used piece of GLF — most of a ruleset is event player
entries connecting hardware events to named game events.

```vbscript
With .EventPlayer()
    .Add "s_Bumper1_active", Array("score_5000", "play_bumper1_show", "backglass_flash1")
End With
```

`.Add key, values` — `key` is the event to listen for, `values` is an array of events to post.

## Conditions and priority

The key accepts the full [event syntax](../concepts/events.md#event-string-syntax):

```vbscript
' only when a player variable matches
.Add "s_Plunger1_inactive{current_player.ball_just_started == 1}", Array("new_ball_active")

' only during a mode
.Add "s_Ramp_active{modes.multiball.active == True}", Array("ramp_jackpot")

' compound
.Add "s_Scoop_active{current_player.locks_lit == 1 && machine.bottom_ball_locked == 0}", Array("lock_ball")

' priority offset — runs before other handlers of the same event in this mode
.Add "text_inputted.1{machine.high_score_initials_chars == 3}", Array("initials_complete")
```

Each `.Add` key must be unique within the player. Two rules on the same event need distinguishing
conditions or priority suffixes — which is usually what you wanted anyway:

```vbscript
.Add "s_Ramp_active{current_player.ramp_level == 1}", Array("ramp_award_small")
.Add "s_Ramp_active{current_player.ramp_level == 2}", Array("ramp_award_big")
```

## Posting with kwargs

Attach data to a posted event with `event:{key: value}`:

```vbscript
.Add "s_left_flipper_active", Array("text_input: {action: ""left""}")
.Add "timer_hs_timeout_complete", Array("text_input_high_score_complete:{text: machine.high_score_initials}")
```

Values may be placeholders. Listeners read them back with `kwargs.<key>`.

## A worked example

Mode start and stop wiring, which nearly every mode needs:

```vbscript
With .EventPlayer()
    .Add "mode_base_started",  Array("new_ball_started", "stop_attract_mode", "play_mus_ambient_loop")
    .Add "mode_base_stopping", Array("stop_mus_ambient_loop", "backglass_off")

    ' per-player display setup
    .Add "mode_base_started{current_player.number == 1}", Array("flash_player1_score")
    .Add "mode_base_started{current_player.number == 2}", Array("flash_player2_score")

    ' playfield
    .Add "s_Bumper1_active",      Array("play_bumper1_show", "score_5000")
    .Add "s_LeftSlingshot_active", Array("backglass_flash3", "score_5000")
End With
```

## Notes

- Fan-out is free: one event can post any number, and several rules can listen to the same event.
- Posted events are queued, not immediate — see [the event system](../concepts/events.md).
- If you need the *next* handler to wait for something, that's the
  [queue relay player](queue-relay-player.md), not this one.

## See also

[Queue event player](queue-event-player.md) ·
[Random event player](random-event-player.md) ·
[Variable player](variable-player.md)
